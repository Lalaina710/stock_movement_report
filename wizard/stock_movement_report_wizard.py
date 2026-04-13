import io
import base64
from itertools import groupby

from datetime import datetime, time

import pytz

from odoo import api, fields, models, _
from odoo.exceptions import UserError


class StockMovementReportWizard(models.TransientModel):
    _name = 'stock.movement.report.wizard'
    _description = 'Wizard Rapport Mouvements de Stock'

    date_from = fields.Date(
        string='Date début', required=True,
        default=lambda self: fields.Date.today().replace(day=1),
    )
    date_to = fields.Date(
        string='Date fin', required=True,
        default=fields.Date.today,
    )
    product_ids = fields.Many2many(
        'product.product', string='Produits',
        domain=[('is_storable', '=', True)],
        help='Laisser vide pour tous les produits stockables',
    )
    warehouse_ids = fields.Many2many(
        'stock.warehouse', string='Dépôt(s)',
        help='Laisser vide pour tous les dépôts',
    )
    lot_id = fields.Many2one(
        'stock.lot', string='Lot/N° de série',
    )
    # Excel download
    report_file = fields.Binary('Fichier', readonly=True)
    report_filename = fields.Char('Nom du fichier', readonly=True)

    # -------------------------------------------------------------------------
    # Actions
    # -------------------------------------------------------------------------

    def action_print_pdf(self):
        self.ensure_one()
        self._validate()
        return self.env.ref(
            'stock_movement_report.action_report_stock_movement'
        ).report_action(self)

    def action_export_excel(self):
        self.ensure_one()
        self._validate()
        data = self._get_report_data()
        content = self._generate_xlsx(data)
        self.report_file = base64.b64encode(content)
        self.report_filename = 'mouvements_stock_%s_%s.xlsx' % (
            self.date_from.strftime('%Y%m%d'),
            self.date_to.strftime('%Y%m%d'),
        )
        return {
            'type': 'ir.actions.act_url',
            'url': '/web/content/?model=%s&id=%d&field=report_file'
                   '&filename_field=report_filename&download=true' % (
                       self._name, self.id),
            'target': 'new',
        }

    def _validate(self):
        if self.date_from > self.date_to:
            raise UserError(_('La date de début doit être antérieure à la date de fin.'))

    # -------------------------------------------------------------------------
    # Report data computation
    # -------------------------------------------------------------------------

    def _get_report_data(self):
        """Build the full report data structure grouped by warehouse then product."""
        self.ensure_one()

        tz = pytz.timezone(self.env.user.tz or 'UTC')
        date_from_dt = tz.localize(datetime.combine(
            self.date_from, time.min,
        )).astimezone(pytz.utc).replace(tzinfo=None)
        date_to_dt = tz.localize(datetime.combine(
            self.date_to, time.max,
        )).astimezone(pytz.utc).replace(tzinfo=None)

        # Determine warehouses to iterate
        warehouses = self.warehouse_ids or self.env['stock.warehouse'].search([
            ('company_id', '=', self.env.company.id),
        ])

        # Pre-compute CMUP for ALL products at date_from (1 single SQL query)
        cmup_cache = self._batch_compute_cmup_at_date(date_from_dt)

        warehouses_data = []
        grand_total_value = 0.0

        for wh in warehouses:
            wh_data = self._compute_warehouse_data(
                wh, date_from_dt, date_to_dt, cmup_cache)
            if wh_data['products']:
                warehouses_data.append(wh_data)
                grand_total_value += wh_data['warehouse_total_value']

        return {
            'company': self.env.company,
            'date_from': self.date_from.strftime('%d/%m/%Y'),
            'date_to': self.date_to.strftime('%d/%m/%Y'),
            'warehouses': warehouses_data,
            'grand_total_value': grand_total_value,
            'print_date': fields.Datetime.context_timestamp(
                self, fields.Datetime.now()
            ).strftime('%d/%m/%Y à %H:%M:%S'),
        }

    def _compute_warehouse_data(self, warehouse, date_from_dt, date_to_dt,
                                cmup_cache):
        """Compute report data for a single warehouse."""
        location_ids = self._get_warehouse_location_ids(warehouse)
        if not location_ids:
            return {'warehouse_name': warehouse.name, 'products': [],
                    'warehouse_total_value': 0.0}

        # Fetch moves in period touching warehouse locations
        domain = [
            ('state', '=', 'done'),
            ('date', '>=', date_from_dt),
            ('date', '<=', date_to_dt),
            '|',
            ('location_id', 'in', location_ids),
            ('location_dest_id', 'in', location_ids),
        ]
        if self.product_ids:
            domain.append(('product_id', 'in', self.product_ids.ids))
        if self.lot_id:
            domain.append(('lot_ids', 'in', [self.lot_id.id]))

        moves = self.env['stock.move'].search(domain, order='product_id, date, id')

        # Filter out moves internal to the same set of locations (net zero)
        location_set = set(location_ids)
        moves = moves.filtered(
            lambda m: not (m.location_id.id in location_set
                          and m.location_dest_id.id in location_set)
        )

        if not moves:
            return {'warehouse_name': warehouse.name, 'products': [],
                    'warehouse_total_value': 0.0}

        # Collect distinct product IDs from filtered moves
        product_ids = list(set(moves.mapped('product_id').ids))

        # Batch opening qty: 1 SQL for all products of this warehouse
        opening_qty_map = self._batch_compute_opening_qty(
            product_ids, location_ids, date_from_dt)

        # Prefetch all valuation layers for these moves in one query
        all_layers = self.env['stock.valuation.layer'].search([
            ('stock_move_id', 'in', moves.ids),
        ])
        layers_by_move = {}
        for layer in all_layers:
            layers_by_move.setdefault(layer.stock_move_id.id, []).append(layer)

        # Prefetch standard_price for fallback
        products_browse = self.env['product.product'].browse(product_ids)
        std_price_map = {p.id: p.standard_price for p in products_browse}

        # Group by product
        products_data = []
        warehouse_total_value = 0.0

        for _pid, grp in groupby(moves, key=lambda m: m.product_id.id):
            product_moves = self.env['stock.move'].concat(*list(grp))
            product = product_moves[0].product_id

            # Opening balance from batch caches
            opening_qty = opening_qty_map.get(product.id, 0.0)
            opening_cmup = cmup_cache.get(product.id,
                                          std_price_map.get(product.id, 0.0))
            opening_value = opening_qty * opening_cmup

            lines = []
            running_qty = opening_qty
            running_value = opening_value
            current_cmup = opening_cmup

            for move in product_moves:
                qty = self._compute_move_qty(move, location_set)
                move_type = self._classify_move(move, location_set)

                # Unit cost from prefetched valuation layers
                move_layers = layers_by_move.get(move.id, [])
                if move_layers:
                    layer_value = sum(l.value for l in move_layers)
                    layer_qty = sum(l.quantity for l in move_layers)
                    move_unit_cost = abs(layer_value / layer_qty) if layer_qty else current_cmup
                else:
                    move_unit_cost = current_cmup

                # Update CMUP on entry (AVCO rule)
                if qty > 0 and (running_qty + qty) > 0:
                    current_cmup = (
                        (running_qty * current_cmup + qty * move_unit_cost)
                        / (running_qty + qty)
                    )

                running_qty += qty
                running_value = running_qty * current_cmup

                lines.append({
                    'date': fields.Date.to_string(move.date),
                    'date_fmt': move.date.strftime('%d/%m/%Y'),
                    'type': move_type,
                    'reference': move.picking_id.name or move.reference or move.name or '',
                    'partner': move.picking_id.partner_id.name or '',
                    'qty': qty,
                    'balance': running_qty,
                    'unit_cost': current_cmup,
                    'stock_value': running_value,
                })

            closing_value = running_qty * current_cmup
            warehouse_total_value += closing_value

            products_data.append({
                'product': product,
                'default_code': product.default_code or '',
                'name': product.name,
                'opening_qty': opening_qty,
                'opening_value': opening_value,
                'opening_cmup': opening_cmup,
                'lines': lines,
                'closing_qty': running_qty,
                'closing_value': closing_value,
                'closing_cmup': current_cmup,
                'total_in': sum(l['qty'] for l in lines if l['qty'] > 0),
                'total_out': sum(l['qty'] for l in lines if l['qty'] < 0),
            })

        return {
            'warehouse_name': warehouse.name,
            'products': products_data,
            'warehouse_total_value': warehouse_total_value,
        }

    # -------------------------------------------------------------------------
    # Helpers
    # -------------------------------------------------------------------------

    def _get_warehouse_location_ids(self, warehouse):
        """Get internal location IDs for a specific warehouse."""
        locations = self.env['stock.location'].search([
            ('id', 'child_of', warehouse.lot_stock_id.id),
            ('usage', '=', 'internal'),
        ])
        return locations.ids

    def _batch_compute_opening_qty(self, product_ids, location_ids, date_from_dt):
        """Batch opening qty for multiple products in 1 SQL query."""
        if not product_ids:
            return {}
        self.env.cr.execute("""
            SELECT sm.product_id,
                COALESCE(SUM(CASE
                    WHEN sm.location_dest_id = ANY(%(locs)s)
                     AND sm.location_id != ALL(%(locs)s)
                    THEN sm.quantity ELSE 0 END), 0)
                -
                COALESCE(SUM(CASE
                    WHEN sm.location_id = ANY(%(locs)s)
                     AND sm.location_dest_id != ALL(%(locs)s)
                    THEN sm.quantity ELSE 0 END), 0)
            FROM stock_move sm
            WHERE sm.state = 'done'
              AND sm.product_id = ANY(%(pids)s)
              AND sm.company_id = %(company_id)s
              AND sm.date < %(dt)s
              AND (sm.location_id = ANY(%(locs)s)
                   OR sm.location_dest_id = ANY(%(locs)s))
            GROUP BY sm.product_id
        """, {
            'locs': location_ids,
            'pids': product_ids,
            'company_id': self.env.company.id,
            'dt': date_from_dt,
        })
        return dict(self.env.cr.fetchall())

    def _batch_compute_cmup_at_date(self, date_dt):
        """Batch CMUP for ALL products in 1 SQL query. Returns {product_id: cmup}."""
        self.env.cr.execute("""
            SELECT svl.product_id,
                   SUM(svl.value),
                   SUM(svl.quantity)
            FROM stock_valuation_layer svl
            LEFT JOIN stock_move sm ON sm.id = svl.stock_move_id
            WHERE svl.company_id = %s
              AND COALESCE(sm.date, svl.create_date) < %s
            GROUP BY svl.product_id
            HAVING SUM(svl.quantity) > 0
        """, (self.env.company.id, date_dt))
        result = {}
        for product_id, total_value, total_qty in self.env.cr.fetchall():
            result[product_id] = total_value / total_qty
        return result

    def _compute_move_qty(self, move, location_set):
        """Signed qty change for our locations. Positive = entry, negative = exit."""
        qty = move.quantity
        if self.lot_id:
            lot_lines = move.move_line_ids.filtered(
                lambda ml: ml.lot_id == self.lot_id
            )
            qty = sum(lot_lines.mapped('quantity'))
        dst_in = move.location_dest_id.id in location_set
        src_in = move.location_id.id in location_set
        if dst_in and not src_in:
            return qty
        elif src_in and not dst_in:
            return -qty
        return 0.0

    def _classify_move(self, move, location_set):
        """Classify a stock move into Sage-style type codes."""
        if move.is_inventory:
            return 'INV'
        if hasattr(move, 'production_id') and (move.production_id or move.raw_material_production_id):
            return 'FAB'
        if move.picking_id and move.picking_id.picking_type_id:
            code = move.picking_id.picking_type_id.code
            origin = (move.picking_id.origin or '').lower()
            is_return = 'return' in origin or 'retour' in origin
            if code == 'incoming':
                return 'RET' if is_return else 'REC'
            elif code == 'outgoing':
                return 'RET' if is_return else 'BL'
            elif code == 'internal':
                return 'INT'
        return 'AUT'

    # -------------------------------------------------------------------------
    # Excel generation
    # -------------------------------------------------------------------------

    def _generate_xlsx(self, data):
        import xlsxwriter

        output = io.BytesIO()
        wb = xlsxwriter.Workbook(output, {'in_memory': True})

        # Formats
        fmt_title = wb.add_format({
            'bold': True, 'font_size': 14, 'align': 'center',
        })
        fmt_header = wb.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'align': 'center', 'text_wrap': True,
        })
        fmt_text = wb.add_format({'border': 1, 'font_size': 10})
        fmt_num = wb.add_format({
            'border': 1, 'font_size': 10, 'num_format': '#,##0.00',
        })
        fmt_num_neg = wb.add_format({
            'border': 1, 'font_size': 10, 'num_format': '#,##0.00',
            'font_color': 'red',
        })
        fmt_product = wb.add_format({
            'bold': True, 'bg_color': '#D9E2F3', 'border': 1,
            'font_size': 11,
        })
        fmt_subtotal = wb.add_format({
            'bold': True, 'bg_color': '#E2EFDA', 'border': 1,
            'font_size': 10, 'num_format': '#,##0.00',
        })
        fmt_subtotal_text = wb.add_format({
            'bold': True, 'bg_color': '#E2EFDA', 'border': 1,
            'font_size': 10,
        })
        fmt_wh_total = wb.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'font_size': 11, 'num_format': '#,##0.00',
        })
        fmt_wh_total_text = wb.add_format({
            'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
            'border': 1, 'font_size': 11,
        })
        fmt_grand = wb.add_format({
            'bold': True, 'bg_color': '#1F3864', 'font_color': 'white',
            'border': 2, 'font_size': 12, 'num_format': '#,##0.00',
        })
        fmt_grand_text = wb.add_format({
            'bold': True, 'bg_color': '#1F3864', 'font_color': 'white',
            'border': 2, 'font_size': 12,
        })

        headers = [
            'Date mouv.', 'Type mouv.', 'N° de pièce',
            'Référence / Tiers', '+/-', 'Solde',
            'P.R. unitaire', 'Stock permanent',
        ]

        for wh_data in data['warehouses']:
            ws = wb.add_worksheet(wh_data['warehouse_name'][:31])

            # Column widths
            ws.set_column(0, 0, 12)
            ws.set_column(1, 1, 8)
            ws.set_column(2, 2, 18)
            ws.set_column(3, 3, 35)
            ws.set_column(4, 4, 12)
            ws.set_column(5, 5, 12)
            ws.set_column(6, 6, 15)
            ws.set_column(7, 7, 18)

            # Title
            ws.merge_range(0, 0, 0, 7, 'Mouvements de stock', fmt_title)
            ws.write(1, 0, data['company'].name, fmt_text)
            ws.write(1, 3, wh_data['warehouse_name'], fmt_text)
            ws.write(1, 6, 'Période du', fmt_text)
            ws.write(1, 7, '%s au %s' % (data['date_from'], data['date_to']), fmt_text)

            row = 3
            for col, h in enumerate(headers):
                ws.write(row, col, h, fmt_header)
            row += 1

            for pdata in wh_data['products']:
                # Product header
                label = pdata['default_code'] or ''
                if label:
                    label += '  '
                label += pdata['name']
                ws.merge_range(row, 0, row, 3, label, fmt_product)
                ws.write(row, 4, '', fmt_product)
                ws.write(row, 5, '', fmt_product)
                ws.write(row, 6, '', fmt_product)
                ws.write(row, 7, '', fmt_product)
                row += 1

                # Opening balance
                ws.write(row, 0, data['date_from'], fmt_text)
                ws.write(row, 1, 'Report', fmt_text)
                ws.write(row, 2, '', fmt_text)
                ws.write(row, 3, 'Stock', fmt_text)
                ws.write(row, 4, '', fmt_text)
                ws.write(row, 5, pdata['opening_qty'], fmt_num)
                ws.write(row, 6, pdata['opening_cmup'], fmt_num)
                ws.write(row, 7, pdata['opening_value'], fmt_num)
                row += 1

                # Move lines
                for line in pdata['lines']:
                    ws.write(row, 0, line['date_fmt'], fmt_text)
                    ws.write(row, 1, line['type'], fmt_text)
                    ws.write(row, 2, line['reference'], fmt_text)
                    ws.write(row, 3, line['partner'], fmt_text)
                    ws.write(row, 4, line['qty'],
                             fmt_num_neg if line['qty'] < 0 else fmt_num)
                    ws.write(row, 5, line['balance'], fmt_num)
                    ws.write(row, 6, line['unit_cost'], fmt_num)
                    ws.write(row, 7, line['stock_value'], fmt_num)
                    row += 1

                # Product subtotal
                code = pdata['default_code'] or pdata['name']
                ws.write(row, 0, '', fmt_subtotal_text)
                ws.write(row, 1, '', fmt_subtotal_text)
                ws.merge_range(row, 2, row, 3,
                               'Total  %s' % code, fmt_subtotal_text)
                ws.write(row, 4, '', fmt_subtotal_text)
                ws.write(row, 5, pdata['closing_qty'], fmt_subtotal)
                ws.write(row, 6, '', fmt_subtotal_text)
                ws.write(row, 7, pdata['closing_value'], fmt_subtotal)
                row += 1
                row += 1  # blank row

            # Warehouse total
            ws.write(row, 0, '', fmt_wh_total_text)
            ws.write(row, 1, '', fmt_wh_total_text)
            ws.merge_range(row, 2, row, 3,
                           'Total  %s' % wh_data['warehouse_name'], fmt_wh_total_text)
            ws.write(row, 4, '', fmt_wh_total_text)
            ws.write(row, 5, '', fmt_wh_total_text)
            ws.write(row, 6, '', fmt_wh_total_text)
            ws.write(row, 7, wh_data['warehouse_total_value'], fmt_wh_total)
            row += 1

            # A reporter
            ws.write(row, 0, '', fmt_wh_total_text)
            ws.write(row, 1, '', fmt_wh_total_text)
            ws.merge_range(row, 2, row, 5, 'A reporter', fmt_wh_total_text)
            ws.write(row, 6, '', fmt_wh_total_text)
            ws.write(row, 7, wh_data['warehouse_total_value'], fmt_wh_total)

        # Summary sheet if multiple warehouses
        if len(data['warehouses']) > 1:
            ws_sum = wb.add_worksheet('Récapitulatif')
            ws_sum.set_column(0, 0, 40)
            ws_sum.set_column(1, 1, 20)
            ws_sum.merge_range(0, 0, 0, 1, 'Récapitulatif par dépôt', fmt_title)
            ws_sum.write(1, 0, '%s au %s' % (data['date_from'], data['date_to']), fmt_text)
            ws_sum.write(3, 0, 'Dépôt', fmt_header)
            ws_sum.write(3, 1, 'Valeur stock', fmt_header)
            row = 4
            for wh_data in data['warehouses']:
                ws_sum.write(row, 0, wh_data['warehouse_name'], fmt_text)
                ws_sum.write(row, 1, wh_data['warehouse_total_value'], fmt_num)
                row += 1
            ws_sum.write(row, 0, 'TOTAL GENERAL', fmt_grand_text)
            ws_sum.write(row, 1, data['grand_total_value'], fmt_grand)

        wb.close()
        return output.getvalue()
