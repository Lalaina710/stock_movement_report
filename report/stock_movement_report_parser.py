from odoo import api, models


class StockMovementReportParser(models.AbstractModel):
    _name = 'report.stock_movement_report.report_stock_movement'
    _description = 'Parser rapport mouvements de stock'

    @api.model
    def _get_report_values(self, docids, data=None):
        wizards = self.env['stock.movement.report.wizard'].browse(docids)
        report_data = {}
        for wizard in wizards:
            report_data[wizard.id] = wizard._get_report_data()
        return {
            'doc_ids': docids,
            'docs': wizards,
            'data': report_data,
        }
