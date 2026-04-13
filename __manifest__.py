{
    'name': 'Mouvements de Stock (Format Sage)',
    'version': '18.0.2.1.0',
    'category': 'Inventory/Reporting',
    'summary': 'Rapport mouvements de stock valorisés avec CMUP, format Sage 100',
    'description': """
Rapport de mouvements de stock au format Sage 100
==================================================
- Fiche de stock valorisée par produit
- Calcul du CMUP (Coût Moyen Unitaire Pondéré) via stock.valuation.layer
- Stock initial (report), détail mouvements, sous-totaux produit, total dépôt
- Export PDF et Excel
- Filtres : période, produit, dépôt, lot
    """,
    'author': 'SOPROMER',
    'depends': ['stock', 'stock_account'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/stock_movement_report_wizard_views.xml',
        'report/stock_movement_report_template.xml',
    ],
    'external_dependencies': {
        'python': ['xlsxwriter'],
    },
    'installable': True,
    'application': False,
    'license': 'LGPL-3',
}
