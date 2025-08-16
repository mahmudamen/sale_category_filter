{
    'name': 'Sale Category Filter',
    'version': '18.0',
    'category': 'Sales/Sales',
    'summary': 'Filter products by category in sale orders with custom report wizard',
    'description': """
        This module adds a category field to sale orders that filters 
        products by category in sale order lines and includes a custom report wizard
        that can generate reports in PDF or DOCX format.
    """,
    'depends': ['base','mail','sale','sale_management'],
    'external_dependencies': {
        'python': ['python-docx'],
    },
    'data': [
        'security/ir.model.access.csv',
        'views/sale_order_views.xml',
        'wizard/wizard_views.xml',
        'report/sale_order_report.xml',
        'report/sale_order_template.xml',
    ],
    'installable': True,
    'application': False,
    'auto_install': False,
    'license': 'LGPL-3',
}
