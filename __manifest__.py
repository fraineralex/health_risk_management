# __manifest__.py
{
    'name': 'ARS Template Export',
    'version': '1.0',
    'category': 'Accounting',
    'summary': 'ARS Template Export Module',
    'description': "ARS Template Export Module",
    'depends': ['account'],
    'data': [
        # views
        'views/ir.ui.menu.xml',

        # wizard views
        'wizard/ars_export_wizard_view.xml',
        
        # security
        'security/ir.model.access.csv',
    ],
    'installable': True,
    'application': True,
    'auto_install': True,
    'license': 'LGPL-3',
}
