# __manifest__.py
{
    'name': 'ARS Template Export',
    'version': '1.0',
    'category': 'Accounting',
    'summary': 'ARS Template Export Module',
    'description': "ARS Template Export Module",
    'data': [
        # views
        'views/ir.ui.menu.xml',

        # wizard views
        'wizard/ars_export_wizard_view.xml',
    ],
    'icon': 'ars_template_export/static/src/img/icon.png',
    'installable': True,
    'application': True,
    'auto_install': True,
    'license': 'LGPL-3',
}
