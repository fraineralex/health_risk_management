{
    'name': 'Health Risk Management Reporter',
    'version': '1.0.0',
    'category': 'Accounting',
    'summary': 'Odoo Module designed to manage Health Risk Management (HRM) reports.',
    'description': "Odoo Module designed to manage Health Risk Management (HRM) reports.",
    'depends': ['account'],
    'author': 'Frainer Encarnación',
    'website': 'https://fraineralex.dev',
    'data': [
        # views
        'views/hrm_report_view.xml',
        'views/hrm_report_menu.xml',

        # security
        'security/ir.model.access.csv',
    ],
    'installable': True,
    'application': True,
    'auto_install': True,
    'license': 'LGPL-3',
}
