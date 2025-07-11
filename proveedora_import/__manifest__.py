# -*- coding: utf-8 -*-
{
    'name': 'Proveedora Import',
    'version': '18.0.1.0.0',
    'category': 'Tools',
    'summary': 'Importador de productos y tarifas desde hoja de cálculo',
    'author': 'Xtendoo Software SLU',
    'website': '',
    'license': 'LGPL-3',
    'depends': [
        'product',
        'stock',
        'sale',
        'base',
        'account_banking_mandate',
    ],
    'data': [
        'security/ir.model.access.csv',
        'wizard/import_products_wizard_view.xml',
        'wizard/import_customers_wizard_view.xml',
        'views/menu.xml',
    ],
    'installable': True,
    'application': False,
}
