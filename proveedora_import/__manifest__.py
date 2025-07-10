# -*- coding: utf-8 -*-
{
    'name': 'Proveedora Import',
    'version': '18.0.1.0.0',
    'category': 'Tools',
    'summary': 'Importador de productos y tarifas desde hoja de cálculo',
    'author': 'Tu Empresa',
    'website': '',
    'depends': [
        'product',
        'stock',
        'sale',
        'base',
    ],
    'data': [
        'wizard/import_products_wizard_view.xml',
        'views/menu.xml',
    ],
    'installable': True,
    'application': False,
}

