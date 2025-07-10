from odoo import models, fields, api
from odoo.exceptions import UserError
import base64
import io
import xlrd

class ImportProductsWizard(models.TransientModel):
    _name = 'proveedora.import.products.wizard'
    _description = 'Wizard para importar productos y tarifas desde Excel'

    file = fields.Binary('Archivo Excel', required=True)
    filename = fields.Char('Nombre del archivo')

    def action_import(self):
        if not self.file:
            raise UserError('Debe adjuntar un archivo.')
        data = base64.b64decode(self.file)
        workbook = xlrd.open_workbook(file_contents=data)
        sheet = workbook.sheet_by_index(0)  # Primera hoja

        # Obtener encabezados de la primera fila
        headers = [sheet.cell_value(0, col) for col in range(sheet.ncols)]

        # Crear tarifas si no existen
        pricelist_obj = self.env['product.pricelist']
        pricelists = {}
        for i in range(1, 11):
            name = f'PVP{i}'
            pricelist = pricelist_obj.search([('name', '=', name)], limit=1)
            if not pricelist:
                pricelist = pricelist_obj.create({'name': name, 'currency_id': self.env.company.currency_id.id})
            pricelists[name] = pricelist

        # Procesar productos
        product_obj = self.env['product.product']
        categ_obj = self.env['product.category']
        tax_obj = self.env['account.tax']
        iva_21_venta = tax_obj.search([
            ('description', '=', 'IVA 21% (Bienes)'),
            ('type_tax_use', '=', 'sale')
        ], limit=1)
        iva_21_compra = tax_obj.search([
            ('description', '=', 'IVA 21% (Bienes)'),
            ('type_tax_use', '=', 'purchase')
        ], limit=1)

        # Procesar cada fila (empezando desde la fila 1, saltando encabezados)
        for row_idx in range(1, sheet.nrows):
            # Crear diccionario de la fila actual
            row = {}
            for col_idx, header in enumerate(headers):
                if col_idx < sheet.ncols:
                    row[header] = sheet.cell_value(row_idx, col_idx)

            if str(row.get('ARTICULO_OBSOLETO', '')).strip().lower() == 'si':
                continue
            # Solo crear o actualizar producto si tiene CODIGO y DESCRIPCION
            codigo = str(row.get('CODIGO', '')).strip()
            descripcion = str(row.get('DESCRIPCION', '')).strip()
            if not codigo or not descripcion:
                continue
            categ = categ_obj.search([('name', '=', row.get('NIVEL1', 'Sin categoría'))], limit=1)
            if not categ:
                categ = categ_obj.create({'name': row.get('NIVEL1', 'Sin categoría')})
            # Precio de venta por defecto: PVP1
            precio_venta = row.get('PVP1', 0)
            # Precio de coste: COSTE_NETO
            precio_coste = row.get('COSTE_NETO', 0)
            vals = {
                'name': descripcion,
                'default_code': codigo,
                'categ_id': categ.id,
                'active': not bool(row.get('ARTICULO_BLOQUEADO', False)),
                'lst_price': precio_venta,
                'standard_price': precio_coste,
                'taxes_id': [(6, 0, [iva_21_venta.id])] if iva_21_venta else False,
                'supplier_taxes_id': [(6, 0, [iva_21_compra.id])] if iva_21_compra else False,
                'invoice_policy': 'delivery',
                'is_storable': True,
            }
            # Buscar producto por CODIGO (default_code)
            product = product_obj.search([('default_code', '=', codigo)], limit=1)
            if product:
                product.write(vals)
            else:
                product = product_obj.create(vals)
            # Stock inicial
            if 'L' in headers:
                qty = row.get('L', 0)
                if qty:
                    self.env['stock.quant'].with_context(inventory_mode=True).create({
                        'product_id': product.id,
                        'location_id': self.env.ref('stock.stock_location_stock').id,
                        'quantity': qty,
                    })
            # Tarifas y descuentos
            for i in range(1, 11):
                pvp_col = f'PVP{i}'
                desc_col = f'DESCUENTO{i}'
                if pvp_col in headers:
                    price = row.get(pvp_col, 0)
                    if price:
                        self.env['product.pricelist.item'].create({
                            'pricelist_id': pricelists[pvp_col].id,
                            'applied_on': '0_product_variant',
                            'product_id': product.id,
                            'fixed_price': price,
                            'min_quantity': 1,
                        })
                if desc_col in headers:
                    discount = row.get(desc_col, 0)
                    if discount:
                        self.env['product.pricelist.item'].create({
                            'pricelist_id': pricelists[pvp_col].id,
                            'applied_on': '0_product_variant',
                            'product_id': product.id,
                            'percent_price': discount,
                            'min_quantity': 1,
                        })
        return {'type': 'ir.actions.act_window_close'}
