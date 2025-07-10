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
        product_obj = self.env['product.template']  # Cambiar a product.template
        categ_obj = self.env['product.category']
        tax_obj = self.env['account.tax']

        # Buscar impuestos de venta y compra por separado
        iva_21_venta = tax_obj.search([
            ('name', '=', '21% G'),
            ('type_tax_use', '=', 'sale'),
            ('amount', '=', 21)
        ], limit=1)

        iva_21_compra = tax_obj.search([
            ('name', '=', '21% G'),
            ('type_tax_use', '=', 'purchase'),
            ('amount', '=', 21)
        ], limit=1)

        # Contadores para el mensaje final
        productos_creados = 0
        productos_actualizados = 0

        # Información de depuración para impuestos
        debug_info = []
        if iva_21_venta:
            debug_info.append(f"IVA Venta encontrado: {iva_21_venta.name} (ID: {iva_21_venta.id})")
        else:
            debug_info.append("IVA Venta: NO ENCONTRADO")

        if iva_21_compra:
            debug_info.append(f"IVA Compra encontrado: {iva_21_compra.name} (ID: {iva_21_compra.id})")
        else:
            debug_info.append("IVA Compra: NO ENCONTRADO")

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
            codigo_raw = row.get('CODIGO', '')
            descripcion_raw = row.get('DESCRIPCION', '')

            # Convertir a string y limpiar, manejando diferentes tipos de datos
            if isinstance(codigo_raw, float):
                if codigo_raw == int(codigo_raw):  # Si es un entero disfrazado de float
                    codigo = str(int(codigo_raw)).strip()
                else:
                    codigo = str(codigo_raw).strip()
            else:
                codigo = str(codigo_raw).strip()

            if isinstance(descripcion_raw, float):
                descripcion = str(descripcion_raw).strip()
            else:
                descripcion = str(descripcion_raw).strip()

            # Validar que no estén vacíos
            if not codigo or not descripcion or codigo == 'nan' or descripcion == 'nan':
                continue

            # Buscar producto existente por CODIGO (default_code)
            product = product_obj.search([('default_code', '=', codigo)], limit=1)

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
                'list_price': precio_venta,
                'standard_price': precio_coste,
                'taxes_id': [(6, 0, [iva_21_venta.id])] if iva_21_venta else False,
                'supplier_taxes_id': [(6, 0, [iva_21_compra.id])] if iva_21_compra else False,
                'invoice_policy': 'delivery',
            }

            # Intentar establecer el tipo de producto como almacenable si es posible
            try:
                # Verificar si existe el campo y qué valores acepta
                if hasattr(product_obj, '_fields') and 'detailed_type' in product_obj._fields:
                    field_selection = product_obj._fields['detailed_type'].selection
                    if 'product' in [x[0] for x in field_selection]:
                        vals['detailed_type'] = 'product'
                    elif 'storable' in [x[0] for x in field_selection]:
                        vals['detailed_type'] = 'storable'
            except:
                # Si hay algún error, simplemente no establecer el tipo
                pass

            # Crear o actualizar producto
            if product:
                product.write(vals)
                productos_actualizados += 1
            else:
                product = product_obj.create(vals)
                productos_creados += 1
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
        # Mensaje de confirmación al finalizar
        debug_message = "\n".join(debug_info)
        message = f"Importación completada:\n• {productos_creados} productos creados\n• {productos_actualizados} productos actualizados\n\nInformación de depuración:\n{debug_message}"
        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': 'Importación finalizada',
                'message': message,
                'type': 'success',
                'sticky': False,
            }
        }
