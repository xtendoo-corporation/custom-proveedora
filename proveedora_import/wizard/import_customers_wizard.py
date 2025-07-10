from odoo import models, fields, api
from odoo.exceptions import UserError
import base64
import io
import xlrd

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

class ImportCustomersWizard(models.TransientModel):
    _name = 'proveedora.import.customers.wizard'
    _description = 'Wizard para importar clientes desde Excel'

    file = fields.Binary('Archivo Excel', required=True)
    filename = fields.Char('Nombre del archivo')

    def _read_excel_file(self, data):
        """Leer archivo Excel usando la librería apropiada según el formato"""

        # Verificar el formato del archivo
        header = data[:8]
        is_xlsx = header.startswith(b'PK')
        is_xls = header.startswith(b'\xd0\xcf\x11\xe0') or header.startswith(b'\x09\x08')

        if is_xlsx:
            # Archivo .xlsx - usar openpyxl
            if not OPENPYXL_AVAILABLE:
                raise UserError('El archivo es .xlsx pero openpyxl no está instalado. '
                              'Por favor, convierte el archivo a .xls o instala openpyxl.')

            try:
                workbook = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
                sheet = workbook.active

                # Obtener encabezados de la primera fila
                headers = []
                for cell in sheet[1]:
                    headers.append(cell.value if cell.value else '')

                # Procesar filas
                rows_data = []
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    row_dict = {}
                    for col_idx, value in enumerate(row):
                        if col_idx < len(headers):
                            row_dict[headers[col_idx]] = value
                    rows_data.append(row_dict)

                return headers, rows_data

            except Exception as e:
                raise UserError(f'Error al leer archivo .xlsx: {str(e)}')

        elif is_xls:
            # Archivo .xls - usar xlrd
            try:
                workbook = xlrd.open_workbook(file_contents=data)
                sheet = workbook.sheet_by_index(0)

                # Obtener encabezados de la primera fila
                headers = [sheet.cell_value(0, col) for col in range(sheet.ncols)]

                # Procesar filas
                rows_data = []
                for row_idx in range(1, sheet.nrows):
                    row_dict = {}
                    for col_idx, header in enumerate(headers):
                        if col_idx < sheet.ncols:
                            row_dict[header] = sheet.cell_value(row_idx, col_idx)
                    rows_data.append(row_dict)

                return headers, rows_data

            except Exception as e:
                if "BIFF2 cell record" in str(e) or "XLRDError" in str(e):
                    raise UserError('⚠️ Error de formato de Excel\n\n'
                                  'El archivo .xls tiene un formato que no se puede leer.\n'
                                  'Solución:\n'
                                  '1. Abre el archivo en Excel\n'
                                  '2. Guárdalo como "Excel 97-2003 Workbook (*.xls)"\n'
                                  '3. Vuelve a intentar la importación')
                else:
                    raise UserError(f'Error al leer archivo .xls: {str(e)}')
        else:
            raise UserError('El archivo no parece ser un Excel válido (.xls o .xlsx). '
                          'Por favor, sube un archivo Excel válido.')

    def action_import(self):
        if not self.file:
            raise UserError('Debe adjuntar un archivo.')

        try:
            data = base64.b64decode(self.file)
            if len(data) < 512:
                raise UserError('El archivo parece estar vacío o corrupto.')

            headers, rows_data = self._read_excel_file(data)

        except UserError:
            raise
        except Exception as e:
            raise UserError(f'Error inesperado al procesar el archivo: {str(e)}')

        # Procesar clientes
        partner_obj = self.env['res.partner']
        country_obj = self.env['res.country']
        state_obj = self.env['res.country.state']

        # Contadores para el mensaje final
        clientes_creados = 0
        clientes_actualizados = 0

        # Procesar cada fila
        for row in rows_data:
            # Solo crear o actualizar cliente si tiene CODIGO y NOMBRE
            codigo_raw = row.get('CODIGO', '')
            nombre_raw = row.get('NOMBRE', '')

            # Convertir a string y limpiar
            if isinstance(codigo_raw, float):
                if codigo_raw == int(codigo_raw):
                    codigo = str(int(codigo_raw)).strip()
                else:
                    codigo = str(codigo_raw).strip()
            else:
                codigo = str(codigo_raw).strip() if codigo_raw else ''

            if isinstance(nombre_raw, float):
                nombre = str(nombre_raw).strip()
            else:
                nombre = str(nombre_raw).strip() if nombre_raw else ''

            # Validar que no estén vacíos
            if not codigo or not nombre or codigo == 'nan' or nombre == 'nan' or codigo == 'None' or nombre == 'None':
                continue

            # Buscar cliente existente por CODIGO (ref)
            partner = partner_obj.search([('ref', '=', codigo)], limit=1)

            # Buscar o crear país
            pais = None
            if row.get('PAIS'):
                pais = country_obj.search([('name', 'ilike', str(row.get('PAIS')).strip())], limit=1)
                if not pais:
                    pais = country_obj.search([('code', 'ilike', str(row.get('PAIS')).strip())], limit=1)

            # Buscar o crear provincia/estado
            provincia = None
            if row.get('PROVINCIA') and pais:
                provincia = state_obj.search([
                    ('name', 'ilike', str(row.get('PROVINCIA')).strip()),
                    ('country_id', '=', pais.id)
                ], limit=1)

            # Preparar valores del cliente
            vals = {
                'name': nombre,
                'ref': codigo,
                'is_company': True,  # Por defecto como empresa
                'customer_rank': 1,  # Marcar como cliente
                'supplier_rank': 0,  # No es proveedor
                'vat': str(row.get('CIF', '')).strip() if row.get('CIF') else False,
                'street': str(row.get('DIRECCION', '')).strip() if row.get('DIRECCION') else False,
                'city': str(row.get('POBLACION', '')).strip() if row.get('POBLACION') else False,
                'zip': str(row.get('C_POSTAL', '')).strip() if row.get('C_POSTAL') else False,
                'country_id': pais.id if pais else False,
                'state_id': provincia.id if provincia else False,
                'phone': str(row.get('TELEFONO1', '')).strip() if row.get('TELEFONO1') else False,
                'mobile': str(row.get('MOVIL', '')).strip() if row.get('MOVIL') else False,
                'email': str(row.get('EMAIL', '')).strip() if row.get('EMAIL') else False,
                'website': str(row.get('WWW', '')).strip() if row.get('WWW') else False,
                'active': str(row.get('ACTIVO', '')).strip().lower() != 'no',
            }

            # Campos adicionales específicos
            if row.get('NOMBRE_COMERCIAL'):
                vals['commercial_company_name'] = str(row.get('NOMBRE_COMERCIAL')).strip()

            if row.get('TELEFONO2'):
                vals['fax'] = str(row.get('TELEFONO2')).strip()

            if row.get('PERSONA_DE_CONTACTO'):
                vals['comment'] = f"Persona de contacto: {str(row.get('PERSONA_DE_CONTACTO')).strip()}"

            # Crear o actualizar cliente
            if partner:
                partner.write(vals)
                clientes_actualizados += 1
            else:
                partner = partner_obj.create(vals)
                clientes_creados += 1

            # Crear dirección comercial si es diferente
            direccion_comercial = str(row.get('DIRECCION_COMERCIAL', '')).strip()
            if direccion_comercial and direccion_comercial != vals.get('street', ''):
                # Buscar país comercial
                pais_comercial = None
                if row.get('PAIS_COMERCIAL'):
                    pais_comercial = country_obj.search([('name', 'ilike', str(row.get('PAIS_COMERCIAL')).strip())], limit=1)
                    if not pais_comercial:
                        pais_comercial = country_obj.search([('code', 'ilike', str(row.get('PAIS_COMERCIAL')).strip())], limit=1)

                # Buscar provincia comercial
                provincia_comercial = None
                if row.get('PROVINCIA_COMERCIAL') and pais_comercial:
                    provincia_comercial = state_obj.search([
                        ('name', 'ilike', str(row.get('PROVINCIA_COMERCIAL')).strip()),
                        ('country_id', '=', pais_comercial.id)
                    ], limit=1)

                # Crear dirección comercial como contacto hijo
                direccion_vals = {
                    'name': f"Dirección comercial - {nombre}",
                    'parent_id': partner.id,
                    'type': 'delivery',
                    'street': direccion_comercial,
                    'city': str(row.get('POBLACION_COMERCIAL', '')).strip() if row.get('POBLACION_COMERCIAL') else False,
                    'zip': str(row.get('C_POSTAL_COMERCIAL', '')).strip() if row.get('C_POSTAL_COMERCIAL') else False,
                    'country_id': pais_comercial.id if pais_comercial else False,
                    'state_id': provincia_comercial.id if provincia_comercial else False,
                    'phone': str(row.get('TELEFONO1_COMERCIAL', '')).strip() if row.get('TELEFONO1_COMERCIAL') else False,
                    'mobile': str(row.get('MOVIL_COMERCIAL', '')).strip() if row.get('MOVIL_COMERCIAL') else False,
                }

                if row.get('PERSONA_DE_CONTACTO_COMERCIAL'):
                    direccion_vals['comment'] = f"Persona de contacto: {str(row.get('PERSONA_DE_CONTACTO_COMERCIAL')).strip()}"

                partner_obj.create(direccion_vals)

        # Mensaje de confirmación al finalizar
        message = f"Importación de clientes completada:\n• {clientes_creados} clientes creados\n• {clientes_actualizados} clientes actualizados"
        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': 'Importación de clientes finalizada',
                'message': message,
                'type': 'success',
                'sticky': False,
            }
        }
