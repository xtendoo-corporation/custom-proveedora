from odoo import models, fields, api
from odoo.exceptions import UserError
import base64
import io
import xlrd
import logging

_logger = logging.getLogger(__name__)

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

    def _safe_str(self, value):
        """Convierte valor a string de forma segura"""
        if value is None or value == '':
            return ''
        if isinstance(value, float):
            if value == int(value):
                return str(int(value)).strip()
            else:
                return str(value).strip()
        return str(value).strip()

    def _validate_vat(self, vat_value):
        """Validar y limpiar número de IVA"""
        if not vat_value:
            return False

        vat_clean = self._safe_str(vat_value).upper()
        if not vat_clean or vat_clean in ['NAN', 'NONE', '']:
            return False

        # Si no tiene prefijo de país, agregar ES
        if len(vat_clean) > 2 and not vat_clean[:2].isalpha():
            vat_clean = 'ES' + vat_clean

        return vat_clean

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
        bank_obj = self.env['res.partner.bank']

        # Verificar si el módulo de mandatos bancarios está disponible
        mandate_obj = None
        try:
            mandate_obj = self.env['account.banking.mandate']
        except KeyError:
            _logger.warning("El módulo de mandatos bancarios no está disponible. Se omitirá la creación de mandatos.")

        # Contadores para el mensaje final
        clientes_creados = 0
        clientes_actualizados = 0
        clientes_con_errores = 0
        errores_detalle = []

        # Procesar cada fila
        for row_idx, row in enumerate(rows_data, start=2):  # +2 porque empezamos en fila 2 del Excel
            try:
                # Solo crear o actualizar cliente si tiene CODIGO y NOMBRE
                codigo_raw = row.get('CODIGO', '')
                nombre_raw = row.get('NOMBRE', '')

                codigo = self._safe_str(codigo_raw)
                nombre = self._safe_str(nombre_raw)

                # Validar que no estén vacíos
                if not codigo or not nombre or codigo.lower() in ['nan', 'none'] or nombre.lower() in ['nan', 'none']:
                    continue

                # Buscar cliente existente por CODIGO (usando ref)
                partner = partner_obj.search([('ref', '=', codigo)], limit=1)

                # Buscar o crear país
                pais = None
                if row.get('PAIS'):
                    pais_name = self._safe_str(row.get('PAIS'))
                    if pais_name and pais_name.lower() not in ['nan', 'none']:
                        pais = country_obj.search([('name', 'ilike', pais_name)], limit=1)
                        if not pais:
                            pais = country_obj.search([('code', 'ilike', pais_name)], limit=1)

                # Buscar o crear provincia/estado
                provincia = None
                if row.get('PROVINCIA') and pais:
                    provincia_name = self._safe_str(row.get('PROVINCIA'))
                    if provincia_name and provincia_name.lower() not in ['nan', 'none']:
                        provincia = state_obj.search([
                            ('name', 'ilike', provincia_name),
                            ('country_id', '=', pais.id)
                        ], limit=1)

                # Preparar valores del cliente
                vals = {
                    'name': nombre,
                    'ref': codigo,
                    'is_company': True,  # Por defecto como empresa
                    'customer_rank': 1,  # Marcar como cliente
                    'supplier_rank': 0,  # No es proveedor
                }

                # Campos opcionales
                vat = self._validate_vat(row.get('CIF'))
                if vat:
                    vals['vat'] = vat

                street = self._safe_str(row.get('DIRECCION'))
                if street and street.lower() not in ['nan', 'none']:
                    vals['street'] = street

                city = self._safe_str(row.get('POBLACION'))
                if city and city.lower() not in ['nan', 'none']:
                    vals['city'] = city

                zip_code = self._safe_str(row.get('C_POSTAL'))
                if zip_code and zip_code.lower() not in ['nan', 'none']:
                    vals['zip'] = zip_code

                if pais:
                    vals['country_id'] = pais.id

                if provincia:
                    vals['state_id'] = provincia.id

                phone = self._safe_str(row.get('TELEFONO1'))
                if phone and phone.lower() not in ['nan', 'none']:
                    vals['phone'] = phone

                mobile = self._safe_str(row.get('MOVIL'))
                if mobile and mobile.lower() not in ['nan', 'none']:
                    vals['mobile'] = mobile

                email = self._safe_str(row.get('EMAIL'))
                if email and email.lower() not in ['nan', 'none'] and '@' in email:
                    vals['email'] = email

                website = self._safe_str(row.get('WWW'))
                if website and website.lower() not in ['nan', 'none']:
                    vals['website'] = website

                # Campo activo
                activo = self._safe_str(row.get('ACTIVO', 'Si'))
                vals['active'] = activo.lower() not in ['no', 'false', '0', 'inactivo']

                # Campos adicionales específicos
                nombre_comercial = self._safe_str(row.get('NOMBRE_COMERCIAL'))
                if nombre_comercial and nombre_comercial.lower() not in ['nan', 'none']:
                    vals['commercial_company_name'] = nombre_comercial

                # Agregar información adicional en comentarios
                comment_parts = []

                telefono2 = self._safe_str(row.get('TELEFONO2'))
                if telefono2 and telefono2.lower() not in ['nan', 'none']:
                    comment_parts.append(f"Teléfono 2: {telefono2}")

                persona_contacto = self._safe_str(row.get('PERSONA_DE_CONTACTO'))
                if persona_contacto and persona_contacto.lower() not in ['nan', 'none']:
                    comment_parts.append(f"Persona de contacto: {persona_contacto}")

                if comment_parts:
                    vals['comment'] = '\n'.join(comment_parts)

                # Crear o actualizar cliente
                if partner:
                    # Actualizar cliente existente
                    partner.with_context(skip_vat_validation=True).write(vals)
                    clientes_actualizados += 1
                    _logger.info(f"Cliente actualizado: {codigo} - {nombre}")
                else:
                    # Crear nuevo cliente
                    partner = partner_obj.with_context(skip_vat_validation=True).create(vals)
                    clientes_creados += 1
                    _logger.info(f"Cliente creado: {codigo} - {nombre}")

                # Procesar IBAN_A como banco del cliente
                iban_a = self._safe_str(row.get('IBAN_A'))
                if iban_a and iban_a.lower() not in ['nan', 'none'] and partner:
                    try:
                        # Buscar si ya existe una cuenta bancaria para este partner
                        existing_bank = bank_obj.search([
                            ('partner_id', '=', partner.id),
                            ('acc_number', '=', iban_a)
                        ], limit=1)

                        bank_account = None
                        if not existing_bank:
                            # Crear nueva cuenta bancaria
                            bank_vals = {
                                'partner_id': partner.id,
                                'acc_number': iban_a,
                            }

                            # Si hay BIC_A, agregarlo
                            bic_a = self._safe_str(row.get('BIC_A'))
                            if bic_a and bic_a.lower() not in ['nan', 'none']:
                                bank_vals['bank_bic'] = bic_a

                            bank_account = bank_obj.create(bank_vals)
                            _logger.info(f"Cuenta bancaria creada para cliente {codigo}: {iban_a}")
                        else:
                            bank_account = existing_bank

                        # Crear mandato bancario automáticamente si no existe
                        if bank_account and mandate_obj:
                            existing_mandate = mandate_obj.search([
                                ('partner_bank_id', '=', bank_account.id),
                                ('partner_id', '=', partner.id)
                            ], limit=1)

                            if not existing_mandate:
                                # Crear nuevo mandato bancario
                                mandate_vals = {
                                    'partner_id': partner.id,
                                    'partner_bank_id': bank_account.id,
                                    'signature_date': fields.Date.today(),
                                    'state': 'valid',
                                    'scheme': 'CORE',  # Esquema SEPA CORE por defecto
                                }

                                mandate_obj.create(mandate_vals)
                                _logger.info(f"Mandato bancario creado para cliente {codigo}: {iban_a}")

                    except Exception as bank_error:
                        _logger.warning(f"Error al crear cuenta bancaria/mandato para cliente {codigo}: {str(bank_error)}")

            except Exception as e:
                # Capturar errores y continuar con el siguiente cliente
                clientes_con_errores += 1
                error_msg = f"Fila {row_idx}: {str(e)}"
                errores_detalle.append(error_msg)
                _logger.error(f"Error al procesar cliente en fila {row_idx}: {str(e)}")
                continue

        # Mensaje de resultado
        mensaje = f"""
✅ Importación de clientes completada

📊 Resumen:
• Clientes creados: {clientes_creados}
• Clientes actualizados: {clientes_actualizados}
• Errores encontrados: {clientes_con_errores}
"""

        if errores_detalle:
            mensaje += f"\n⚠️ Errores encontrados:\n" + "\n".join(errores_detalle[:10])
            if len(errores_detalle) > 10:
                mensaje += f"\n... y {len(errores_detalle) - 10} errores más"

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': 'Importación completada',
                'message': mensaje,
                'type': 'success',
                'sticky': True,
            }
        }
