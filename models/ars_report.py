from odoo import models, fields, api, _
from odoo.exceptions import ValidationError
from xlwt import Workbook, easyxf
import base64
import io

class ArsReport(models.Model):
    _name = 'ars.report'
    _description = 'ARS Report'
    
    claimant_code = fields.Char(string='Código Reclamante')
    name = fields.Char(string='Periodo', placeholder='Mes/Año')
    #insurer_id = fields.Many2one('medical.insurance.company', string='Aseguradora')
    insurer_id = fields.Selection([
        ('1', 'Aseguradora 1'),
        ('2', 'Aseguradora 2'),
    ], string='Aseguradora')
    claimant_type = fields.Selection([
        ('medico', 'MEDICO'),
        ('no_medico', 'NO_MEDICO'),
    ], string='Tipo Reclamante')
    line_ids = fields.One2many('ars.report.line', 'report_id', string='Líneas de Reporte')
    
    @api.constrains('name')
    def _check_name(self):
        for record in self:
            if record.name and not self._check_format(record.name):
                raise ValidationError(_("El formato del periodo debe ser 'Mes/Año'"))

    def _check_format(self, name):
        month, year = int(name.split('/')[0]), int(name.split('/')[1])
        return month <= 12 and month > 0 and year > 0
    
    @api.model
    def create(self, vals):
        ars_report = super().create(vals)
        
        history_moves = self.env['account.move'].search(
            [('state', '=', 'posted')])
        line_ids = []
        for move in history_moves:
            line_ids = []
            """ values = {
                'authorization_insurer': move.auth_num,
                'service_date': move.invoice_date,
                'affiliate': move.afiliacion,
                'insured_name': move.partner_id.name,
                'id_number': move.partner_id.vat,
                'total_claimed': move.cober,
                'service_amount': move.service_total_amount,
                'goods_amount': move.good_total_amount,
                'total_to_pay': move.service_total_amount + move.good_total_amount,
                'affiliate_difference': move.cober_diference,
                'invoice': move.name,
                'invoice_date': move.invoice_date,
                'service_types': move.service_type,
                'subservice_types': '',
                'credit_fiscal_ncf_date': move.invoice_date,
                'credit_fiscal_ncf': move.ref,
                'document_type': 'F' if move.type == 'out_invoice' else
                                 'D' if move.is_debit_note else 
                                 'C' if move.type == 'out_invoice' else
                                 '',
                'ncf_expiration_date': move.ncf_expiration_date,
                'modified_ncf_nc_or_db': move.l10n_do_origin_ncf,
                'nc_or_db_amount': move.amount_total,
                'itbis_amount': move.invoiced_itbis,
                'isc_amount': move.selective_tax,
                'other_taxes_amount': move.other_taxes,
                'phone': move.partner_id.phone,
                'cell_phone': move.partner_id.mobile,
                'email': move.partner_id.email,
            } """
            
            values = {
                'authorization_insurer': 'move.auth_num',
                'service_date': move.invoice_date,
                'affiliate': 'move.afiliacion',
                'insured_name': 'move.partner_id.name',
                'id_number': 'move.partner_id.vat',
                'total_claimed': 0.0,
                'service_amount': move.amount_total,
                'goods_amount': 0.0,
                'total_to_pay': 0.0,
                'affiliate_difference': 0.0,
                'invoice': move.name,
                'invoice_date': move.invoice_date,
                'service_types': 'move.service_type',
                'subservice_types': '',
                'credit_fiscal_ncf_date': move.invoice_date,
                'credit_fiscal_ncf': 'move.ref',
                'document_type': 'F',
                'ncf_expiration_date': move.invoice_date,
                'modified_ncf_nc_or_db': 'move.l10n_do_origin_ncf',
                'nc_or_db_amount': move.amount_total,
                'itbis_amount': move.amount_total,
                'isc_amount': move.amount_total,
                'other_taxes_amount': move.amount_total,
                'phone': 'move.partner_id.phone',
                'cell_phone': 'move.partner_id.mobile',
                'email': 'move.partner_id.email',
            }
            
            created_line = self.env['ars.report.line'].create({
                'report_id': ars_report.id,
                **values
            })
            
            line_ids.append(created_line.id)   
            
        ars_report.write({'line_ids': [(6, 0, line_ids)]})

        return ars_report
    
    def action_open_lines(self):
        self.ensure_one()
        return {
            'name': 'Línes del Reporte',
            'type': 'ir.actions.act_window',
            'view_mode': 'tree',
            'res_model': 'ars.report.line',
            'domain': [('report_id', '=', self.id)],
        }
        
    
    def export_to_xlsx(self):
        headers, title = self._get_headers_and_title()
        workbook = self._create_and_populate_xlsx(headers, title)
        workbook_data = io.BytesIO()
        workbook.save(workbook_data)
        workbook_data.seek(0)

        report_file = base64.b64encode(workbook_data.getvalue())
        filename = title + '.xls'
        
        return self._download_report_file(report_file, filename)
    
    
    def _get_headers_and_title(self):
        title = 'Reporte ARS'
        # xls headers
        headers = [
            'AUTORIZACION ASEGURADORA',
            'FECHA SERVICIO',
            'AFILIADO',
            'NOMBRE ASEGURADO',
            'NO. CEDULA DE IDENTIDAD',
            'TOTAL RECLAMADO',
            'MONTO SERVICIO',
            'MONTO BIENES',
            'TOTAL A PAGAR',
            'DIFERENCIA AFILIADO',
            'FACTURA',
            'FECHA FACTURA',
            'TIPOS DE SERVICIOS',
            'SUB-TIPO DE SERVICIOS',
            'FECHA NCF Credito Fiscal',
            'NCF Credito Fiscal',
            'TIPO COMPROBANTE',
            'FECHA VENCIMIENTO NCF',
            'NCF Modificado (NC y/o DB)',
            'Monto NC y/o DB',
            'MONTO ITBIS',
            'MONTO ISC',
            'MONTO OTROS IMPUESTOS',
            'Teléfono',
            'Celular',
            'Correo Electrónico'
        ]
        
        return [headers, title]
    
    
    def _create_and_populate_xlsx(self, headers, title):
        workbook = Workbook()
        worksheet = workbook.add_sheet(title)
        excel_units = 256
        column_width = 30 * excel_units
        header_style = easyxf(
            'pattern: pattern solid, fore_colour blue; font: colour white, bold True;')
        
        # write headers with their styles
        for col_num, header in enumerate(headers):
            worksheet.col(col_num).width = column_width
            worksheet.write(0, col_num, header, header_style)
        
        for col_num, line in enumerate(self.line_ids, start=1):
            values = {
                'authorization_insurer': line.authorization_insurer,
                'service_date': line.service_date,
                'affiliate': line.affiliate,
                'insured_name': line.insured_name,
                'id_number': line.id_number,
                'total_claimed': line.total_claimed,
                'service_amount': line.service_amount,
                'goods_amount': line.goods_amount,
                'total_to_pay': line.total_to_pay,
                'affiliate_difference': line.affiliate_difference,
                'invoice': line.invoice,
                'invoice_date': line.invoice_date,
                'service_types': line.service_types,
                'subservice_types': line.subservice_types,
                'credit_fiscal_ncf_date': line.credit_fiscal_ncf_date,
                'credit_fiscal_ncf': line.credit_fiscal_ncf,
                'document_type': line.document_type,
                'ncf_expiration_date': line.ncf_expiration_date,
                'modified_ncf_nc_or_db': line.modified_ncf_nc_or_db,
                'nc_or_db_amount': line.nc_or_db_amount,
                'itbis_amount': line.itbis_amount,
                'isc_amount': line.isc_amount,
                'other_taxes_amount': line.other_taxes_amount,
                'phone': line.phone,
                'cell_phone': line.cell_phone,
                'email': line.email,
            }
                        
            # write values in the worksheet
            for row_num, (key, value) in enumerate(values.items()):
                worksheet.write(col_num, row_num, value or '')
        
        return workbook
    
    def _download_report_file(self, report_file, filename):
        # Crea un objeto Attachments para el archivo adjunto
        attachment = self.env['ir.attachment'].create({
            'name': filename,
            'datas': report_file,
            'mimetype': 'application/vnd.ms-excel' if filename.endswith('.xls') else 'text/plain',
            'res_model': self._name,
            'res_id': self.id,
        })

        # Devuelve la acción para descargar el archivo adjunto
        return {
            'name': 'Download',
            'type': 'ir.actions.act_url',
            'url': '/web/content/%s?download=true' % attachment.id,
            'target': 'self',
        }
        
        
    def export_to_txt(self):
        txt_lines = ''
        for col_num, line in enumerate(self.line_ids, start=1):
            values = {
                'authorization_insurer': line.authorization_insurer,
                'service_date': line.service_date,
                'affiliate': line.affiliate,
                'insured_name': line.insured_name,
                'id_number': line.id_number,
                'total_claimed': line.total_claimed,
                'service_amount': line.service_amount,
                'goods_amount': line.goods_amount,
                'total_to_pay': line.total_to_pay,
                'affiliate_difference': line.affiliate_difference,
                'invoice': line.invoice,
                'invoice_date': line.invoice_date,
                'service_types': line.service_types,
                'subservice_types': line.subservice_types,
                'credit_fiscal_ncf_date': line.credit_fiscal_ncf_date,
                'credit_fiscal_ncf': line.credit_fiscal_ncf,
                'document_type': line.document_type,
                'ncf_expiration_date': line.ncf_expiration_date,
                'modified_ncf_nc_or_db': line.modified_ncf_nc_or_db,
                'nc_or_db_amount': line.nc_or_db_amount,
                'itbis_amount': line.itbis_amount,
                'isc_amount': line.isc_amount,
                'other_taxes_amount': line.other_taxes_amount,
                'phone': line.phone,
                'cell_phone': line.cell_phone,
                'email': line.email,
            }
            
            txt_lines += self._create_txt_line(values)
        
        txt_file = io.BytesIO()
        txt_file.write(txt_lines.encode('utf-8'))
        txt_file.seek(0)
        txt_file_base64 = base64.b64encode(txt_file.read()).decode('utf-8')
        
        return self._download_report_file(txt_file_base64, 'Reporte ARS.txt')
        
    def _create_txt_line(self, values):
        txt_line = ''
        for key, value in values.items():
            # if value is None, replace it with an empty string
            chunk = str(value or '') + '   '
            if value == '':
                chunk = '      ' # double tab
            txt_line += chunk 
        return txt_line[:-1] + '\n'
        
        

class ArsReportLines(models.Model):
    _name = 'ars.report.line'
    _description = 'ARS Report Lines'

    report_id = fields.Many2one('ars.report', string='Reporte ARS',
        index=True, required=True, readonly=True, auto_join=True, ondelete="cascade",
        help="El reporte ARS al que pertenece esta línea.")
    authorization_insurer = fields.Char('Número de Autorización del Seguro')
    service_date = fields.Date('Fecha del Servicio')
    affiliate = fields.Char('Afiliado')
    insured_name = fields.Char('Nombre del Asegurado')
    id_number = fields.Char('Número de Identificación')
    total_claimed = fields.Float('Total Reclamado')
    service_amount = fields.Float('Monto del Servicio')
    goods_amount = fields.Float('Monto de los Bienes')
    total_to_pay = fields.Float('Total a Pagar')
    affiliate_difference = fields.Float('Diferencia del Afiliado')
    invoice = fields.Char('Factura')
    invoice_date = fields.Date('Fecha de la Factura')
    service_types = fields.Char('Tipos de Servicio')
    subservice_types = fields.Char('Tipos de Subservicio')
    credit_fiscal_ncf_date = fields.Date('Fecha del NCF del Crédito Fiscal')
    credit_fiscal_ncf = fields.Char('NCF del Crédito Fiscal')
    document_type = fields.Selection([
        ('F', 'Factura'),
        ('D', 'Nota de Débito'),
        ('C', 'Nota de Crédito'),
        ('', 'Ninguno')
    ], 'Tipo de Documento')
    ncf_expiration_date = fields.Date('Fecha de Expiración del NCF')
    modified_ncf_nc_or_db = fields.Char('NCF Modificado NC o DB')
    nc_or_db_amount = fields.Float('Monto NC o DB')
    itbis_amount = fields.Float('Monto ITBIS')
    isc_amount = fields.Float('Monto ISC')
    other_taxes_amount = fields.Float('Monto de Otros Impuestos')
    phone = fields.Char('Teléfono')
    cell_phone = fields.Char('Teléfono Móvil')
    email = fields.Char('Correo Electrónico')

    def export_to_xlsx(self):
        print('export_to_xlsx')
        
    def export_to_txt(self):
        print('export_to_txt')
