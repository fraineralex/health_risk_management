from odoo import models, fields, api, _
from odoo.exceptions import ValidationError
from xlwt import Workbook, easyxf
import base64
import io
from datetime import datetime, timedelta


class ArsTemplateReport(models.Model):
    _name = 'ars.template.report'
    _description = 'ARS Template Report'

    claimant_code = fields.Char(string='Código Reclamante')
    name = fields.Char(string='Periodo', placeholder='Mes/Año')
    insurer_id = fields.Char(string='Aseguradora') #fields.Many2one('medical.insurance.company', string='Aseguradora', required=True)
    claimant_type = fields.Selection([
        ('medico', 'MEDICO'),
        ('no_medico', 'NO_MEDICO'),
    ], string='Tipo Reclamante')
    line_ids = fields.One2many('ars.template.report.line', 'report_id', string='Líneas de Reporte')
    date_from = fields.Char(string='Fecha Inicio', required=True)
    date_to = fields.Char(string='Fecha Fin', required=True)

    @api.constrains('date_from', 'date_to')
    def _check_name(self):
        for record in self:
            if record.date_from and record.date_to:
                self._check_format(record.date_from, record.date_to)

    def _check_format(self, date_from, date_to):
        month_from, year_from = int(date_from.split('/')[0]), int(date_from.split('/')[1])
        month_to, year_to = int(date_to.split('/')[0]), int(date_to.split('/')[1])

        if month_from < 1 or month_from > 12 or month_to < 1 or month_to > 12:
            raise ValidationError(_("Mes invalido, debes ingresar un mes entre 1-12."))
        elif year_from < 1 or year_to < 1:
            raise ValidationError(_("Año invalido, debes ingresar un valor mayor que cero."))
        elif year_from > year_to or (year_from == year_to and month_from > month_to):
            raise ValidationError(_("Valores invalidos, la fecha final debe ser mayor que la fecha de inicio."))
        else:
            self.name = '{}/{}-{}/{}'.format(month_from, year_from, month_to, year_to)

    @api.model
    def create(self, vals):
        report = super().create(vals)
        if not report.date_from or not report.date_to or not report.insurer_id:
            raise ValidationError(_("Los campos 'Fecha Inicio', 'Fecha Fin' y 'Aseguradora' son requeridos."))
        
        self.generate_report(report.id)
        return report
    
    @api.model
    def generate_report(self, report_id):
        report = self.browse(report_id)
        
        date_from = datetime.strptime('01/' + report.date_from, '%d/%m/%Y')
        date_to = datetime.strptime('01/' + report.date_to, '%d/%m/%Y')

        if date_to.month != 12:
            next_month = date_to.replace(month=date_to.month + 1, day=1)
            last_day_of_month = (next_month - timedelta(days=1)).day
        else:
            next_year = date_to.replace(month=1, day=1, year=date_to.year + 1)
            last_day_of_month = (next_year - timedelta(days=1)).day
            
        date_to = date_to.replace(day=last_day_of_month)
        
        history_moves = self.env['account.move'].search([
            ('state', '=', 'posted'),
            ('invoice_date', '>=', date_from),
            ('invoice_date', '<=', date_to),
            #('ars', '=', report.insurer_id.id)
        ])

        line_ids = []
        for move in history_moves:
            values = {
                'report_id': report.id,
                'authorization_insurer': "move.auth_num",
                'service_date': move.invoice_date,
                'affiliate': "move.afiliacion",
                'insured_name': "move.partner_id.name",
                'id_number': 10 or move.partner_id.vat,
                'total_claimed': 10 or move.cober,
                'service_amount': 10 or move.service_total_amount,
                'goods_amount': 10 or move.good_total_amount,
                'total_to_pay': 10 or move.service_total_amount + move.good_total_amount,
                'affiliate_difference': 10 or move.cober_diference,
                'invoice': "move.name",
                'invoice_date': move.invoice_date,
                'service_types': "move.service_type",
                'subservice_types': "move.subservice_type",
                'credit_fiscal_ncf_date': move.invoice_date,
                'credit_fiscal_ncf': "move.ref",
                'document_type': 'F' if move.type == 'out_invoice' else
                #'D' if move.is_debit_note else
                #'C' if move.type == 'out_invoice' else
                '',
                'ncf_expiration_date': move.invoice_date or move.ncf_expiration_date,
                'modified_ncf_nc_or_db': 'move.l10n_do_origin_ncf',
                'nc_or_db_amount': 10 or move.amount_total,
                'itbis_amount': 10 or move.invoiced_itbis,
                'isc_amount': 10 or move.selective_tax,
                'other_taxes_amount': 10 or move.other_taxes,
                'phone': 'move.partner_id.phone',
                'cell_phone': 'move.partner_id.mobile',
                'email': 'move.partner_id.email',
            }

            created_line = self.env['ars.template.report.line'].create(values)
            line_ids.append(created_line.id)

        report.write({'line_ids': [(6, 0, line_ids)]})

        return report

    def action_open_lines(self):
        self.ensure_one()
        return {
            'name': 'Línes del Reporte',
            'type': 'ir.actions.act_window',
            'view_mode': 'tree',
            'res_model': 'ars.template.report.line',
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
            values = self._map_line_values(line)
            
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
            values = self._map_line_values(line)
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
                chunk = '      '  # double tab
            txt_line += chunk
        return txt_line[:-1] + '\n'
    
    def _map_line_values(self, line):
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
        return values


class ArsTemplateReportLines(models.Model):
    _name = 'ars.template.report.line'
    _description = 'ARS Report Lines'

    report_id = fields.Many2one('ars.template.report', string='Reporte ARS',
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
