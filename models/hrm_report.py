from odoo import models, fields, api, _
from odoo.exceptions import ValidationError
from xlwt import Workbook, easyxf
import base64
import io
from datetime import datetime, timedelta


class HrmTemplateReport(models.Model):
    _name = 'hrm.template.report'
    _description = 'HRM Template Report'

    claimant_code = fields.Char(string='Claimant Code')
    name = fields.Char(string='Period', placeholder='Month/Year')
    insurer_id = fields.Many2one('medical.insurance.company', string='Insurer', required=True)
    claimant_type = fields.Selection([
        ('medico', 'MEDICAL'),
        ('no_medico', 'NON_MEDICAL'),
    ], string='Claimant Type')
    line_ids = fields.One2many('hrm.template.report.line', 'report_id', string='Report Lines')
    date_from = fields.Char(string='Start Date', required=True)
    date_to = fields.Char(string='End Date', required=True)

    @api.constrains('date_from', 'date_to')
    def _check_name(self):
        for record in self:
            if record.date_from and record.date_to:
                self._check_format(record.date_from, record.date_to)

    def _check_format(self, date_from, date_to):
        month_from, year_from = int(date_from.split('/')[0]), int(date_from.split('/')[1])
        month_to, year_to = int(date_to.split('/')[0]), int(date_to.split('/')[1])

        if month_from < 1 or month_from > 12 or month_to < 1 or month_to > 12:
            raise ValidationError(_("Invalid month, you must enter a month between 1-12."))
        elif year_from < 1 or year_to < 1:
            raise ValidationError(_("Invalid year, you must enter a value greater than zero."))
        elif year_from > year_to or (year_from == year_to and month_from > month_to):
            raise ValidationError(_("Invalid values, the end date must be greater than the start date."))
        else:
            self.name = '{}/{}-{}/{}'.format(month_from, year_from, month_to, year_to)

    @api.model
    def create(self, vals):
        report = super().create(vals)
        if not report.date_from or not report.date_to or not report.insurer_id:
            raise ValidationError(_("The fields 'Start Date', 'End Date' and 'Insurer' are required."))
        
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
            #('hrm', '=', report.insurer_id.id) # add this line to filter by insurer
        ])

        line_ids = []
        for move in history_moves:
            values = {
                'report_id': report.id,
                'authorization_insurer': move.name, # change this to the field that contains the insurer authorization
                'service_date': move.invoice_date,
                'affiliate': move.name, # change this to the field that contains the affiliate
                'insured_name': move.partner_id.name, # change this to the field that contains the insured name
                'id_number': move.partner_id.vat, # change this to the field that contains the id number
                'total_claimed': move.amount_total, # change this to the field that contains the total claimed
                'service_amount': move.amount_total, 
                'goods_amount': move.amount_total, # change this to the field that contains the goods amount
                'total_to_pay': move.amount_total + move.amount_tax, # change this to the field that contains the total to pay
                'affiliate_difference': move.amount_total, # change this for move.service_total_amount + move.good_total_amount,
                'invoice': move.name,
                'invoice_date': move.invoice_date,
                'service_types': move.move_type,
                'subservice_types': move.move_type, # change this to the field that contains the subservice types
                'credit_fiscal_ncf_date': move.invoice_date,
                'credit_fiscal_ncf': move.ref,
                'document_type': 'F' if move.move_type == 'out_invoice' else
                'C' if move.move_type == 'out_invoice' else
                '',
                'ncf_expiration_date': move.invoice_date, # change this to the field that contains the ncf expiration date
                'modified_ncf_nc_or_db': move.name, # change this to the field that contains the modified ncf
                'nc_or_db_amount': move.amount_total,
                'itbis_amount': move.amount_tax, # change this to the field that contains the itbis amount
                'isc_amount': move.amount_tax, # change this to the field that contains the isc amount
                'other_taxes_amount': move.amount_tax, # change this to the field that contains the other taxes amount
                'phone': move.partner_id.phone,
                'cell_phone': move.partner_id.mobile,
                'email': move.partner_id.email,
            }

            created_line = self.env['hrm.template.report.line'].create(values)
            line_ids.append(created_line.id)

        report.write({'line_ids': [(6, 0, line_ids)]})

        return report

    def action_open_lines(self):
        self.ensure_one()
        return {
            'name': 'Report Lines',
            'type': 'ir.actions.act_window',
            'view_mode': 'tree',
            'res_model': 'hrm.template.report.line',
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
        title = 'Health Risk Managers Report'
        headers = [
            'INSURER AUTHORIZATION',
            'SERVICE DATE',
            'AFFILIATE',
            'INSURED NAME',
            'IDENTITY CARD NO.',
            'TOTAL CLAIMED',
            'SERVICE AMOUNT',
            'GOODS AMOUNT',
            'TOTAL TO PAY',
            'AFFILIATE DIFFERENCE',
            'INVOICE',
            'INVOICE DATE',
            'TYPES OF SERVICES',
            'SUB-TYPE OF SERVICES',
            'NCF Tax Credit Date',
            'NCF Tax Credit',
            'TYPE OF VOUCHER',
            'NCF EXPIRATION DATE',
            'Modified NCF (NC and/or DB)',
            'Amount NC and/or DB',
            'ITBIS AMOUNT',
            'ISC AMOUNT',
            'OTHER TAXES AMOUNT',
            'Phone',
            'Cellphone',
            'Email'
        ]

        return [headers, title]

    def _create_and_populate_xlsx(self, headers, title):
        workbook = Workbook()
        worksheet = workbook.add_sheet(title)
        excel_units = 256
        column_width = 30 * excel_units
        header_style = easyxf(
            'pattern: pattern solid, fore_colour blue; font: colour white, bold True;')

        for col_num, header in enumerate(headers):
            worksheet.col(col_num).width = column_width
            worksheet.write(0, col_num, header, header_style)

        for col_num, line in enumerate(self.line_ids, start=1):
            values = self._map_line_values(line)
            
            for row_num, (key, value) in enumerate(values.items()):
                worksheet.write(col_num, row_num, value or '')

        return workbook

    def _download_report_file(self, report_file, filename):
        attachment = self.env['ir.attachment'].create({
            'name': filename,
            'datas': report_file,
            'mimetype': 'application/vnd.ms-excel' if filename.endswith('.xls') else 'text/plain',
            'res_model': self._name,
            'res_id': self.id,
        })

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

        return self._download_report_file(txt_file_base64, 'Health Risk Managers Report.txt')

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


class HrmTemplateReportLines(models.Model):
    _name = 'hrm.template.report.line'
    _description = 'HRM Report Lines'
    
    report_id = fields.Many2one('hrm.template.report', string='HRM Report',
                                index=True, required=True, readonly=True, auto_join=True, ondelete="cascade",
                                help="The HRM report to which this line belongs.")
    authorization_insurer = fields.Char('Insurer Authorization Number')
    service_date = fields.Date('Service Date')
    affiliate = fields.Char('Affiliate')
    insured_name = fields.Char('Insured Name')
    id_number = fields.Char('Identification Number')
    total_claimed = fields.Float('Total Claimed')
    service_amount = fields.Float('Service Amount')
    goods_amount = fields.Float('Goods Amount')
    total_to_pay = fields.Float('Total to Pay')
    affiliate_difference = fields.Float('Affiliate Difference')
    invoice = fields.Char('Invoice')
    invoice_date = fields.Date('Invoice Date')
    service_types = fields.Char('Service Types')
    subservice_types = fields.Char('Subservice Types')
    credit_fiscal_ncf_date = fields.Date('NCF Tax Credit Date')
    credit_fiscal_ncf = fields.Char('NCF Tax Credit')
    document_type = fields.Selection([
        ('F', 'Invoice'),
        ('D', 'Debit Note'),
        ('C', 'Credit Note'),
        ('', 'None')
    ], 'Document Type')
    ncf_expiration_date = fields.Date('NCF Expiration Date')
    modified_ncf_nc_or_db = fields.Char('Modified NCF NC or DB')
    nc_or_db_amount = fields.Float('NC or DB Amount')
    itbis_amount = fields.Float('ITBIS Amount')
    isc_amount = fields.Float('ISC Amount')
    other_taxes_amount = fields.Float('Other Taxes Amount')
    phone = fields.Char('Phone')
    cell_phone = fields.Char('Mobile Phone')
    email = fields.Char('Email')
