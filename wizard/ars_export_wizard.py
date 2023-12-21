from odoo import models, fields
from xlwt import Workbook, easyxf
from odoo.exceptions import UserError
import base64
from werkzeug import urls
import io


class ArsExportWizard(models.TransientModel):
    _name = 'ars.export.wizard'
    _description = 'Wizard to export ARS templates in xlsx and txt formats'

    def create_report_xlsx(self):
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

        title = 'Reporte'

        workbook = self._create_workbook(title, headers)
        workbook_data = io.BytesIO()
        workbook.save(workbook_data)
        workbook_data.seek(0)

        excel_base64 = base64.b64decode(workbook_data.getvalue())
        filename = urls.url_quote(title + '.xls')
        report_file = excel_base64
        report_filename = filename

        return self._action_save(workbook_data, report_filename)

    def create_report_txt(self):
        print("create_report_txt")

    def _create_workbook(self, title, headers):
        workbook = Workbook()
        worksheet = workbook.add_sheet(title)
        excel_units = 256
        column_width = 30 * excel_units
        header_style = easyxf(
            'pattern: pattern solid, fore_colour blue; font: colour white, bold True;')
        for col_num, header in enumerate(headers):
            worksheet.col(col_num).width = column_width
            worksheet.write(0, col_num, header, header_style)

        history_moves = self.env['account.move'].search(
            [('state', '=', 'posted')])

        for col_num, move in enumerate(history_moves, start=1):
            authorization_insurer = ''
            service_date = move.date or ''
            affiliate = ''
            insured_name = move.partner_id.name or ''
            id_number = move.partner_id.vat or ''
            total_claimed = move.amount_total_signed or ''
            service_amount = move.amount_total or ''
            goods_amount = move.good_total_amount or ''
            total_to_pay = move.amount_total or ''
            affiliate_difference = ''
            invoice = move.name or ''
            invoice_date = move.invoice_date or ''
            service_types = move.service_type or ''
            subservice_types = ''
            credit_fiscal_ncf_date = move.l10n_do_ecf_sign_date or ''
            credit_fiscal_ncf = ''
            document_type = move.l10n_latam_document_type_id.name or ''
            ncf_expiration_date = move.ncf_expiration_date or ''
            modified_ncf_nc_or_db = ''
            nc_or_db_amount = move.amount_total or ''
            itbis_amount = move.cost_itbis or ''
            isc_amount = move.amount_tax or ''
            other_taxes_amount = move.other_taxes or ''
            phone = move.partner_id.phone or ''
            cell_phone = move.partner_id.mobile or ''
            email = move.partner_id.email or ''

            worksheet.write(col_num, 0, authorization_insurer)
            worksheet.write(col_num, 1, service_date)
            worksheet.write(col_num, 2, affiliate)
            worksheet.write(col_num, 3, insured_name)
            worksheet.write(col_num, 4, id_number)
            worksheet.write(col_num, 5, total_claimed)
            worksheet.write(col_num, 6, service_amount)
            worksheet.write(col_num, 7, goods_amount)
            worksheet.write(col_num, 8, total_to_pay)
            worksheet.write(col_num, 9, affiliate_difference)
            worksheet.write(col_num, 10, invoice)
            worksheet.write(col_num, 11, invoice_date)
            worksheet.write(col_num, 12, service_types)
            worksheet.write(col_num, 13, subservice_types)
            worksheet.write(col_num, 14, credit_fiscal_ncf_date)
            worksheet.write(col_num, 15, credit_fiscal_ncf)
            worksheet.write(col_num, 16, document_type)
            worksheet.write(col_num, 17, ncf_expiration_date)
            worksheet.write(col_num, 18, modified_ncf_nc_or_db)
            worksheet.write(col_num, 19, nc_or_db_amount)
            worksheet.write(col_num, 20, itbis_amount)
            worksheet.write(col_num, 21, isc_amount)
            worksheet.write(col_num, 22, other_taxes_amount)
            worksheet.write(col_num, 23, phone)
            worksheet.write(col_num, 24, cell_phone)
            worksheet.write(col_num, 25, email)

        return workbook

    def _action_save(self, report_file, report_filename):
        report_file_base64 = base64.b64encode(
            report_file.read()).decode('utf-8')
        attachment = self.env['ir.attachment'].create({
            'name': report_filename,
            'datas': report_file_base64,
            'mimetype': 'application/vnd.ms-excel',
            'res_model': self._name,
            'res_id': self.id,
        })

        return {
            'name': 'Download',
            'type': 'ir.actions.act_url',
            'url': '/web/content/%s?download=true' % attachment.id,
        }
