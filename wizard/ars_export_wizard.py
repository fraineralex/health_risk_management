from odoo import models, fields
from xlwt import Workbook, easyxf
from odoo.exceptions import UserError
import base64
import io


class ArsExportWizard(models.TransientModel):
    _name = 'ars.export.wizard'
    _description = 'Wizard to export ARS templates in xlsx and txt formats'
    
    txt_binary = fields.Binary(string='Archivo TXT')
    txt_filename = fields.Char()
    
    xls_binary = fields.Binary(string='Archivo XLS')
    xls_filename = fields.Char()

    def generate_reports(self):
        headers, title = self._get_headers_and_title()
        self._save_reports(title, headers)
        
        return {
            'context': self.env.context,
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'ars.export.wizard',
            'res_id': self.id,
            'view_id': False,
            'type': 'ir.actions.act_window',
            'target': 'new',
        }


    def _get_headers_and_title(self):
        title = 'Reporte'
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


    def _save_reports(self, title, headers):
        workbook, worksheet = self._generate_workbook(headers, title)
        txt_lines = ''

        history_moves = self.env['account.move'].search(
            [('state', '=', 'posted')])

        for col_num, move in enumerate(history_moves, start=1):
            values = {
                'authorization_insurer': '  ',
                'service_date': move.date or '  ',
                'affiliate': '  ',
                'insured_name': move.partner_id.name or '  ',
                'id_number': move.partner_id.vat or '  ',
                'total_claimed': move.amount_total_signed or '  ',
                'service_amount': move.amount_total or '  ',
                'goods_amount': move.good_total_amount or '  ',
                'total_to_pay': move.amount_total or '  ',
                'affiliate_difference': '  ',
                'invoice': move.name or '  ',
                'invoice_date': move.invoice_date or '  ',
                'service_types': move.service_type or '  ',
                'subservice_types': '  ',
                'credit_fiscal_ncf_date': move.l10n_do_ecf_sign_date or '  ',
                'credit_fiscal_ncf': '  ',
                'document_type': move.l10n_latam_document_type_id.name or '  ',
                'ncf_expiration_date': move.ncf_expiration_date or '  ',
                'modified_ncf_nc_or_db': '  ',
                'nc_or_db_amount': move.amount_total or '  ',
                'itbis_amount': move.cost_itbis or '  ',
                'isc_amount': move.amount_tax or '  ',
                'other_taxes_amount': move.other_taxes or '  ',
                'phone': move.partner_id.phone or '  ',
                'cell_phone': move.partner_id.mobile or '  ',
                'email': move.partner_id.email or '  ',
            }
            
            for row_num, (key, value) in enumerate(values.items()):
                worksheet.write(col_num, row_num, value)
            
            txt_lines += self._create_txt_line(values)
        
        self._generate_txt_file(txt_lines, title)
        self._generate_xls_file(workbook, title)


    def _generate_workbook(self, headers, title):
        workbook = Workbook()
        worksheet = workbook.add_sheet(title)
        excel_units = 256
        column_width = 30 * excel_units
        header_style = easyxf(
            'pattern: pattern solid, fore_colour blue; font: colour white, bold True;')
        for col_num, header in enumerate(headers):
            worksheet.col(col_num).width = column_width
            worksheet.write(0, col_num, header, header_style)
        
        return [workbook, worksheet]
    

    def _generate_xls_file(self, workbook, title):
        workbook_data = io.BytesIO()
        workbook.save(workbook_data)
        workbook_data.seek(0)

        report_file_base64 = base64.b64encode(
            workbook_data.read()).decode('utf-8')
        
        self.write({
            'xls_filename': title + '.xls',
            'xls_binary': report_file_base64
        })

 
    def _create_txt_line(self, values):
        txt_line = ''
        for key, value in values.items():
            txt_line += str(value) + '   '
        return txt_line[:-1] + '\n'


    def _generate_txt_file(self, txt_lines, title):
        txt_file = io.BytesIO()
        txt_file.write(txt_lines.encode('utf-8'))
        txt_file.seek(0)
        txt_file_base64 = base64.b64encode(txt_file.read()).decode('utf-8')
            
        self.write({
            'txt_filename': title + '.txt',
            'txt_binary': txt_file_base64
        })
