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
