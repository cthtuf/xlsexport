from io import BytesIO
import xlsxwriter

from django.shortcuts import render
from django.http import HttpResponse
from django.views import View


def WriteToExcel(export_data):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output)

    worksheet = workbook.add_worksheet("Order")
    title = workbook.add_format({
        'bold': True,
        'font_size': 14,
        'align': 'center',
        'valign': 'vcenter',
    })

    header = workbook.add_format({
        'bg_color': '#F7F7F7',
        'color': '000000',
        'align': 'center',
        'valign': 'top',
        'border': 1,
    })

    title_text = "%s %s" % ('Заявка', 'выфв')
    worksheet.merge_range('B2:H2', title_text, title)
    worksheet.write(4, 0, 'tadam', header)
    worksheet.write(4, 1, 'tadam', header)
    worksheet.write(4, 2, 'tadam', header)

    workbook.close()
    xlsx_data = output.getvalue()
    return xlsx_data

class Export(View):

    def get(self, request, *args, **kwargs):
        response = HttpResponse(content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment; filename=Order.xlsx'
        params = request.GET
        xlsx_data = WriteToExcel(params)
        response.write(xlsx_data)
        return response
