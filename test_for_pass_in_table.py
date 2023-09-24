from docxtpl import DocxTemplate
import openpyxl


def main():
    wb = openpyxl.load_workbook('data_staff_pass_fishport.xlsx', data_only=True)
    sheet = wb['список_людей']

    data_list = [
        {
            'id': i - 1,
            **{sheet.cell(1, col).value: sheet.cell(i, col).value for col in range(2, 11)}
        }
        for i in range(2, sheet.max_row + 1)
    ]

    context = {'data_list': data_list}
    doc = DocxTemplate('template_order_staff_pass_fishport.docx')
    doc.render(context)
    doc.save('data_list_result.docx')

if __name__ == '__main__':
    main()
