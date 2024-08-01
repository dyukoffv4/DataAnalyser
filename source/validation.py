import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation


def add_validation(book: openpyxl.Workbook):
    sheet = book['ТИ']

    rule = DataValidation(type="list", formula1='=\'Тип ТИ\'!$A$2:$A$43', allow_blank=True)
    sheet.add_data_validation(rule)
    rule.add("G1:G1048576")

    rule = DataValidation(type="list", formula1='=\'Тип ТИ\'!$L$2:$L$16', allow_blank=True)
    sheet.add_data_validation(rule)
    rule.add("H1:H1048576")

    sheet = book['ТС']

    rule = DataValidation(type="list", formula1='=\'Тип ТИ\'!$L$2:$L$16', allow_blank=True)
    sheet.add_data_validation(rule)
    rule.add("F1:F1048576")

    rule = DataValidation(type="list", formula1='=\'Тип ТИ\'!$D$2:$D$43', allow_blank=True)
    sheet.add_data_validation(rule)
    rule.add("D1:D1048576")

    rule = DataValidation(type="list", formula1='=\'Тип ТИ\'!$I$2:$I$193', allow_blank=True)
    sheet.add_data_validation(rule)
    rule.add("E1:E1048576")

    rule = DataValidation(type="list", formula1='=\'Тип ТИ\'!$G$2:$G$6', allow_blank=True)
    sheet.add_data_validation(rule)
    rule.add("Q1:Q1048576")

    return book