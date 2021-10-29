#! /usr/bin/python

import os, sys
import pdfrw


ANNOT_KEY = '/Annots'           # key for all annotations within a page
ANNOT_FIELD_KEY = '/T'          # Name of field. i.e. given ID of field
ANNOT_FORM_type = '/FT'         # Form type (e.g. text/button)
ANNOT_FORM_button = '/Btn'      # ID for buttons, i.e. a checkbox
ANNOT_FORM_text = '/Tx'         # ID for textbox
SUBTYPE_KEY = '/Subtype'
WIDGET_SUBTYPE_KEY = '/Widget'

data_dict={
    'page0_Line4_DaytimeTelephoneNumber[0]':'Mouren',
    'page0_Pt1Line2c_MiddleName[0]':'Mail'
}
#data_dict={
#}

def string_escape(s, encoding='utf-8'):
    return (s.encode('latin1')         # To bytes, required by 'unicode-escape'
             .decode('unicode-escape') # Perform the actual octal-escaping decode
             .encode('latin1')         # 1:1 mapping back to bytes
             .decode(encoding))        # Decode original encoding

from openpyxl import Workbook
def inspect(input_pdf_path, input_excel_path=None):
    key_list = []
    template_pdf=pdfrw.PdfReader(input_pdf_path)
    for page_number, page in enumerate(template_pdf.pages):
        annotations=page[ANNOT_KEY]
        for annotation in annotations:
            if annotation[SUBTYPE_KEY]==WIDGET_SUBTYPE_KEY:
                key=annotation[ANNOT_FIELD_KEY][1:-1]
                try:
                    key = str(bytes.fromhex(key).decode('utf-16'))
                except ValueError:
                    print('Please decrpt the file before inspect.')
                    print('Using python autofill.py decrypt command')
                cat_key = 'page' + str(page_number) + '_' +key
                print(cat_key)
                key_list.append(cat_key)
    if input_excel_path is None:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = os.path.basename(input_pdf_path)[0:-4]
        for key_row, key in enumerate(key_list):
            sheet['A'+str(key_row + 1)] = key
        path = os.getcwd() + '/'
        workbook.save(filename = path + os.path.basename(input_pdf_path)[0:-4] + '.xlsx')
    else:
        workbook = load_workbook(input_excel_path)
        sheet_name = os.path.basename(input_pdf_path)[0:-4]
        sheet = workbook.create_sheet(title=sheet_name)
        for key_row, key in enumerate(key_list):
            sheet['A'+str(key_row + 1)] = key
        print('Data write to sheet ' + sheet.title)
        workbook.save(filename = path + os.path.basename(input_excel_path))


def write_fillable_pdf(input_pdf_path,output_pdf_path,data_dict):
    if not os.path.isfile(input_pdf_path):
        print('Could not find ' + input_pdf_path + '. Skip.')
        return
    template_pdf=pdfrw.PdfReader(input_pdf_path)
    for page_number, page in enumerate(template_pdf.pages):
        annotations=page[ANNOT_KEY]
        for annotation in annotations:
            if annotation[SUBTYPE_KEY]==WIDGET_SUBTYPE_KEY:
                key=annotation[ANNOT_FIELD_KEY][1:-1]
                try:
                    key = str(bytes.fromhex(key).decode('utf-16'))
                except ValueError:
                    print('Please decrpt the file before inspect.')
                    print('Using python autofill.py decrypt command')
                key = 'page' + str(page_number) + '_' +key
                try:
                    if annotation[ANNOT_FORM_type] == ANNOT_FORM_button:
                        annotation.update(
                            pdfrw.PdfDict(V=pdfrw.PdfName(data_dict[key]), AS=pdfrw.PdfName(data_dict[key])) # default checkbox value is 'Off'
                        )
                    else:
                        annotation.update(
                            pdfrw.PdfDict(V='{}'.format(data_dict[key]))
                        )
                except KeyError:
                    print('Missing in: ' + input_pdf_path)
                    print('Missing: ' + key)
                        
    pdfrw.PdfWriter().write(output_pdf_path,template_pdf)
    
import pikepdf
def decrpt(input_pdf_path):
    pdf = pikepdf.open(input_pdf_path, allow_overwriting_input=True)
    pdf.save(input_pdf_path)


from openpyxl import load_workbook
def read_excel(input_excel_path):
    workbook = load_workbook(filename=input_excel_path)
    sheet = workbook.active
    print('Extracting from ' + sheet.title)
    for data in sheet.iter_rows(values_only=True):
        print(data[0], data[1])
        data_dict[data[0]] = data[1]

def run_all(input_excel_path):
    path = os.getcwd() + '/'
    workbook = load_workbook(filename=input_excel_path)
    for sheet_name_index, sheet_name in enumerate(workbook.sheetnames):
        if sheet_name != 'Custom Info':
            data_dict = {}
            sheet = workbook.worksheets[sheet_name_index]
            print('Extracting from ' + sheet.title)
            for data in sheet.iter_rows(min_col=1, max_col=3, values_only=True):
                if data[1] is None:
                    print(data[0], '')
                    data_dict[data[0]] = ''
                else:
                    print(data[0], data[1])
                    data_dict[data[0]] = data[1]
            write_fillable_pdf(path + sheet_name + '.pdf', path + sheet_name + '-fill.pdf', data_dict)
    
    
if __name__ == '__main__':
    if sys.argv[1] == 'inspect':
        if len(sys.argv) == 3:
            inspect(sys.argv[2])
        elif len(sys.argv) == 4:
            inspect(sys.argv[2], sys.argv[3])
    elif sys.argv[1] == 'decrpt':
        decrpt(sys.argv[2])
    elif sys.argv[1] == 'write':
        write_fillable_pdf(sys.argv[2], sys.argv[3], data_dict)
    elif sys.argv[1] == 'read_excel':
        read_excel(sys.argv[2])
    else:
        run_all(sys.argv[1])
