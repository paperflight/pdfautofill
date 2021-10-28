# pdfautofill
An auto filling script from excel data to pdf

Use data from excel sheets and fill the pdf with the same name of the sheet. The output is name-fill.pdf.

## Usage
Inspect the possible filling elements:

      python autofill.py 'inspect' <pdf path>
  
Read data from active sheet:
   
      python autofill.py 'read_excel' <excel path>

Fill PDF:
   
      python autofill.py <excel path>
