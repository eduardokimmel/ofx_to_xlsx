import os, re, openpyxl
import datetime as datetime


#  Uncomment if name == main
#import gettext
#gettext.install('ofx_to_xlsx')
#t = gettext.translation('gui_i18n', 'locale', fallback=True)
#_ = t.gettext


def run(ofx):
    os.chdir(os.path.dirname(ofx)) 
    regex = re.compile(r"(<\w+>)(.+)(</\w+>)")
    #arquivo = open(ofx,'r')
    with open(ofx, 'r') as arquivo:
            
        texto = arquivo.read()
        texto = texto.split('\n')
    
        #  Creates workbook, sets as active and set columns titles
        
        wb = openpyxl.Workbook()
        ws = wb.active
        
        ws['A1'] = _('Bank Code')
        ws['B1'] = _('AccType')
        ws['C1'] = _('Date')
        ws['D1'] = _('Transaction Type')
        ws['E1'] = _('Value')
        ws['F1'] = _('Memo')
        ws['G1'] = _('ID')
        
        n = 2 # Line on xlsx
        
        for i in texto:
            try:
                reg = regex.findall(i)
                if reg[0][0] == '<BANKID>':
                    bco = reg[0][1]
                if reg[0][0] == '<ACCTID>':
                    cc = reg[0][1]
                if reg[0][0] == '<TRNTYPE>':
                    ws['D'+str(n)] = reg[0][1]
                if reg[0][0] == '<DTPOSTED>':
                    reg = reg[0][1]
                    ws['C'+str(n)] = datetime.date(int(reg[0:4]), int(reg[4:6]), int(reg[6:8]))
                    ws['C'+str(n)].number_format = 'dd/mm/yyyy'
                if reg[0][0] == '<TRNAMT>':
                    reg = reg[0][1]
                    ws['E'+str(n)] = float(reg.replace(",","."))
                if reg[0][0] == '<FITID>':
                    ws['G'+str(n)] = reg[0][1]
                if reg[0][0] == '<MEMO>':
                    ws['F'+str(n)] = reg[0][1]
                    ws['A'+str(n)] = bco
                    ws['B'+str(n)] = cc
                    n += 1
                    
            except IndexError:
                continue
        
        
        #  Save File
        dest_filename = os.path.basename(ofx[:-4]) + '.xlsx'
        wb.save(filename = dest_filename)
        wb.close()
    
