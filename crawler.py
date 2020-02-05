# Importing libraries
import imaplib, email, getpass, html2text, re, datetime, os.path, xlwt, xlrd, itertools
from xlutils.copy import copy as xl_copy

#user = 'andre.cavicchiolli@usp.br'
imap_url = 'imap.gmail.com'
user = ''
password = ''
meses = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
meses1 = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']
year = datetime.datetime.today().year
sheet_name = str(year)
  
#pega informacoes de login e as guarda num txt  
def get_pass():
    
    user = input("Entre com o novo usuário: ")
    
    password = getpass.getpass("Senha: ")

    f = open("Pass.txt", "w")
    f.write(user+","+password)
    f.close()

    return user, password
    
#mostra informacoes do txt para saber se deve manter o login ou nao
def login_info():
    global user,password
    f = open("Pass.txt", "r")
    UserNpass = f.read()
    if(UserNpass):
        user,password = UserNpass.split(',')
        placeholder = '*' * len(password)
        
        print("Usuário: "+ user +"\nSenha: " + placeholder)
        
        if(input("Manter login?")  == '1'):
            print("\nOk")

        else:
            get_pass()

    else:
        get_pass()

# Function to get email content part i.e its body part 
def get_body(msg): 
    if msg.is_multipart(): 
        return get_body(msg.get_payload(0)) 
    else: 
        return msg.get_payload(None, True) 
  
# Function to search for a key value pair  
def search(key, value, con):  
    result, data = con.search(None, key, '"{}"'.format(value)) 
    #result, data = con.search(None, 'ALL') 
    return data 
  
#Decodifica strings ainda codificadas
def decode_mime_words(s):
    return u''.join(
        word.decode(encoding or 'utf8') if isinstance(word, bytes) else word
        for word, encoding in email.header.decode_header(s))

# Function to get the list of emails under this label 
def get_emails(result_bytes): 
    msgs = [] # all the email data are pushed inside an array 
    for num in result_bytes[0].split(): 
        typ, data = con.fetch(num, '(RFC822)') 
        msgs.append(data) 
        
        for response_part in data:
            #passa por cada mensagem do email
            if isinstance(response_part, tuple):
                part = response_part[1].decode('utf-8')
                msg = email.message_from_string(part)
                regex = 'Nubank'
                FROM = re.findall(regex,msg['from'])
                print(define_subject(msg))
                print("Delivered to: " + msg['Delivered-To'])
                print("Date: " + msg['Date'])
                timestamp = datetime.datetime.strptime(msg['Date'].split(', ')[1].split(' +')[0], '%d %b %Y %H:%M:%S')
                #print("MÊS: " + str(timestamp.month))
                #print(timestamp.strftime('%s'))
                cat_transaction(msg, timestamp)
                print("===========================================================================================================================")

    return msgs 

#cria aba na planilha
def create_sheet(sheet_name):
    # Insere regras e headers da planilha
    book1 = xlrd.open_workbook('base.xlsx',  formatting_info=True)
    book = xl_copy(book1)
    sheet = book.add_sheet(sheet_name)
    #print(book1.sheet_names())
    #fill_sheet(book1)
    fill_sheet(sheet, book)

#cria arquivo
def create_file(sheet_name):
    book = xlwt.Workbook()
    sheet = book.add_sheet(sheet_name)
    fill_sheet(sheet,book)

#preenche planilha
def fill_sheet(sheet, book):
    #campos e formulas
    sheet.write(0, 1,   'Janeiro')
    sheet.write(0, 2,   'Fevereiro')
    sheet.write(0, 3,   'Março')
    sheet.write(0, 4,   'Abril')
    sheet.write(0, 5,   'Maio')
    sheet.write(0, 6,   'Junho')
    sheet.write(0, 7,   'Julho')
    sheet.write(0, 8,   'agosto')
    sheet.write(0, 9,   'Setembro')
    sheet.write(0, 10,  'Outubro')
    sheet.write(0, 11,  'Novembro')
    sheet.write(0, 12,  'Dezembro')
    sheet.write(1,0,    'Salário')
    sheet.write(99,0,   'Total Mensal')
    sheet.write(100,0,  'Total')
    #Totais Mensais
    sheet.write(99, 1,  xlwt.Formula('SUM(B2:B98)'))
    sheet.write(99, 2,  xlwt.Formula('SUM(C2:C98)'))
    sheet.write(99, 3,  xlwt.Formula('SUM(D2:D98)'))
    sheet.write(99, 4,  xlwt.Formula('SUM(E2:E98)'))
    sheet.write(99, 5,  xlwt.Formula('SUM(F2:F98)'))
    sheet.write(99, 6,  xlwt.Formula('SUM(G2:G98)'))
    sheet.write(99, 7,  xlwt.Formula('SUM(H2:H98)'))
    sheet.write(99, 8,  xlwt.Formula('SUM(I2:I98)'))
    sheet.write(99, 9,  xlwt.Formula('SUM(J2:J98)'))
    sheet.write(99, 10, xlwt.Formula('SUM(K2:K98)'))
    sheet.write(99, 11, xlwt.Formula('SUM(L2:L98)'))
    sheet.write(99, 12, xlwt.Formula('SUM(M2:M98)'))
    #Totais
    sheet.write(100, 1,  xlwt.Formula('SUM(B100;A101)'))
    sheet.write(100, 2,  xlwt.Formula('SUM(C100;B101)'))
    sheet.write(100, 3,  xlwt.Formula('SUM(D100;C101)'))
    sheet.write(100, 4,  xlwt.Formula('SUM(E100;D101)'))
    sheet.write(100, 5,  xlwt.Formula('SUM(F100;E101)'))
    sheet.write(100, 6,  xlwt.Formula('SUM(G100;F101)'))
    sheet.write(100, 7,  xlwt.Formula('SUM(H100;G101)'))
    sheet.write(100, 8,  xlwt.Formula('SUM(I100;H101)'))
    sheet.write(100, 9,  xlwt.Formula('SUM(J100;I101)'))
    sheet.write(100, 10, xlwt.Formula('SUM(K100;J101)'))
    sheet.write(100, 11, xlwt.Formula('SUM(L100;K101)'))
    sheet.write(100, 12, xlwt.Formula('SUM(M100;L101)'))
    first_col = sheet.col(0)
    first_col.width = 256 * 50
    book.save("base.xlsx")

#insere os valores da planilha das transacoes
def get_values(msg, category, timestamp):
    c = get_body(msg).decode('utf-8')
    message = html2text.html2text(c) 

    #filtra o valor transferido/pago da mensagem
    pattern_of_value = re.compile(r'(((\d){1,3}(\.\d\d\d)*)|\d+)(,\d+)') 
    value = pattern_of_value.findall(message)
    value = value[0][0] + value[0][len(value[0])-1]
    print("Valor: " + str(value))
   
    #filtra a entidade da qual recebeu/enviou tal valor
    pattern_of_entity = re.compile(r'\*+([A-Za-z\s^\nÀ-ÖØ-öø-ÿ!,]*?)\*+')
    match_entity = pattern_of_entity.findall(message)
    pattern_of_entity2 = re.compile(r'\*+\s*(.*?)\n?(.*?)\s*\*+')
    match_entity2 = pattern_of_entity2.findall(message)

    if (category == 1 ):
        #procura entidade
        entity = ''.join(match_entity2[2])
        print("ACHEI A ENTIDADE2: " + entity)
        insert_sheet(category, timestamp, value[0], entity)

    elif (category == 2):
        #procura entidade
        print("ACHEI A ENTIDADE: " + match_entity[1])
        insert_sheet(category, timestamp, value[0], match_entity[1])

    elif (category == 3):
        #procura entidade
        print("ACHEI A ENTIDADE: " + match_entity[1])
        insert_sheet(category, timestamp, value[0], match_entity[1])

    elif (category == 4):
        #fatura  
        print('Entidade: Nubank')
        insert_sheet(category, timestamp, value[0], 'Nubank')

#preenche o subject com valor sem codificacao
def define_subject(msg):
    if(msg['Subject'].find('=?UTF-8?') != -1 or msg['Subject'].find('=?utf-8?') != -1):
        return decode_mime_words(msg['Subject'])
    else:
        return msg['Subject']

    return 'Error Subject'

#categoriza transacao
def cat_transaction(msg,timestamp):
    category= 0
    if(define_subject(msg).find('Pagamento realizado') != -1):
        category = 1
        get_values(msg, category,timestamp)
    elif(define_subject(msg).find('Transferência realizada') != -1):
        category = 2
        get_values(msg, category,timestamp)
    elif(define_subject(msg).find('recebeu uma transferência!') != -1):
        category = 3
        get_values(msg, category,timestamp)
    elif(define_subject(msg).find('Pagamento de fatura') != -1):
        category = 4
        get_values(msg, category,timestamp)    

#pega o nome da planilha do ano atual
def get_sheet_by_name(book):
    try:
        for idx in itertools.count():
            sheet = book.get_sheet(idx)
            if sheet.name == sheet_name:
                return sheet
    except IndexError:
        print("Planilha não existe.")
        return None

#devolve o indice da entidade da qual fez a transacao
def get_entity_index(entity):
    book = xlrd.open_workbook('base.xlsx')
    temporary_sheet = book.sheet_by_name(sheet_name)
    #print(type(temporary_sheet))
    #num_rows = temporary_sheet.nrow
    num_rows = 100
    for index in range(2,num_rows):
        cell_val= temporary_sheet.cell(index, 0).value
        #caso a entidade ja exista ou n 
        if(cell_val == entity or cell_val == ''):
            break
    return index
    
#insere na linha a transacao 
def insert_sheet(category, timestamp, value, entity):
    book1 = xlrd.open_workbook('base.xlsx',  formatting_info=True)
    #copia formatacao da planilha
    book = xl_copy(book1)
    sheet = get_sheet_by_name(book)
    #insere entidade numa coluna fixa
    setOutCell(sheet, 0, get_entity_index(entity), entity)
    #insere valor
    setOutCell(sheet, timestamp.month, get_entity_index(entity), value)
    book.save("base.xlsx")
    
def get_current_cell_value (col, row):
    book = xlrd.open_workbook('base.xlsx')
    temporary_sheet = book.sheet_by_name(sheet_name)
    cell_val = temporary_sheet.cell(row, col).value

    return cell_val

def _getOutCell(outSheet, colIndex, rowIndex):
    """ HACK: Extract the internal xlwt cell representation. """
    row = outSheet._Worksheet__rows.get(rowIndex)
    if not row: return None

    cell = row._Row__cells.get(colIndex)
    return cell

def setOutCell(outSheet, col, row, value):
    """ Change cell value without changing formatting. """
    # HACK to retain cell style.
    #print(type(outSheet))
    previousCell = _getOutCell(outSheet, col, row)
    # END HACK, PART I
    old_value = get_current_cell_value(col,row)
    #print(type(old_value))
    if old_value != '' and type(old_value) == int:
        outSheet.write(row, col, int(value) + int(old_value))
    else:
        outSheet.write(row, col, value)

    # HACK, PART II
    if previousCell:
        newCell = _getOutCell(outSheet, col, row)
        if newCell:
            newCell.xf_idx = previousCell.xf_idx
    # END HACK
'''
setOutCell(outSheet, 5, 5, 'Test')
outBook.save('output.xls')    
'''

# this is done to make SSL connnection with GMAIL 
con = imaplib.IMAP4_SSL(imap_url)  
# Pega informacoes de login ou mantem as atuais
login_info()
# logging the user in 
con.login(user,password)
  
# calling function to check for email under this label 

con.select('Inbox')  
  
# verifica se a planilha com os dados ja existe
if (os.path.isfile("base.xlsx") == False):
    print('cria')
    create_file(sheet_name)
elif ((sheet_name in xlrd.open_workbook("base.xlsx").sheet_names())== False):
    print('cria sheet')
    create_sheet(sheet_name)

 # fetching emails from this user "tu**h*****1@gmail.com" 

msgs = get_emails(search('FROM', 'todomundo@nubank.com.br', con)) 


#msgs = get_emails(search(None,'all' con))   


