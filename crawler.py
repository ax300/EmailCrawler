# Importing libraries
import imaplib, email, getpass, html2text, re, datetime, os.path, xlwt, xlrd


#user = 'andre.cavicchiolli@usp.br'
imap_url = 'imap.gmail.com'
user = ''
password = ''
  
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
                #msg = email.message_from_string(str(response_part[1]).strip())
                #print("subject: " + msg['Subject'])
                regex = 'Nubank'
                FROM = re.findall(regex,msg['from'])
                #print("from: " + FROM[0])
                #print("Delivered to: " + msg['Delivered-To'])
                #print("Date: " + msg['Date'])
                timestamp = datetime.datetime.strptime(msg['Date'].split(', ')[1].split(' +')[0], '%d %b %Y %H:%M:%S')
                #print(timestamp.strftime('%s'))
                c = get_body(msg)
                #print(html2text.html2text(c.decode('utf-8')))
                #print("===========================================================================================================================")
    return msgs 

#cria planilha e a preenche com as formulas
def create_sheet(file_name):
    f = open("base.xls", "w")
    # Insere regras e headers da planilha
    book = xlwt.Workbook()
    #campos e formulas
    sheet = book.add_sheet(file_name)
    sheet.write(0, 1, 'Janeiro')
    sheet.write(0, 2, 'Fevereiro')
    sheet.write(0, 3, 'Março')
    sheet.write(0, 4, 'Abril')
    sheet.write(0, 5, 'Maio')
    sheet.write(0, 6, 'Junho')
    sheet.write(0, 7, 'Julho')
    sheet.write(0, 8, 'agosto')
    sheet.write(0, 9, 'Setembro')
    sheet.write(0, 10, 'Outubro')
    sheet.write(0, 11, 'Novembro')
    sheet.write(0, 12, 'Dezembro')
    sheet.write(1,0,'Salário')
    sheet.write(4,0,'Total Mensal')
    sheet.write(5,0,'Total')
    #Totais Mensais
    sheet.write(4, 1,  '=SOMA(B2:B4)')
    sheet.write(4, 2,  '=SOMA(C2:C4)')
    sheet.write(4, 3,  '=SOMA(D2:D4)')
    sheet.write(4, 4,  '=SOMA(F2:F4)')
    sheet.write(4, 5,  '=SOMA(G2:G4)')
    sheet.write(4, 6,  '=SOMA(H2:H4)')
    sheet.write(4, 7,  '=SOMA(I2:I4)')
    sheet.write(4, 8,  '=SOMA(J2:J4)')
    sheet.write(4, 9,  '=SOMA(K2:K4)')
    sheet.write(4, 10, '=SOMA(L2:L4)')
    sheet.write(4, 11, '=SOMA(M2:M4)')
    sheet.write(4, 12, '=SOMA(N2:N4)')
    #Totais
    sheet.write(5, 1,  '=SOMA(B5;A6)')
    sheet.write(5, 2,  '=SOMA(C5;B6)')
    sheet.write(5, 3,  '=SOMA(D5;C6)')
    sheet.write(5, 4,  '=SOMA(F5;D6)')
    sheet.write(5, 5,  '=SOMA(G5;E6)')
    sheet.write(5, 6,  '=SOMA(H5;F6)')
    sheet.write(5, 7,  '=SOMA(I5;G6)')
    sheet.write(5, 8,  '=SOMA(J5;H6)')
    sheet.write(5, 9,  '=SOMA(K5;I6)')
    sheet.write(5, 10, '=SOMA(L5;J6)')
    sheet.write(5, 11, '=SOMA(M5;K6)')
    sheet.write(5, 12, '=SOMA(N5;L6)')
    book.save("base.xls")

# this is done to make SSL connnection with GMAIL 
con = imaplib.IMAP4_SSL(imap_url)  
# Pega informacoes de login ou mantem as atuais
login_info()
# logging the user in 
con.login(user,password)
  
# calling function to check for email under this label 

con.select('Inbox')  
  
 # fetching emails from this user "tu**h*****1@gmail.com" 
msgs = get_emails(search('FROM', 'todomundo@nubank.com.br', con)) 
#msgs = get_emails(search(None,'all' con))   

# verifica se a planilha com os dados ja existe
year = datetime.datetime.today().year
file_name = str(year)
if (os.path.isfile("base.xls") == False):
    print('cria')
    create_sheet(file_name)
    