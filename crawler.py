# Importing libraries
import imaplib, email, getpass, html2text, re
import urllib.request
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
                print("subject: " + msg['Subject'])
                #print("To:" + str(msg['To']))
                regex = 'Nubank'
                FROM = re.findall(regex,msg['from'])
                print("from: " + FROM[0])
                print("Delivered to:" + msg['Delivered-To'])
                print("Date: " + msg['Date'])
                c = get_body(msg)
                #q = quote(c, safe=' <>="/:!')
                print(html2text.html2text(c.decode('utf-8')))
                print("===========================================================================================================================")
    return msgs 
  
# this is done to make SSL connnection with GMAIL 
con = imaplib.IMAP4_SSL(imap_url)  

login_info()
# logging the user in 
con.login(user,password)
  
# calling function to check for email under this label 

con.select('Inbox')  
  
 # fetching emails from this user "tu**h*****1@gmail.com" 
msgs = get_emails(search('FROM', 'todomundo@nubank.com.br', con)) 
#msgs = get_emails(search(None,'all' con))   

# Uncomment this to see what actually comes as data  
#print(msgs) 
#  

"""
x = 1
for msg in msgs:
    
    print('Message ' + str(x) + ': ')
    print(msg['from'])
    x+=1
"""


# Finding the required content from our msgs 
# User can make custom changes in this part to 
# fetch the required content he / she needs 
  
# printing them by the order they are displayed in your gmail  
"""
for msg in msgs[::-1]:  
    for sent in msg: 
        if type(sent) is tuple:  
  
            # encoding set as utf-8 
            content = str(sent[1], 'utf-8')  
            data = str(content) 
  
            # Handling errors related to unicodenecode 
            try:  
                indexstart = data.find("ltr") 
                data2 = data[indexstart + 5: len(data)] 
                indexend = data2.find("</div>") 
  
                # printtng the required content which we need 
                # to extract from our email i.e our body 
                print(data2[0: indexend]) 
  
            except UnicodeEncodeError as e: 
                pass

"""