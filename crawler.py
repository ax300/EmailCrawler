# Importing libraries
import imaplib, email, getpass, html2text

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
                print("subject: " + str(msg['Subject']))
                #print("To:" + str(msg['To']))
                print("from: " + str(msg['from']))
                print("Delivered to:" + str(msg['Delivered-To']))
                #print("Received:" + str(msg['Received']))
                #print("X-Google-Smtp-Source:" + str(msg['X-Google-Smtp-Source']))
                #print("X-Received:" + str(msg['X-Received']))
                #print("ARC-Seal:" + str(msg['ARC-Seal']))
                #print("ARC-Message-Signature:" + str(msg['ARC-Message-Signature']))
                #print("ARC-Authentication-Results:" + str(msg['ARC-Authentication-Results']))
                #print("Return-Path:" + str(msg['Return-Path']))
                #print("Received-SPF:" + str(msg['Received-SPF']))
                #print("Authentication-Results:" + str(msg['Authentication-Results']))
                #print("DKIM-Signature:" + str(msg['DKIM-Signature']))
                #print("X-Report-Abuse:" + str(msg['X-Report-Abuse']))
                #print("X-Mandrill-User:" + str(msg['X-Mandrill-User']))
                #print("Message-Id:" + str(msg['Message-Id']))
                print("Date:" + str(msg['Date']))
                #print("MIME-Version:" + str(msg['MIME-Version']))
                #print("Content-Type:" + str(msg['Content-Type']))  
                c=msg.get_payload(None, True)
                print(type(c))
                print(html2text.html2text(str(c)))
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