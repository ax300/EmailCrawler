from bs4 import BeautifulSoup as soup  # HTML data structure
from urllib.request import urlopen as uReq  # Web client
import mechanize #sudo pip install python-mechanize

page_url = 'https://mail.google.com/mail/u/0/#inbox'

def logaConta():
    br = mechanize.Browser() #initiating a browser
    br.set_handle_robots(False) #ignore robots.txt

    br.addheaders = [("User-agent","Chrome")] #our identity
    br.open(page_url) #requesting the google base url

    br.select_form(nr=0)
    br.form['Email'] = 'andre.cavicchiolli@viaconsulting.com.br'
    #para acessar control sem nome <None> eh necessario guardar forms em lista e iterar pelos controles ou acessar o None no indice especifico
    forms = [f for f in br.forms()]
    forms[0].controls[17].value = 'ax300naruto'
    #print(*forms)

    #for x in range(len(forms[0].controls)):
        #print(forms[0].controls[x].name)
        #if(forms[0].controls[x].name == 'Email'):
            #print(br.form[0].controls[x])
            #br.form[0].controls[x] = 'andre.cavicchiolli@usp.br'
    br.submit()

    return br





# opens the connection and downloads html page from url
#uClient = uReq(page_url)

# parses html into a soup data structure to traverse html
# as if it were a json data type.

print(logaConta())

page_soup = soup(logaConta().response().read(), "html.parser")

#uClient.close()

conteiners = page_soup.findAll("tbody")

#print(page_soup.find(id=':1t'))

#print(len(conteiners))
