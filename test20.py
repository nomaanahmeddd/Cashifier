from bs4 import BeautifulSoup
import requests
import time
import smtplib
import ssl
from email.message import EmailMessage
import openpyxl

# Open the workbook
wb = openpyxl.load_workbook('SCRAPF.xlsx')

# Get the sheet
sheet=wb['Sheet2']

# Iterate through the rows
for row in sheet.iter_rows():
    # Get the values for each cell in the row
    values = [cell.value for cell in row]
    
    # Store the values as variables
    var1, var2, var3, var4 = values
receiver=var3
URL=var2
target=var4

headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"}
page=requests.get(URL, headers=headers) 
soup1=BeautifulSoup(page.text,"html.parser")

### amazon ###
def amazon():
    print('₹',pricea.replace(',','').replace('.',''))

###  flipkart  ###
def flipkart():
    print('₹',pricef.replace(',','').replace('₹',''))


if URL[12]=='a':
    prod=soup1.find(id='productTitle').text
    pricea=soup1.find(class_='a-price-whole').text
    price=float(pricea.replace(',','').replace('₹',''))
    print(prod.strip())
    amazon()
elif URL[12]=='f':
    prod=soup1.find(class_='B_NuCI').text
    pricef=soup1.find(class_='_30jeq3 _16Jk6d').text
    price=float(pricef.replace(',','').replace('₹',''))
    print(prod.strip())
    flipkart()

## email ##
def em():
  if price<target:
          sender="scraptivists@gmail.com"
          password="gcxnqozqluclhyxm"
          subject = "Price drop in your selected product!!"
          body="{}'s price has dropped to ₹{} currently\n.Hurry up and buy it soon!! \nBuy here: {}".format(prod,price,URL)
          em=EmailMessage()
          em['From']=sender
          em['To']=receiver
          em['Subject']=subject
          em.set_content(body)
          context=ssl.create_default_context()
          with smtplib.SMTP_SSL('smtp.gmail.com',465,context=context) as smtp:
              smtp.login(sender,password)
              smtp.sendmail(sender,receiver,em.as_string())
em()