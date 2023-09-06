from urllib.request import urlopen, Request
from bs4 import BeautifulSoup
import win32com.client as win32

'''
Não deixar vix passar de 30% da margem
Rolar de qualquer forma se passar de menos que 30 dias uteis
'''

list =['pair_2', 'pair_3', 'pair_4', 'pair_5', 'pair_6', 'pair_7',
       'pair_8', 'pair_9', 'pair_10']

r = Request('https://br.investing.com/indices/us-spx-vix-futures-contracts',
            headers={'User-Agent': 'Mozilla/5.0'})
response = urlopen(r).read()
soup = BeautifulSoup(response, "html.parser")

#mandaemail('VEnda coberta Dupla', f'<p>Abeve chegou a meta {result}</p>')
def mandaemail(titulo, mensagem):
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    email.to = 'aventureiro-msn@hotmail.com'
    email.Subject = titulo
    email.HTMLBody = mensagem
    email.Send()

def vix():
    lista = ['pair_2', 'pair_3', 'pair_4', 'pair_5', 'pair_6', 'pair_7',
             'pair_8', 'pair_9', 'pair_10']
    lista2 = []
    for lis in lista:
        table = soup.find(id=lis)
        valor = round(float(table.text[7:13]), 1)
        lista2.append(valor)
    if lista2[0] >= lista2[1]:
        print('rolagem primeiro mes')
        mandaemail('VIX', '<p> Rolagem possivel</p>')
    elif lista2[1] >= lista2[2]:
        print('rolagem segundo mes')
        mandaemail('VIX', '<p> Rolagem possivel</p>')
    elif lista2[2] >= lista2[3]:
        print('rolagem terceiro mes')
        mandaemail('VIX', '<p> Rolagem possivel</p>')
    elif lista2[3] >= lista2[4]:
        print('rolagem quarto mes')
        mandaemail('VIX', '<p> Rolagem possivel</p>')
    elif lista2[4] >= lista2[5]:
        print('rolagem quinto mes')
        mandaemail('VIX', '<p> Rolagem possivel</p>')
    elif lista2[5] >= lista2[6]:
        print('rolagem sexto mes')
        mandaemail('VIX', '<p> Rolagem possivel</p>')
    elif lista2[6] >= lista2[7]:
        print('rolagem setimo mes')
        mandaemail('VIX', '<p> Rolagem possivel</p>')
    elif lista2[7] >= lista2[8]:
        print('rolagem oitavo mes')
        mandaemail('VIX', '<p> Rolagem possivel</p>')
    elif lista2[8] >= lista2[9]:
        print('rolagem nono mes')
        mandaemail('VIX', '<p> Rolagem possivel</p>')
    print('*'*100)
    return print(lista2), mandaemail('VIX', f'<p> {lista2} </p>')

vix()
