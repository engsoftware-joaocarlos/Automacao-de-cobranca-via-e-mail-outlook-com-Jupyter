#!/usr/bin/env python
# coding: utf-8

# Instalando a biblioteca:

# In[1]:


get_ipython().system('pip install pywin32 ')


# Importando as bibliotecas:

# In[5]:


import win32com.client as client
import pandas as pd
import datetime as dt


# Lendo o arquivo excel:

# In[9]:


tabela = pd.read_excel('Contas a Receber.xlsx')
display(tabela)
tabela.info()


# Verificando data de hoje:

# In[10]:


hoje = dt.datetime.now()
print(hoje)


# Coletando apenas os dados de clientes que estão devendo:

# In[12]:


tabela_devedores = tabela.loc[tabela['Status']=='Em aberto']
display(tabela_devedores)
tabela_devedores = tabela_devedores.loc[tabela_devedores['Data Prevista para pagamento']<hoje]
display(tabela_devedores)


# Enviar um e-mail via Outlook:

# In[13]:


outlook = client.Dispatch('Outlook.Application')


# In[33]:


emissor = outlook.session.Accounts['engsoftware.joaocarlos@gmail.com']


# In[38]:


dados= tabela_devedores[['Valor em aberto','Data Prevista para pagamento','E-mail','NF']].values.tolist()


# Enviando o e-mail para todos os destinatarios que estao com o status Em aberto:

# In[43]:


for dado in dados:
    destinatario = dado[2]
    nf=dado[3]
    prazo=dado[1]
    prazo = prazo.strftime("%d/%m/%Y")
    valor = dado[0]
    assunto = 'Atraso de pagamento'
    mensagem = outlook.CreateItem(0)
    mensagem.Display()
    mensagem.To = destinatario
    mensagem.Subject = destinatario
    corpo_mensagem = f'''
    Prezado Cliente,

    Verificamos um atraso no pagamento referente a NF {nf} com vencimento em {prazo} e valor total de R${valor:.2f}.
    
    Em caso de dúvidas, é só entrar em contato com nosso time atráves do e-mail financeiro@gmail.com
    
    Att,
    
    Joao Carlos
    '''
    mensagem.Body = corpo_mensagem
    mensagem.Save()
    mensagem.Send()


# In[ ]:




