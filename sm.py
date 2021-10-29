import pandas as pd
import math
import os.path, time
import pickle
import time

import email.utils
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from smtplib import SMTP


sheet1, sheet2 = None, None

file = r'R:\Quant\python\controls\inputs\fund_fs.xlsm'

with pd.ExcelFile(file) as reader:
    sheet1 = pd.read_excel(reader, sheet_name='fs_data',convert_float=False)
    #sheet2 = pd.read_excel(reader, sheet_name='Sheet2')

sheet1 = sheet1.loc[sheet1['GROUP'].str.startswith('SC')]

columns = sheet1.columns
date_=int(columns[63])
date_=str(date_)  
    
    
mensaje = MIMEMultipart("plain")
mensaje['From']=email.utils.formataddr(('Cygnus Quant','ialonso@cygnus-am.com'))
#mensaje['To']="front@cygnus-am.com"
mensaje['To']="ialonso@cygnus-am.com"
    
# Target Price

sel ='SM'

col = ['GROUP','NAME','ANALISTAS',sel,sel+"_5D",sel+"_10D",sel+"_1M",sel+"_3M"]   
df = sheet1[col]
df = df.round(decimals=2)
df['chg']=df[sel]-df[sel+'_5D']

trig = 0.15

df_sel = df.loc[(df['chg']<-trig) | (df['chg']>trig)]

if len(df_sel)==0:
    pass
else:
    
    texto = "Cambios Recomendaciones - Consensus (+/-15%) - " +date_ 
    mensaje['Subject']=texto
    
    df_sel['chg'] = df_sel['chg'].astype(float).map("{:.2%}".format)

    html = df_sel.to_html()
    cuerpo = MIMEText(html, 'html')
    mensaje.attach(cuerpo)

    smtp = SMTP("smtp.office365.com", 587)
    smtp.starttls()
    smtp.login("ialonso@cygnus-am.com",'Lanas2021&..')
    #smtp.sendmail("ialonso@cygnus-am.com",["lamusategui@cygnus-am.com","sruizdegaribay@cygnus-am.com","ialonso@cygnus-am.com"],mensaje.as_string())
    smtp.sendmail("ialonso@cygnus-am.com",["ialonso@cygnus-am.com","igor_athletic@yahoo.es"],mensaje.as_string())
    smtp.quit()

