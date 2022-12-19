# encoding=utf-8
# -*- coding: encoding -*-
import unicodedata
from docx import Document
from docx.shared import Inches
import xlrd
import xlwt
import os
import time
import datetime
import calendar
import sys
import win32api
import shutil
import random
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email import Encoders
import smtplib
import time
import datetime
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email import Encoders
import os

####stile foglio excel
style = xlwt.XFStyle()

# font
font = xlwt.Font()
font.bold = True
font.height = 320
style.font = font

# borders

borders = xlwt.Borders()
borders.bottom = xlwt.Borders.THIN
borders.top = xlwt.Borders.THIN
borders.right = xlwt.Borders.THIN
borders.left = xlwt.Borders.THIN
style.borders = borders


style1 = xlwt.XFStyle()

# font
font = xlwt.Font()
font.height = 320
style1.font = font
# borders

borders = xlwt.Borders()
borders.bottom = xlwt.Borders.THIN
borders.top = xlwt.Borders.THIN
borders.right = xlwt.Borders.THIN
borders.left = xlwt.Borders.THIN
style1.borders = borders

style2 = xlwt.XFStyle()
# font
font = xlwt.Font()
font.height = 320
style2.font = font

#style2 = xlwt.easyxf('font: bold 1, color red;')
style2 = xlwt.easyxf('pattern: pattern solid, fore_colour red;')

# borders
borders = xlwt.Borders()
borders.bottom = xlwt.Borders.THIN
borders.top = xlwt.Borders.THIN
borders.right = xlwt.Borders.THIN
borders.left = xlwt.Borders.THIN
style2.borders = borders

#### fine stile foglio excel

def convertXLS2CSV_manipolazione(aFile,nomefile):
    '''converts a MS Excel file to csv w/ the same name in the same directory'''
    print "------ beginning to convert XLS to CSV ------"
    try:
        import win32com.client, os
        excel = win32com.client.Dispatch('Excel.Application')
        fileDir, fileName = os.path.split(aFile)
        nameOnly = os.path.splitext(fileName)
        newName = "C:\\Users\\f.altarocca\\Desktop\\"+nomefile+".csv"
        outCSV = os.path.join(fileDir, newName)
        workbook = excel.Workbooks.Open(aFile)
        workbook.SaveAs(outCSV, FileFormat=24) # 24 represents xlCSVMSDOS
        workbook.Close(False)
        excel.Quit()
        del excel
        print "...Converted " + nameOnly [0]+ " to CSV"
    except:
        print ">>>>>>> FAILED to convert " + aFile + " to CSV!"

def manipolazionexls(nomelistino):
    aFile='C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\'+nomelistino+'.xls'
    nomefile=nomelistino
    convertXLS2CSV_manipolazione(aFile,nomefile)
    f=open('C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\'+nomelistino+'.csv','r')
    ordini=f.readlines()
    f.close()
    linee=[]
    for elemento in ordini:
        riga=elemento.split(';')
        if len(riga)>28:
            linee.append(riga[0])
            linee.append(riga[6])
            linee.append(riga[15])
            linee.append(riga[26])
            linee.append(riga[29])

    del linee[0:5]

    now = datetime.datetime.today()
    anno=str(now)[0:4]
    mese=str(now)[5:7]
    giorno=str(now)[8:10]
    MyFile = xlwt.Workbook()
    MyFoglio = MyFile.add_sheet(nomelistino+'_'+giorno+"_"+mese+"_"+anno, cell_overwrite_ok = True)
    MyFoglio.col(0).width = 6500
    MyFoglio.col(1).width = 6500
    MyFoglio.col(2).width = 6500
    MyFoglio.col(3).width = 6500
    MyFoglio.col(4).width = 6500
    MyFoglio.write(0,0,'CODICE',style=style)
    MyFoglio.write(0,1,'DESCRIZIONE',style=style)
    MyFoglio.write(0,2,'PREZZO',style=style)
    MyFoglio.write(0,3,'PRODUTTORE',style=style)
    MyFoglio.write(0,4,'PROVENIENZA',style=style)
    #print len(linee)
    riga=1
    contatore=0
    while contatore<=len(linee)-5:
        colonna=0
        while colonna<=4:
            MyFoglio.write(riga,colonna,linee[contatore].decode("iso-8859-1"),style=style1)
            colonna=colonna+1
            contatore=contatore+1
        riga=riga+1
    os.rename('C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\'+nomelistino+'.xls','C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\'+nomelistino+'_origin.xls')
    MyFile.save(r'C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\'+nomelistino+'.xls')




gmail_user = "felice@biosolidale.it"
gmail_pwd = "BioFelice"

def mail(to, subject,text,nomefile):
    msg = MIMEMultipart()
    msg['From'] = gmail_user
    msg['To'] = to
    msg['Subject'] = subject
    msg.attach(MIMEText(text))
    mail_file = MIMEBase('application', 'csv')
    mail_file.set_payload(open('C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\'+nomefile, 'rb').read())
    mail_file.add_header('Content-Disposition', 'attachment', filename=nomefile)
    Encoders.encode_base64(mail_file)
    msg.attach(mail_file)
    mailServer = smtplib.SMTP("smtp.gmail.com", 587)
    mailServer.ehlo()
    mailServer.starttls()
    mailServer.ehlo()
    mailServer.login(gmail_user, gmail_pwd)
    mailServer.sendmail(gmail_user, to, msg.as_string())
    mailServer.close()

def convertXLS2CSV(aFile,nomefile):
    '''converts a MS Excel file to csv w/ the same name in the same directory'''
    print "------ beginning to convert XLS to CSV ------"
    try:
        import win32com.client, os
        excel = win32com.client.Dispatch('Excel.Application')
        fileDir, fileName = os.path.split(aFile)
        nameOnly = os.path.splitext(fileName)
        newName = "C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\"+nomefile+".csv"
        outCSV = os.path.join(fileDir, newName)
        workbook = excel.Workbooks.Open(aFile)
        workbook.SaveAs(outCSV, FileFormat=24) # 24 represents xlCSVMSDOS
        workbook.Close(False)
        #excel.Quit()
        del excel
        print "...Converted " + nameOnly [0]+ " to CSV"
    except:
        print ">>>>>>> FAILED to convert " + aFile + " to CSV!"



###conversione xlsx-->xls
listafilesdacambiare=os.listdir("C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\")
for elemento in listafilesdacambiare:
    rigaelemento=elemento.split('.')
    rigaelemento[1]=rigaelemento[1].upper()
    if rigaelemento[1]=='XLSX':
        os.rename("C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\"+elemento,"C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\"+rigaelemento[0]+".xls")
listafilesdacambiare=os.listdir("C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\")
###conversione xlsx-->xls



#execfile('C:\\Users\\f.altarocca\\Desktop\\Scripts\\Lavorazione Dati\\controllo magazzino\\controllo magazzino2.9.py')
#execfile('C:\\Users\\f.altarocca\\Desktop\\Scripts\\Lavorazione Dati\\controllo magazzino\\Non_aperti_2.1.py')

#crea csv
indirizzostr="C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\neg.xls"
convertXLS2CSV(indirizzostr,"neg")
indirizzostr="C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\gr.xls"
convertXLS2CSV(indirizzostr,"gr")
indirizzostr="C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\gas.xls"
convertXLS2CSV(indirizzostr,"gas")
indirizzostr="C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\pri20.xls"
convertXLS2CSV(indirizzostr,"pri20")
#crea csv

nomelistino_man='gas'
manipolazionexls(nomelistino_man)
nomelistino_man='gr'
manipolazionexls(nomelistino_man)
#nomelistino_man='neg'
#manipolazionexls(nomelistino_man)

today = datetime.date.today()
differenza_inizio = datetime.timedelta(days=3)
differenza_fine = datetime.timedelta(days=9)
inizio_settimana = today + differenza_inizio
fine_settimana = today + differenza_fine
if inizio_settimana.month != fine_settimana.month:
    mese=fine_settimana.month
else:
    mese=inizio_settimana.month


soggetto='listino dal '+ str(inizio_settimana.day) + ' al '+ str(fine_settimana.day) + ' ' + str(mese) + ' ' + str(fine_settimana.year)
document='In allegato il listino della prossima settimana'

files=os.listdir('C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\')
f=open('C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\destinatari\\destinatari.txt','r')
destinatari=f.readlines()
f.close()



esiste=os.path.exists('C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\listini\\'+str(today.day)+'_'+str(today.month)+'_'+str(today.year))
if not esiste:
    os.mkdir('C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\listini\\'+str(today.day)+'_'+str(today.month)+'_'+str(today.year))


f1=open('C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\listini\\'+str(today.day)+'_'+str(today.month)+'_'+str(today.year)+'\\log.txt','a')
#Spedizione listini
i=0
while i<=len(destinatari)-1:
    if '@' not in destinatari[i]:
        nome_tipo=destinatari[i].split('-')
        estensione=nome_tipo[0]
        nomelistino=nome_tipo[1]
        filedaspedire=nomelistino+'.'+estensione
    else:

        #listamail=destinatari[i+1].split(';')
        #listamail[len(listamail)-1]=listamail[len(listamail)-1][:-1]
        #filedaspedire=nomelistino+'.'+estensione
        #for sendto in listamail:

        if '\n' in destinatari[i]:
            destinatari[i]=destinatari[i][:-1]
        f1.write('DESTINATARIO: '+destinatari[i]+'\n')
        #f1.write(soggetto+'\n')
        #f1.write(document+'\n')
        f1.write('FILE SPEDITO: '+filedaspedire+'\n\n')
        soggetto2='DESTINATARIO: '+destinatari[i]+'\n'+ 'FILE SPEDITO: '+filedaspedire+'\n\n'
        mail(destinatari[i],soggetto,document,filedaspedire)
        mail('felice@biosolidale.it',soggetto2,document,filedaspedire)
    i=i+1
#Spedizione listini
f1.close()

"""
--- GOOGLE DRIVE ---
esiste=os.path.exists('C:\\Users\\f.altarocca\\Google Drive\\listini\\'+str(inizio_settimana.year))
if not esiste:
    os.mkdir('C:\\Users\\f.altarocca\\Google Drive\\listini\\'+str(inizio_settimana.year))
esiste=os.path.exists('C:\\Users\\f.altarocca\\Google Drive\\listini\\'+str(inizio_settimana.year)+'\\'+str(inizio_settimana.month)+'\\')
if not esiste:
    os.mkdir('C:\\Users\\f.altarocca\\Google Drive\\listini\\'+str(inizio_settimana.year)+'\\'+str(inizio_settimana.month))
esiste=os.path.exists('C:\\Users\\f.altarocca\\Google Drive\\listini\\'+str(inizio_settimana.year)+'\\'+str(inizio_settimana.month)+'\\'+str(inizio_settimana.day)+'-'+str(fine_settimana.day))
if not esiste:
    os.mkdir('C:\\Users\\f.altarocca\\Google Drive\\listini\\'+str(inizio_settimana.year)+'\\'+str(inizio_settimana.month)+'\\'+str(inizio_settimana.day)+'-'+str(fine_settimana.day))
"""

#shutil.copy('C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\NEG.PDF','C:\\Users\\f.altarocca\\Google Drive\\listini\\'+str(inizio_settimana.year)+'\\'+str(inizio_settimana.month)+'\\'+str(inizio_settimana.day)+'-'+str(fine_settimana.day)+'\\NEG.PDF')
#shutil.copy('C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\NEG2.PDF','C:\\Users\\f.altarocca\\Google Drive\\listini\\'+str(inizio_settimana.year)+'\\'+str(inizio_settimana.month)+'\\'+str(inizio_settimana.day)+'-'+str(fine_settimana.day)+'\\NEG2.PDF')

#copia csv
shutil.copy("C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\neg.csv","C:\\Users\\f.altarocca\\Desktop\\neg.csv")
shutil.copy("C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\gas.csv","C:\\Users\\f.altarocca\\Desktop\\gas.csv")
shutil.copy("C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\gr.csv","C:\\Users\\f.altarocca\\Desktop\\gr.csv")
shutil.copy("C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\neg2.csv","C:\\Users\\f.altarocca\\Desktop\\neg2.csv")
shutil.copy("C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\privati.csv","C:\\Users\\f.altarocca\\Desktop\\privati.csv")
shutil.copy("C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\pri20.csv","C:\\Users\\f.altarocca\\Desktop\\pri20.csv")
#copia csv

files=os.listdir('C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\')
for elemento in files:
    if elemento != 'neg.csv':
        shutil.move("C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\"+elemento,'C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\listini\\'+str(today.day)+'_'+str(today.month)+'_'+str(today.year)+'\\'+elemento)
    else:
        os.unlink("C:\\Users\\f.altarocca\\Desktop\\Scripts\\Gestione listini\\files\\"+elemento)

#execfile('C:\\Users\\f.altarocca\\Desktop\\Scripts\\Lavorazione Dati\\controllo magazzino\\controllo magazzino2.9.py')
#execfile('C:\\Users\\f.altarocca\\Desktop\\Scripts\\Lavorazione Dati\\controllo magazzino\\Non_aperti_2.1.py')
