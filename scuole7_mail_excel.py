# -*- coding: utf-8 -*-
import xlrd
import xlwt
import os
import time
import datetime
import calendar
import sys
import win32api
import shutil
import string
import smtplib
import time
import datetime
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email import Encoders
import os

def cercagiornisettimana(listagiornisettimana):
    nomegiorno=[]
    for elemento in listagiornisettimana:
        app=elemento.split('_')
        #print app
        anno=app[3][:4]
        mese=app[2]
        giorno=app[1]
        #print anno
        #print mese
        #print giorno
        wd=calendar.weekday(int(anno), int(mese), int(giorno))
        #print wd
        if wd==0:
            nomegiorno.append("LUNEDI")
        elif wd==1:
            nomegiorno.append("MARTEDI")
        elif wd==2:
            nomegiorno.append("MERCOLEDI")
        elif wd==3:
            nomegiorno.append("GIOVEDI")
        elif wd==4:
            nomegiorno.append("VENERDI")
    #print nomegiorno 
    return nomegiorno



gmail_user = "gas@biosolidale.it"
gmail_pwd = "BIO.Gas.10"

def mail(to, subject, text):
    msg = MIMEMultipart()
    msg['From'] = gmail_user
    msg['To'] = to
    msg['Subject'] = subject
    msg.attach(MIMEText(text))
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

def scriviexcel(fileinput):
    listaperexcel=[]
    #print type(document) # report giornaliero
    listapp=fileinput.split(':')
    #print listapp
    for elemento in listapp:
        if '\n' in elemento:
            splitta=elemento.split('\n')
            listaperexcel.append(splitta[0][1:])
            listaperexcel.append(splitta[1])
        else:
            listaperexcel.append(elemento)
    listaperexcel.pop(len(listaperexcel)-1)
    #print listaperexcel
    return listaperexcel




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
#### fine stile foglio excel


#trovare settimana di riferimento per nome file

MyFile = xlwt.Workbook()
MyFoglio = MyFile.add_sheet("SCUOLE_", cell_overwrite_ok = True)
MyFoglio.col(0).width = 6500
"""
MyFoglio.col(5).width = 6500
MyFoglio.col(4).width = 6500
MyFoglio.col(3).width = 6500
MyFoglio.col(2).width = 6500
MyFoglio.col(1).width = 6500
MyFoglio.col(0).width = 6500
MyFoglio.write(0,1,'LUNEDI',style=style)
MyFoglio.write(0,2,'MARTEDI',style=style)
MyFoglio.write(0,3,'MERCOLEDI',style=style)
MyFoglio.write(0,4,'GIOVEDI',style=style)
MyFoglio.write(0,5,'TOTALI',style=style)
"""


f=open('C:\\Users\\f.altarocca\\Desktop\\Scripts\\lock\\lock.txt','r')
contenuto=f.readlines()
if contenuto[0].upper()=='ON':
    exit='n'
    """
    listafilecreati=[]
    listagiorni=[]
    listapp=[]
    dizionario={}
    contagiorno=1
    listagiorniexcel=''
    """

    while exit=='n' or exit=='N':
        MyFile = xlwt.Workbook()
        MyFoglio = MyFile.add_sheet("SCUOLE_", cell_overwrite_ok = True)
        MyFoglio.col(0).width = 6500
        datixexcel=[]
        datixexceltotali=[]
        listafilecreati=[]
        listagiorni=[]
        listapp=[]
        dizionario={}
        contagiorno=1
        listagiorniexcel=''
        listaappgiorni=[]
        listanumeriscuole=[]
        dizionariototale=[]
        nomefileassociazioni =''
        listanomifile=[]
        listafilecreati=[]
        totalone=[]
        totaloneprodotti=[]
        filestipo=[]
        contatore=1
        storicochiavi=[]
        intestazione='CODICE ADHOC,SEDE,CENTRO COSTO,COD.PRODOTTO,PEZZI,DATA EVASIONE\n'
        totalone.append(intestazione)
        setogior=raw_input('settimanale (1) o giornaliero(2)? ')
        tipo=raw_input('Laziobio (1) o biosolidale (2)?')
        if tipo=='1':
            tipo='laziobio'
        elif tipo=='2':
            tipo='biosolidale'   
        #tipo='biosolidale'
        gruppo=raw_input('inserisci il gruppo: ')
        anno=raw_input('inserisci anno: ')
        mese=raw_input('inserisci mese: ')
        if setogior=='2':
            giorno=raw_input('inserisci giorno: ')   
        #ricerca files
        if setogior=='1':
            files=os.listdir("\\\\DC02\\bioso-car\\carico scuole\\"+tipo+"\\"+anno+"\\"+mese+"\\"+'\\da fare')
            #print files
            ###
            i=0
            while i<=len(files)-1:
                esplosione=files[i].split('_')
                if esplosione[0]!=gruppo:
                    indicelemento=files.index(files[i])
                    files.pop(indicelemento)
                else:
                    i=i+1
            ###
            for elementofiles in files:
                isdir = os.path.isdir("\\\\DC02\\bioso-car\\carico scuole\\biosolidale\\2020\\01\\da fare\\"+elementofiles) 
                if isdir:
                    indiceelemento=files.index(elementofiles)
                    files.pop(indiceelemento)
            #print files
            nomi_giorni_settimana=cercagiornisettimana(files) ###x
            index_col=1
            #print nomi_giorni_settimana
            for giorno in nomi_giorni_settimana:
                MyFoglio.write(0,index_col,giorno,style=style)
                MyFoglio.col(index_col).width = 6500
                index_col=index_col+1
            MyFoglio.write(0,index_col,'TOTALI',style=style)
            MyFoglio.col(index_col).width = 6500
            for elemento in files:
                riga=elemento.split('_')
                if gruppo.capitalize() in riga[0].capitalize():
                    filestipo.append(elemento)
        else:
            filestipo.append(gruppo.title()+"_"+giorno+"_"+mese+"_"+anno+'xls')
        #fine ricerca files
        i=0
        for elemento in filestipo:
            totalegiornosingolo=[]
            d={}
            if setogior =='1':
                giorno=elemento.split('_')[1]
            listagiorniexcel=listagiorniexcel+giorno+'_'
            nomefile=gruppo+"_"+giorno+"_"+mese+"_"+anno
            data=giorno+"/"+mese+"/"+anno
            listagiorni.append(giorno)
            f1=open('C:\\Users\\f.altarocca\\Desktop\\Totale'+'_'+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'.csv','w')
            f1.write(intestazione)
            #funzione per ricerca codice e centro di costo
            if gruppo.upper()!='NG':
                f=open('\\\\DC02\\bioso-car\\carico scuole\\utility\\codadhoc'+tipo+'.csv', 'r')
                codici=f.readlines()
                f.close()
                codice=''
                centrocosto=''
                for elemento in codici:
                    riga=elemento.split(';')
                    riga[2]=riga[2][:-1]
                    if riga[0]==gruppo.upper():
                       codice=riga[1]
                       centrocosto=riga[2]
            
            #fine funzione per ricerca codice e centro di costo
            esite=os.path.exists("\\\\DC02\\bioso-car\\carico scuole\\"+tipo+"\\"+anno+"\\"+mese)
            if not esite:
                os.mkdir ("\\\\DC02\\bioso-car\\carico scuole\\"+tipo+"\\"+anno+"\\"+mese) 

            esite=os.path.exists("\\\\DC02\\bioso-car\\carico scuole\\"+tipo+"\\"+anno+"\\"+mese+"\\"+'\\da fare')
            if not esite:
                os.mkdir ("\\\\DC02\\bioso-car\\carico scuole\\"+tipo+"\\"+anno+"\\"+mese+"\\"+'\\da fare')

            esite=os.path.exists('\\\\DC02\\bioso-car\\carico scuole\\'+tipo+'\\'+anno+'\\'+mese+'\\csv\\')
            if not esite:
                os.mkdir ('\\\\DC02\\bioso-car\\carico scuole\\'+tipo+'\\'+anno+'\\'+mese+'\\csv\\')

            esite=os.path.exists('\\\\DC02\\bioso-car\\carico scuole\\'+tipo+'\\'+anno+'\\'+mese+'\\csv\\'+gruppo)
            if not esite:
                os.mkdir ('\\\\DC02\\bioso-car\\carico scuole\\'+tipo+'\\'+anno+'\\'+mese+'\\csv\\'+gruppo)


            
            convertXLS2CSV("\\\\DC02\\bioso-car\\carico scuole\\"+tipo+"\\"+anno+"\\"+mese+'\\da fare\\'+nomefile+'.xls', nomefile)
            f=open("C:\\Users\\f.altarocca\\Desktop\\"+nomefile+".csv",'r')
            ordine=f.readlines()
            f.close()

            #pulizia lista ordine
            #print ordine
            for elemento in ordine:
                #print elemento
                ind=ordine.index(elemento)
                ordine[ind]=elemento[:-1]
            #print ordine

            #fine pulizia lista ordine
            scuole=ordine[1].split(';')
            scuole.pop(0)
            scuole.pop(0)
            #print scuole
            #creazione ordini singole scuole
            nomiscuole=[]
            contatore=1
            if gruppo.upper()!='NG':
                f=open('\\\\DC02\\bioso-car\\carico scuole\\utility\\scuole_'+gruppo+'_'+tipo+'.csv')
                numeriscuole=f.readlines()
                f.close()
            for elemento in scuole:
                #ricerca scuole senza gruppo
                if gruppo.upper()=='NG':
                    f=open('\\\\DC02\\bioso-car\\carico scuole\\utility\\codadhocng'+tipo+'.csv', 'r')
                    codici=f.readlines()
                    f.close()
                    codice=''
                    centrocosto=''
                    for cod in codici:
                        riga=cod.split(';')
                        riga[2]=riga[2][:-1]
                        if riga[0]==elemento.upper():
                           codice=riga[1]
                           #per ora: nel caso di un non gruppo, inserisce il numero scuola ugale al numero del gruppo
                           codicenumericoscuola='' #riga[1]
                           centrocosto=riga[2]
                    
                #ricerca scuole senza gruppo

                #ricerca numero scuola
                #print numeriscuole
                if gruppo.upper()!='NG':
                    for numero in numeriscuole:
                        riga=numero.split(';')
                        #print riga[1][:-1]
                        if elemento==riga[1][:-1]:
                            codicenumericoscuola=riga[0]
                            #print codicenumericoscuola
                #ricerca numero scuola fine
                #digits=string.digits ###           
                ordinefinale=[]
                indicescuola=scuole.index(elemento)
                i=2
                while i<=len(ordine)-1:
                    riga=ordine[i].split(';')
                    #print riga
                    #print riga [indicescuola+2]
                    if riga[indicescuola+2]!='' and riga[indicescuola+2].isalnum(): ###
                        #print riga [indicescuola+2]
                        ordinefinale.append(riga[1])#codice
                        if ',' in riga[indicescuola+2]:
                            numero=riga[indicescuola+2]
                            nuovo=''
                            for car in numero:
                                if car ==',':
                                    nuovo=nuovo+'.'
                                else:
                                    nuovo=nuovo+car
                            riga[indicescuola+2]=nuovo
                        if (riga[1]=='BANES' or riga[1]=='BAN' or riga[1]=='BANFT') and (gruppo=='pedevilla' or gruppo=='PEDEVILLA' or gruppo=='pedevilla2' or gruppo=='PEDEVILLA2'or gruppo=='bioristororoma' or gruppo=='BIORISTOROROMA') and tipo=='biosolidale':
                            riga[indicescuola+2]=str(int(round(float(riga[indicescuola+2])*0.17)))
                        if (riga[1]=='PREZCS' or riga[1]=='PREZ') and (gruppo=='ELIOR' or gruppo=='elior'):                     
                            numeropezzi=riga[indicescuola+2].split('.')
                            if '.' in riga[indicescuola+2]:
                                riga[indicescuola+2]=numeropezzi[1]
                            else:
                                riga[indicescuola+2]=numeropezzi[0]
                        if (riga[1]=='BASICS') and (gruppo=='ELIOR' or gruppo=='elior'):                     
                            numeropezzi=riga[indicescuola+2].split('.')
                            if '.' in riga[indicescuola+2]:
                                riga[indicescuola+2]=numeropezzi[1]
                            else:
                                riga[indicescuola+2]=numeropezzi[0]
                        if (riga[1]=='ROSM') and (gruppo=='ELIOR' or gruppo=='elior'):
                            numeropezzi=riga[indicescuola+2].split('.')
                            riga[indicescuola+2]=numeropezzi[1]
                        """
                        elif (riga[1]=='ARNA8' or riga[1]=='ARNA4') and (gruppo=='BIORISTORO' or gruppo=='bioristoro'):
                            riga[indicescuola+2]=str(int(round(float(riga[indicescuola+2])*0.150)))
                        """
                        ordinefinale.append(riga[indicescuola+2])#quantita
                    i=i+1
                #print ordinefinale
                if len(ordinefinale)>1:
                    #esite=os.path.exists("C:\\Users\\f.altarocca\\Desktop\\"+tipo+'_'+gruppo+'_'+giorno+'_'+mese+'_'+anno)
                    #if not esite:
                    #    os.mkdir ("C:\\Users\\f.altarocca\\Desktop\\"+tipo+'_'+gruppo+'_'+giorno+'_'+mese+'_'+anno)
                    #f=open("C:\\Users\\f.altarocca\\Desktop\\"+tipo+'_'+gruppo+'_'+giorno+'_'+mese+'_'+anno+'\\'+str(contatore)+'_'+elemento+'_'+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'.csv','w')
                    nomiscuole.append(elemento)
                    j=0
                    #f.write(intestazione)
                    while j<=len(ordinefinale)-2:
                        #f.write(codice+','+codicenumericoscuola+','+centrocosto+','+ordinefinale[j]+','+ordinefinale[j+1]+','+data+'\n')
                        #f1.write(codice+','+codicenumericoscuola+','+centrocosto+','+ordinefinale[j]+','+ordinefinale[j+1]+','+data+'\n')
                        #totalone.append(codice+','+codicenumericoscuola+','+centrocosto+','+ordinefinale[j]+','+ordinefinale[j+1]+','+data+'\n')
                        if ordinefinale[j] in totaloneprodotti:
                            #print ordinefinale[j]
                            ind=totaloneprodotti.index(ordinefinale[j])
                            quantita=float(totaloneprodotti[ind+1])
                            #print totaloneprodotti[ind+1]
                            totaloneprodotti[ind+1]=str(quantita+float(ordinefinale[j+1]))
                        else:
                            totaloneprodotti.append(ordinefinale[j])
                            totaloneprodotti.append(ordinefinale[j+1])
                        if ordinefinale[j] in totalegiornosingolo:
                            indgs=totalegiornosingolo.index(ordinefinale[j])
                            quantita=float(totalegiornosingolo[indgs+1])
                            totalegiornosingolo[indgs+1]=str(quantita+float(ordinefinale[j+1]))
                        else:
                            totalegiornosingolo.append(ordinefinale[j])
                            totalegiornosingolo.append(ordinefinale[j+1])
                        j=j+2
                    #f.close()
                    
                    d[codicenumericoscuola]=elemento
                    #shutil.copy("C:\\Users\\f.altarocca\\Desktop\\"+tipo+'_'+gruppo+'_'+giorno+'_'+mese+'_'+anno+'\\'+str(contatore)+'_'+elemento+'_'+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'.csv','\\\\DC02\\bioso-car\\carico scuole\\'+tipo+'\\'+anno+'\\'+mese+'\\csv\\'+gruppo+'\\'+str(contatore)+'_'+elemento+'_'+gruppo+'_'+tipo+'_'+'_'+giorno+'_'+mese+'_'+anno+'.csv')
                    contatore=contatore+1
            #print d

            #fine creazione file corrispondenze
            #fine creazione ordini singole scuole     
                
            f1.close()
            #shutil.copy('\\\\DC02\\bioso-car\\carico scuole\\'+tipo+'\\'+anno+'\\'+mese+'\\da fare\\'+gruppo+'_'+giorno+'_'+mese+'_'+anno+'.xls','\\\\DC02\\bioso-car\\carico scuole\\'+tipo+'\\'+anno+'\\'+mese+'\\'+gruppo+'_'+giorno+'_'+mese+'_'+anno+'.xls')
            #shutil.copy('C:\\Users\\f.altarocca\\Desktop\\Totale'+'_'+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'.csv','C:\\Users\\f.altarocca\\Desktop\\'+tipo+'_'+gruppo+'_'+giorno+'_'+mese+'_'+anno+'\\Totale'+'_'+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'.csv')
            os.unlink('C:\\Users\\f.altarocca\\Desktop\\Totale'+'_'+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'.csv')
            os.unlink("C:\\Users\\f.altarocca\\Desktop\\"+nomefile+".csv")
            #os.unlink('\\\\DC02\\bioso-car\\carico scuole\\'+tipo+'\\'+anno+'\\'+mese+'\\da fare\\'+gruppo+'_'+giorno+'_'+mese+'_'+anno+'.xls')
        #print listanomifile
            if totalegiornosingolo: #manda la mail per ogni giorno
                soggetto='Calcolo kg prodotti per scuole '+gruppo+' '+giorno+' '+mese+' '+anno
                f=open('C:\\Users\\f.altarocca\\Desktop\\Totale '+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'_parziale.txt','w')
                i=0
                while i<=len(totalegiornosingolo)-2:
                    f.write(totalegiornosingolo[i]+': '+totalegiornosingolo[i+1]+'\n')
                    i=i+2
                f.close() 
                file=open('C:\\Users\\f.altarocca\\Desktop\\Totale '+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'_parziale.txt','r')
                document = file.read()
                file.close()
                datixexcel=scriviexcel(document)
                #print datixexcel
                ########creazione dizionario
                i=0
                while i<= len(datixexcel)-1:
                    chiave=datixexcel[i]
                    if dizionario.has_key(chiave):
                        listapp=dizionario[chiave]
                        listapp.append(datixexcel[i+1])
                        dizionario[chiave]=listapp
                    else:
                        listapp=[]
                        listapp.append(datixexcel[i+1])
                        dizionario[chiave]=listapp
                    listapp=[]
                    i=i+2
                listadiappoggio=[]
                """
                print contagiorno
                print 'pre'
                for chiave, valore in dizionario.items():
                    print chiave, valore
                """
                #print dizionario
                for chiave in dizionario:
                    #print contagiorno
                    #print chiave
                    #print storicochiavi
                    listadiappoggio=dizionario[chiave]
                    #print listadiappoggio
                    if len(listadiappoggio)<contagiorno:
                        if chiave not in storicochiavi:
                            j=1
                            while j<contagiorno:
                                listadiappoggio.insert(0,'')
                                j=j+1
                        else:
                            listadiappoggio.append('')
                            dizionario[chiave]=listadiappoggio
                    if chiave not in storicochiavi:
                        storicochiavi.append(chiave)
                contagiorno=contagiorno+1
                #print storicochiavi
                """
                print 'post'
                for chiave, valore in dizionario.items():
                    print chiave, valore
                """
                ########creazione dizionario
                
                #mail("felice@biosolidale.it",soggetto,document)
                #mail("acquisti@biosolidale.it",soggetto,document)
                os.unlink('C:\\Users\\f.altarocca\\Desktop\\Totale '+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'_parziale.txt')
        #if setogior=='1':
        #    print totaloneprodotti
        #spedizione mail
        f=open('C:\\Users\\f.altarocca\\Desktop\\Totale '+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'.txt','w')
        i=0
        while i<=len(totaloneprodotti)-2:
            f.write(totaloneprodotti[i]+': '+totaloneprodotti[i+1]+'\n')
            i=i+2
        f.close()   
        import smtplib
        import time
        import datetime
        from email.MIMEMultipart import MIMEMultipart
        from email.MIMEBase import MIMEBase
        from email.MIMEText import MIMEText
        from email import Encoders
        import os
        gmail_user = "gas@biosolidale.it"
        gmail_pwd = "BIO.Gas.10"
        def mail(to, subject, text):
            msg = MIMEMultipart()
            msg['From'] = gmail_user
            msg['To'] = to
            msg['Subject'] = subject
            msg.attach(MIMEText(text))
            mailServer = smtplib.SMTP("smtp.gmail.com", 587)
            mailServer.ehlo()
            mailServer.starttls()
            mailServer.ehlo()
            mailServer.login(gmail_user, gmail_pwd)
            mailServer.sendmail(gmail_user, to, msg.as_string())
            mailServer.close()
        #print datafilereale
        file=open('C:\\Users\\f.altarocca\\Desktop\\Totale '+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'.txt','r')
        document = file.read()
        file.close()
        datixexceltotali=scriviexcel(document)
        #print datixexceltotali
        i=0
        while i<=len(datixexceltotali)-1:
            chiave=datixexceltotali[i]
            listadiappoggiofinale=dizionario[chiave]
            listadiappoggiofinale.append(datixexceltotali[i+1])
            dizionario[chiave]=listadiappoggiofinale
            i=i+2

        #print document # report totale settimana
        
        soggetto='Calcolo kg prodotti per scuole '+gruppo+' - settimana dal '+listagiorni[0]+' al '+listagiorni[len(listagiorni)-1]+' '+mese+' '+anno
        #mail("qualita@biosolidale.it",soggetto,document)
        #mail("acquisti@biosolidale.it",soggetto,document)
        #mail("felice@biosolidale.it",soggetto,document)
        
        os.unlink('C:\\Users\\f.altarocca\\Desktop\\Totale '+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'.txt')
        #fine spedizione mail
        #print dizionario
        rig=1
        for chiave, valore in dizionario.items():
            col=0
            MyFoglio.write(rig,col,chiave,style=style)
            for elemento in valore:
                #print type(elemento)
                if '.' in elemento:
                    appEle=elemento.split('.')
                    elemento=appEle[0]
                MyFoglio.write(rig,col+1,elemento,style=style)
                col=col+1
            rig=rig+1
            #print chiave, valore
            
        esite=os.path.exists("\\\\DC02\\bioso-car\\ACQUISTI\\Scuole"+"\\"+anno)
        if not esite:
            os.mkdir ("\\\\DC02\\bioso-car\\ACQUISTI\\Scuole"+"\\"+anno)
        esite=os.path.exists("\\\\DC02\\bioso-car\\ACQUISTI\\Scuole"+"\\"+anno+"\\"+mese)
        if not esite:
            os.mkdir ("\\\\DC02\\bioso-car\\ACQUISTI\\Scuole"+"\\"+anno+"\\"+mese)
            
        MyFile.save(r"\\DC02\bioso-car\ACQUISTI\Scuole\\"+anno+"\\"+mese+"\\"+gruppo+"_"+listagiorniexcel+"_"+mese+"_"+anno+".xls")


        try:
            soggetto='File excel prospetti scuole pronto'
            document='Il file excel dei prospetti delle scuole per i giorni: '+listagiorniexcel+"//"+mese+"//"+anno+' e\' disponibile qui:\n\n \\DC02\\bioso-car\ACQUISTI\Scuole\\'+anno+"\\"+mese+'\\'
            mail("acquisti@biosolidale.it",soggetto,document)
            mail("vendite@biosolidale.it",soggetto,document)
            mail("felice@biosolidale.it",soggetto,document)
        except:
            execfile('C:\\Users\\f.altarocca\\Desktop\\Scripts\\GUI\\tkinter\\alertNOMAIL.py')
          
        exit=raw_input('uscire(s/n)?')