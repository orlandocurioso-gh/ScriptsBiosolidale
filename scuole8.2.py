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
f=open('C:\\Users\\f.altarocca\\Desktop\\Scripts\\lock\\lock.txt','r')
contenuto=f.readlines()
if contenuto[0].upper()=='ON':
   
    listafilecreati=[]
    listafiledacancellare=[]
    exit='n'
    while exit=='n' or exit=='N':
        lista_giorni_print=[]
        totalenomisettimana=[]
        listafiledacancellare=[]
        listaappgiorni=[]
        listanumeriscuole=[]
        #dizionariototale=[]
        nomefileassociazioni =''
        listanomifile=[]
        listafilecreati=[]
        totalone=[]
        filestipo=[]
        contatore=1
        trasformati=''
        intestazione='CODICE ADHOC,SEDE,CENTRO COSTO,COD.PRODOTTO,PEZZI,DATA EVASIONE,ORARIO\n'
        totalone.append(intestazione)
        setogior=raw_input('settimanale (1) o giornaliero(2)? ')
        tipo=raw_input('Laziobio (1) o biosolidale (2)?')
        if tipo=='1':
            tipo='laziobio'
        elif tipo=='2':
            tipo='biosolidale'   
        #tipo='biosolidale'
        gruppo=raw_input('inserisci il gruppo: ')
        trasformati=raw_input('Trasformati?(s/n): ')
        trasformati=trasformati.upper()
        anno=raw_input('inserisci anno: ')
        mese=raw_input('inserisci mese: ')
        if setogior=='2':
            giorno=raw_input('inserisci giorno: ')   
        #ricerca files
        if setogior=='1':
            files=os.listdir("\\\\Dc02\\bioso-car\\carico scuole\\"+tipo+"\\"+anno+"\\"+mese+"\\"+'\\da fare')
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
    #pulizia lista files ---
            for elementofiles in files:
                isdir = os.path.isdir("\\\\DC02\\bioso-car\\carico scuole\\biosolidale\\2020\\01\\da fare\\"+elementofiles) 
                if isdir:
                    indiceelemento=files.index(elementofiles)
                    files.pop(indiceelemento)
    #pulizia lista files ---
    #creazione lista giorni per la stampa finale---
            for elementofiles in files:
                lista_nome_files=elementofiles.split('_')
                #print lista_nome_files
                lista_giorni_print.append(lista_nome_files[1])
    #creazione lista giorni per la stampa finale---      
            for elemento in files:
                riga=elemento.split('_')
                if gruppo.capitalize() in riga[0].capitalize():
                    filestipo.append(elemento)
        else:
            filestipo.append(gruppo.title()+"_"+giorno+"_"+mese+"_"+anno+'xls')
        #fine ricerca files
        i=0
        for elemento in filestipo:
            d={}
            if setogior =='1':
                giorno=elemento.split('_')[1]
            nomefile=gruppo+"_"+giorno+"_"+mese+"_"+anno
            data=giorno+"/"+mese+"/"+anno
            f1=open('C:\\Users\\f.altarocca\\Desktop\\Totale'+'_'+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'.csv','w')
            f1.write(intestazione)
            #funzione per ricerca codice e centro di costo
            if gruppo.upper()!='NG':
                f=open('\\\\Dc02\\bioso-car\\carico scuole\\utility\\codadhoc'+tipo+'.csv', 'r')
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
            esite=os.path.exists("\\\\Dc02\\bioso-car\\carico scuole\\"+tipo+"\\"+anno+"\\"+mese)
            if not esite:
                os.mkdir ("\\\\Dc02\\bioso-car\\carico scuole\\"+tipo+"\\"+anno+"\\"+mese) 

            esite=os.path.exists("\\\\Dc02\\bioso-car\\carico scuole\\"+tipo+"\\"+anno+"\\"+mese+"\\"+'\\da fare')
            if not esite:
                os.mkdir ("\\\\Dc02\\bioso-car\\carico scuole\\"+tipo+"\\"+anno+"\\"+mese+"\\"+'\\da fare')

            esite=os.path.exists('\\\\Dc02\\bioso-car\\carico scuole\\'+tipo+'\\'+anno+'\\'+mese+'\\csv\\')
            if not esite:
                os.mkdir ('\\\\Dc02\\bioso-car\\carico scuole\\'+tipo+'\\'+anno+'\\'+mese+'\\csv\\')

            esite=os.path.exists('\\\\Dc02\\bioso-car\\carico scuole\\'+tipo+'\\'+anno+'\\'+mese+'\\csv\\'+gruppo)
            if not esite:
                os.mkdir ('\\\\Dc02\\bioso-car\\carico scuole\\'+tipo+'\\'+anno+'\\'+mese+'\\csv\\'+gruppo)



            
            convertXLS2CSV("\\\\Dc02\\bioso-car\\carico scuole\\"+tipo+"\\"+anno+"\\"+mese+'\\da fare\\'+nomefile+'.xls', nomefile)
            f=open("C:\\Users\\f.altarocca\\Desktop\\"+nomefile+".csv",'r')
            ordine=f.readlines()
            f.close()

            #pulizia lista ordine
            
            for elemento in ordine:
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
                f=open('\\\\Dc02\\bioso-car\\carico scuole\\utility\\scuole_'+gruppo+'_'+tipo+'.csv')
                numeriscuole=f.readlines()
                f.close()
            for elemento in scuole:
                #ricerca scuole senza gruppo
                if gruppo.upper()=='NG':
                    f=open('\\\\Dc02\\bioso-car\\carico scuole\\utility\\codadhocng'+tipo+'.csv', 'r')
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
                        #print elemento
                        #print riga[1][:-1]
                        if elemento==riga[1][:-1]:
                            codicenumericoscuola=riga[0]
                            #print codicenumericoscuola
                #ricerca numero scuola fine
                            
                ordinefinale=[]
                indicescuola=scuole.index(elemento)
                i=2
                while i<=len(ordine)-1:
                    riga=ordine[i].split(';')
                    if riga[indicescuola+2]!='':
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
                        if (riga[1]=='BANES' or riga[1]=='BAN' or riga[1]=='BANFT') and (gruppo=='pedevilla' or gruppo=='PEDEVILLA' or gruppo=='pedevilla2' or gruppo=='PEDEVILLA2') and tipo=='biosolidale':
                            riga[indicescuola+2]=str(int(round(float(riga[indicescuola+2])*0.170)))
                        if (riga[1]=='OLEV5CS') and (gruppo=='pedevilla' or gruppo=='PEDEVILLA') and tipo=='biosolidale':
                            lattefinali=[]
                            numlitri=float(riga[indicescuola+2])
                            latte=numlitri/5
                            prima=str(latte)
                            testquantita=prima.split('.')
                            #print testquantita
                            if int(testquantita[1])== 0:
                                riga[indicescuola+2]=testquantita[0]
                            else:
                                if int(testquantita[1])> 6:
                                    riga[indicescuola+2]= str(int(testquantita[0])+1)
                                else:
                                    riga[indicescuola+2]= testquantita[0]
                            
                        if (riga[1]=='PREZCS' or riga[1]=='PREZ' or riga[1]=='BASICS') and (gruppo=='ELIOR' or gruppo=='elior'):                     
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
                    dizionarioordinamento={}
                    esite=os.path.exists("C:\\Users\\f.altarocca\\Desktop\\"+tipo+'_'+gruppo+'_'+giorno+'_'+mese+'_'+anno)
                    if not esite:
                        listafiledacancellare.append("C:\\Users\\f.altarocca\\Desktop\\"+tipo+'_'+gruppo+'_'+giorno+'_'+mese+'_'+anno)###lfc
                        os.mkdir ("C:\\Users\\f.altarocca\\Desktop\\"+tipo+'_'+gruppo+'_'+giorno+'_'+mese+'_'+anno)
                    f=open("C:\\Users\\f.altarocca\\Desktop\\"+tipo+'_'+gruppo+'_'+giorno+'_'+mese+'_'+anno+'\\'+str(contatore)+'_'+elemento+'_'+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'.csv','w')

                    nomiscuole.append(elemento)
                    j=0
                    f.write(intestazione)
                    while j<=len(ordinefinale)-2:
                        if trasformati=='N':
                            f.write(codice+','+codicenumericoscuola+','+centrocosto+','+ordinefinale[j]+','+ordinefinale[j+1]+','+data+'\n')
                            f1.write(codice+','+codicenumericoscuola+','+centrocosto+','+ordinefinale[j]+','+ordinefinale[j+1]+','+data+'\n')
                            totalone.append(codice+','+codicenumericoscuola+','+centrocosto+','+ordinefinale[j]+','+ordinefinale[j+1]+','+data+'\n')
                        else:
                            f.write(codice+','+codicenumericoscuola+','+centrocosto+','+ordinefinale[j]+','+ordinefinale[j+1]+','+data+',,TRASF\n')
                            f1.write(codice+','+codicenumericoscuola+','+centrocosto+','+ordinefinale[j]+','+ordinefinale[j+1]+','+data+',,TRASF\n')
                            totalone.append(codice+','+codicenumericoscuola+','+centrocosto+','+ordinefinale[j]+','+ordinefinale[j+1]+','+data+',,TRASF\n')
                        j=j+2
                    f.close()
                    d[codicenumericoscuola]=elemento
                    shutil.copy("C:\\Users\\f.altarocca\\Desktop\\"+tipo+'_'+gruppo+'_'+giorno+'_'+mese+'_'+anno+'\\'+str(contatore)+'_'+elemento+'_'+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'.csv','\\\\Dc02\\bioso-car\\carico scuole\\'+tipo+'\\'+anno+'\\'+mese+'\\csv\\'+gruppo+'\\'+str(contatore)+'_'+elemento+'_'+gruppo+'_'+tipo+'_'+'_'+giorno+'_'+mese+'_'+anno+'.csv')
                    contatore=contatore+1

            for elemento in nomiscuole:
                totalenomisettimana.append(elemento)
            totalenomisettimana.append('*')
            listaappgiorni.append(giorno)
            #fine creazione file corrispondenze
            #fine creazione ordini singole scuole     
            #print listafilecreati
            #print nomiscuole    
            f1.close()
            shutil.copy('\\\\Dc02\\bioso-car\\carico scuole\\'+tipo+'\\'+anno+'\\'+mese+'\\da fare\\'+gruppo+'_'+giorno+'_'+mese+'_'+anno+'.xls','\\\\Dc02\\bioso-car\\carico scuole\\'+tipo+'\\'+anno+'\\'+mese+'\\'+gruppo+'_'+giorno+'_'+mese+'_'+anno+'.xls')
            shutil.copy('C:\\Users\\f.altarocca\\Desktop\\Totale'+'_'+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'.csv','C:\\Users\\f.altarocca\\Desktop\\'+tipo+'_'+gruppo+'_'+giorno+'_'+mese+'_'+anno+'\\Totale'+'_'+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'.csv')
            os.unlink('C:\\Users\\f.altarocca\\Desktop\\Totale'+'_'+gruppo+'_'+tipo+'_'+giorno+'_'+mese+'_'+anno+'.csv')
            os.unlink("C:\\Users\\f.altarocca\\Desktop\\"+nomefile+".csv")
            os.unlink('\\\\Dc02\\bioso-car\\carico scuole\\'+tipo+'\\'+anno+'\\'+mese+'\\da fare\\'+gruppo+'_'+giorno+'_'+mese+'_'+anno+'.xls')
        #print listanomifile
        if setogior=='1':
            #####
            if (gruppo=='pedevilla' or gruppo=='PEDEVILLA')and (tipo=='biosolidale' or tipo=='BIOSOLIDALE'):
                totaloneAnto=[]
                totalenomisettimanaAnto=[]
                ford=open('\\\\Dc02\\bioso-car\\carico scuole\\utility\\scuole_pedevilla_biosolidale_proAnto.csv')
                ordinaAnto=ford.readlines()
                ford.close()
                for elemento in listaappgiorni:               
                    for school in ordinaAnto:
                        sede=school.split(';')[1]
                        entrato=0
                        for pezzo in totalone:
                            if pezzo.split(',')[1]==sede and pezzo.split(',')[5]==elemento+'/'+mese+'/'+anno+'\n':
                                totaloneAnto.append(pezzo)
                                entrato=1
                        if entrato==1:
                            totalenomisettimanaAnto.append(school.split(';')[2][:-1])
                    totalenomisettimanaAnto.append('*')
                #print totalenomisettimanaAnto
                f=open('C:\\Users\\f.altarocca\\Desktop\\Totale Settimana Antonello'+'_'+gruppo+'_'+tipo+'.csv','w')
                listafiledacancellare.append('C:\\Users\\f.altarocca\\Desktop\\Totale Settimana Antonello'+'_'+gruppo+'_'+tipo+'.csv')
                f.write('CODICE ADHOC,SEDE,CENTRO COSTO,COD.PRODOTTO,PEZZI,DATA EVASIONE\n')
                for elemento in totaloneAnto:
                    f.write(elemento)
                f.close()
            else:
                #crea totale settimana
                f=open('C:\\Users\\f.altarocca\\Desktop\\Totale Settimana'+'_'+gruppo+'_'+tipo+'.csv','w')
                for elemento in totalone:
                    #print elemento
                    f.write(elemento)
                f.close()
                #fine crea totale settimana
                listafiledacancellare.append('C:\\Users\\f.altarocca\\Desktop\\Totale Settimana'+'_'+gruppo+'_'+tipo+'.csv')  ###lfc     
        #print totalenomisettimana
        #print totalenomisettimana
        log=raw_input('inserisci file di log? ')
        if log=='s' or log=='S':
            #estrazione numeri da ORWEB
            f=open("C:\\Users\\f.altarocca\\Desktop\\ORDWEB.LOG",'r')
            file=f.readlines()
            f.close()
            file.pop(0)
            file.pop(0)
            file.pop(0)
            listaordiniadhoc=[]
            #print file
            for elemento in file:
                ord=elemento.split(':')[1].split('del')[0]
                ord2=ord.split(' ')
                listaordiniadhoc.append(ord2[1]+' '+ord2[2]+ord2[len(ord2)-2])
            #print listaordiniadhoc
            #fine estrazione numeri da ORWEB
            #if setogior=='1':
            esiste=os.path.exists('\\\\Dc02\\bioso-car\\carico scuole\\associazioni\\'+tipo+'\\'+gruppo+'\\')
            if not esiste:
                os.mkdir('\\\\Dc02\\bioso-car\\carico scuole\\associazioni\\'+tipo+'\\'+gruppo+'\\')
            esiste=os.path.exists('\\\\Dc02\\bioso-car\\carico scuole\\associazioni\\'+tipo+'\\'+gruppo+'\\'+anno)
            if not esiste:
                os.mkdir('\\\\Dc02\\bioso-car\\carico scuole\\associazioni\\'+tipo+'\\'+gruppo+'\\'+anno)
            esiste=os.path.exists('\\\\Dc02\\bioso-car\\carico scuole\\associazioni\\'+tipo+'\\'+gruppo+'\\'+anno+'\\'+mese)
            if not esiste:
                os.mkdir('\\\\Dc02\\bioso-car\\carico scuole\\associazioni\\'+tipo+'\\'+gruppo+'\\'+anno+'\\'+mese)
            #print listaappgiorni
            for elemento in listaappgiorni:
                esiste=os.path.exists('\\\\Dc02\\bioso-car\\carico scuole\\associazioni\\'+tipo+'\\'+gruppo+'\\'+anno+'\\'+mese+'\\'+elemento)
                #print esiste
                if not esiste:
                    os.mkdir('\\\\Dc02\\bioso-car\\carico scuole\\associazioni\\'+tipo+'\\'+gruppo+'\\'+anno+'\\'+mese+'\\'+elemento)
                else:
                    #print 'in'
                    random.seed()
                    numerocasuale=str(random.randint(0, 100))
                    os.rename('\\\\Dc02\\bioso-car\\carico scuole\\associazioni\\'+tipo+'\\'+gruppo+'\\'+anno+'\\'+mese+'\\'+elemento,'\\\\Dc02\\bioso-car\\carico scuole\\associazioni\\'+tipo+'\\'+gruppo+'\\'+anno+'\\'+mese+'\\'+elemento+'_'+numerocasuale)
                    os.mkdir('\\\\Dc02\\bioso-car\\carico scuole\\associazioni\\'+tipo+'\\'+gruppo+'\\'+anno+'\\'+mese+'\\'+elemento)
            i=0
            j=0
            f=open('\\\\Dc02\\bioso-car\\carico scuole\\associazioni\\'+tipo+'\\'+gruppo+'\\'+anno+'\\'+mese+'\\'+listaappgiorni[j]+'\\ordinamento.txt','w')

            if (gruppo=='pedevilla' or gruppo=='PEDEVILLA')and (tipo=='biosolidale' or tipo=='BIOSOLIDALE'):
                contatorescuole=0
                f.write('Lista ordini '  + tipo + ' ' + gruppo+' del '+ listaappgiorni[j]+'/'+mese+'/'+anno+'\n'+'\n')
                for elemento in totalenomisettimanaAnto:
                    if elemento !='*':
                        f.write(elemento + '--->' + listaordiniadhoc[i]+'\n\n\n')
                        print elemento + '--->' + listaordiniadhoc[i]
                        contatorescuole=contatorescuole+1
                        i=i+1
                    else:
                        f.write('\n')
                        f.write('Totale scuole: '+ str(contatorescuole))
                        contatorescuole=0
                        f.close()
                        if gruppo.capitalize()=='Elior' or gruppo.capitalize()=='Pedevilla' or gruppo.capitalize()=='Bioristororoma':
                            win32api.ShellExecute(0,"print",'\\\\Dc02\\bioso-car\\carico scuole\\associazioni\\'+tipo+'\\'+gruppo+'\\'+anno+'\\'+mese+'\\'+listaappgiorni[j]+'\\ordinamento.txt',None,".",0)
                        if len(listaappgiorni)>1 and j<len(listaappgiorni)-1:
                            j=j+1
                            f=open('\\\\Dc02\\bioso-car\\carico scuole\\associazioni\\'+tipo+'\\'+gruppo+'\\'+anno+'\\'+mese+'\\'+listaappgiorni[j]+'\\ordinamento.txt','w')
                            f.write('Lista ordini ' + tipo + ' ' + gruppo+' del '+ listaappgiorni[j]+'/'+mese+'/'+anno+'\n'+'\n')
            else:
                contatorescuole=0
                f.write('Lista ordini ' + tipo + ' ' + gruppo+' del '+ listaappgiorni[j]+'/'+mese+'/'+anno+'\n'+'\n')
                for elemento in totalenomisettimana:
                    if elemento !='*':
                        f.write(elemento + '--->' + listaordiniadhoc[i]+'\n'+'\n')
                        print elemento + '--->' + listaordiniadhoc[i]
                        contatorescuole=contatorescuole+1
                        i=i+1
                    else:
                        f.write('\n')
                        f.write('Totale scuole: '+ str(contatorescuole))
                        contatorescuole=0
                        f.close()
                        if gruppo.capitalize()=='Elior' or gruppo.capitalize()=='Pedevilla' or gruppo.capitalize()=='Bioristororoma':
                            win32api.ShellExecute(0,"print",'\\\\Dc02\\bioso-car\\carico scuole\\associazioni\\'+tipo+'\\'+gruppo+'\\'+anno+'\\'+mese+'\\'+listaappgiorni[j]+'\\ordinamento.txt',None,".",0)
                        if len(listaappgiorni)>1 and j<len(listaappgiorni)-1:
                            j=j+1
                            f=open('\\\\Dc02\\bioso-car\\carico scuole\\associazioni\\'+tipo+'\\'+gruppo+'\\'+anno+'\\'+mese+'\\'+listaappgiorni[j]+'\\ordinamento.txt','w')
                            f.write('Lista ordini '  + tipo + ' ' + gruppo+' del '+ listaappgiorni[j]+'/'+mese+'/'+anno+'\n'+'\n')
             #fine genera file totale associazione della settimana

        def stampaassociazioni(nomefileassociazioni,gruppo,tipo):
            fname = nomefileassociazioni
            #print '-' * 40                          # Genera 40 trattini
            F1 = open(fname,'r')
            data = F1.read()
            print gruppo+' - '+tipo+' '+giorno+'_'+mese+'_'+anno+'\n'
            print data
        if nomefileassociazioni !='':
            stampaassociazioni(nomefileassociazioni,gruppo,tipo)
        """
        if gruppo.capitalize()=='Pedevilla':
            creaexcel=raw_input('creare excel? ')
            if creaexcel.capitalize()=='S':
                execfile('C:\\Users\\f.altarocca\\Desktop\\Scripts\\testXL\\testexcel2.py',{'filestipo':filestipo})
        """
        exit=raw_input('uscire(s/n)?')

    #--- pulizia desktop---

        for elemento in listafiledacancellare:
            print elemento
            if 'biosolidale_' in elemento:
                shutil.rmtree(elemento)     
            else:
                os.unlink(elemento)
        os.unlink('C:\\Users\\f.altarocca\\Desktop\\ORDWEB.LOG')

    #--- pulizia desktop---

    files=os.listdir('C:\\Users\\f.altarocca\\Desktop\\')

    #print listafiledacancellare


    ###---STAMPA ASSOCIAZIONI---###
    if gruppo.capitalize()!='Pedevilla':
        for elemento_giorno in lista_giorni_print:
            pathAssociazioni='\\\\Dc02\\bioso-car\\carico scuole\\associazioni\\'+tipo+'\\'+gruppo+'\\'+anno+'\\'+mese+'\\'+elemento_giorno
            filedastampare=pathAssociazioni+'\\'+'ordinamento.txt'
            win32api.ShellExecute(0,"print",filedastampare,None,".",0)
