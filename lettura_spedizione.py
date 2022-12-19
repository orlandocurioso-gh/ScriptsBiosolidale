import spedisciMail
import os
import archiviazione

def mailInviata(risposte):
    f=open('C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//log_invio.txt',"w")
    if risposte==[]:
        f.write('Mail inviate')
    else:
        f.write('Mail Non inviate')
    f.close()

risposte=[]
soggetto="listino settimanale"
testo='In allegato il listino della prossima settimana'
#print ('---')
f=open("C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//parametri.txt","r")
flagmensile=f.readlines()
f.close()
files=os.listdir("C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//destinatari")
#print (files)
#print (flagmensile)
if flagmensile[0]=='0':
    files.pop(1)
    files.pop(1)
#print (files)
for elemento in files:
    listaPulitaDestinatari=[]
    f=open("C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//destinatari//"+elemento, 'r')
    riga=f.readlines()
    f.close()
    nomelistino=elemento.split('_')
    for item in riga:
        destinatari=item.split(';')
        for indirizzo in destinatari:
            if '\n' not in indirizzo:
                listaPulitaDestinatari.append(indirizzo)
            else:
                if len(indirizzo) != 1:
                    listaPulitaDestinatari.append(indirizzo[:-1])
    nomefile=nomelistino[0]+'.pdf'
    print (nomelistino[0])
    for indirizzo in listaPulitaDestinatari:
        print(indirizzo)
        risposta=spedisciMail.mail(indirizzo,soggetto,testo,nomefile)
        if risposta !={}:
            risposte.append(risposta)
mailInviata(risposte)
archiviazione.archiviazioneFiles()