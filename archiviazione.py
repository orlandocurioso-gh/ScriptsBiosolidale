import os
import time
import shutil

def archiviazioneFiles():

    data=time.strftime("%d/%m/%Y")
    dataesplosa=data.split('/')
    giorno=dataesplosa[0]
    mese=dataesplosa[1]
    anno=dataesplosa[2]

    esite=os.path.exists('C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//spediti//'+anno+'//')
    if not esite:
        os.mkdir ('C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//spediti//'+anno+'//')
    esite=os.path.exists('C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//spediti//'+anno+'//'+mese+'//')
    if not esite:
        os.mkdir ('C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//spediti//'+anno+'//'+mese+'//')
    esite=os.path.exists('C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//spediti//'+anno+'//'+mese+'//'+giorno+'//')
    if not esite:
        os.mkdir ('C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//spediti//'+anno+'//'+mese+'//'+giorno+'//')
    target='C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//spediti//'+anno+'//'+mese+'//'+giorno+'//'

    files=os.listdir("C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//Listini//")
    source="C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//Listini//"
    for elemento in files:
        shutil.move(source+elemento,target+elemento)