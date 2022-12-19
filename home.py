import tkinter as tk
import os


def mensile():
    pass
    """
    if (var1.get() == 1):
        print(var1.get())
    """
def lancio_spedizione():
    #print("lanciata")
    f=open("C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//parametri.txt","w")
    f.write(str(var1.get()))
    f.close()
    exec(open("C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//lettura_spedizione.py").read())
    f=open('C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//log_invio.txt',"r")
    fileInvio=f.readlines()
    logInvio=fileInvio[0]
    f.close()
    if logInvio=='Mail inviate':
        l2 = tk.Label(window, bg='white', width=20, text='Listini Spediti', font =("Helvetica", 15))
        l2.pack()
    else:
        l2 = tk.Label(window, bg='white', width=20, text='Listini Non Spediti', font =("Helvetica", 15))
        l2.pack()

    os.unlink('C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//log_invio.txt')
    os.unlink("C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//parametri.txt")

f=open('C:\\Users\\f.altarocca\\Desktop\\Scripts\\lock\\lock.txt','r')
contenuto=f.readlines()
if contenuto[0].upper()=='ON':

    window = tk.Tk()
    window.title('Spedizione Listini')
    window.geometry('400x200')
     
    l = tk.Label(window, bg='white', width=20, text='Spedizione Listini', font =("Helvetica", 15))
    l.pack()

    b=tk.Button(text="Spedisci i listini", command=lancio_spedizione)
    b.pack()

    var1 = tk.IntVar()
    c1 = tk.Checkbutton(window, text='Listini Mensili',variable=var1, onvalue=1, offvalue=0, command=mensile)
    c1.pack()
    
    files=os.listdir("C://Users//f.altarocca//Documents//coding//Python//scripts38//ListiniAManoSender//Listini//")
    if files==[]:
        l3 = tk.Label(window, bg='white', width=30, text='Attenzione Listini Mancanti', font =("Helvetica", 15))
        l3.pack()
    else:
        l3 = tk.Label(window, bg='white', width=30, text='Listini Pronti', font =("Helvetica", 15))
        l3.pack()
        
    window.mainloop()