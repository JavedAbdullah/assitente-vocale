from tkinter import *
import time
import speech_recognition as sr
import risposte
import threading as t
import win32com.client as wincl

### funzione per riproduce vocalmente il testo e viene attivata tramite un thread
def speak(testo):
    speaker = wincl.Dispatch("SAPI.SpVoice")
    speaker.speak(testo)

##### creazione della schermata
root = Tk() # creazione della grafica tkinter
root.configure(bg="black")  # impostazione dello sfondo nero
root.resizable(False, False)    # non è possibile ingrandire la schermata
root.title("Elsa Jean") # impostazione del titolo della finestra
#####

window = Frame(root, bg="black")    # creazione di un frame all'interno del root grafico
window.pack()   # inserimento della finestra nel root grafico

S = Scrollbar(window)   # creazione della barra di scorrimento all'interno della finestra
T = Text(window, bg="black", fg="white", font=("bold", 12)) # creazione della widget di testo contenente il testo, tale widget avrà lo sfondo nero, il font "bold" con grandezza 12 e colore bianco
S.pack(side=RIGHT, fill="both") # inserisce la barra di scorrimento a sinistra
T.pack(side=LEFT, fill="both", expand=1)    # inserisce il widget di testo a destra

### aggiunge la funzionalità di scorrimento al widget testuale
S.config(command=T.yview)
T.config(yscrollcommand=S.set)
###

#### set di tag con impostaazione dei colori per la grafica testuale
T.tag_config("yellow", foreground="yellow")
T.tag_config("green", foreground="green")
T.tag_config("red", foreground="red")
T.tag_config("orange", foreground="orange")
####

######## impostazioni riconoscimento vocale
r = sr.Recognizer() # creazione del riconoscitore vocale
r.energy_threshold = 4000  # impostazione del rumore di sottofondo
r.dynamic_energy_threshold = True  # volume di sottofondo varia in base al rumore di sottofondo riconosciuto
source = sr.Microphone()    # imposta il microfono come sorgente di registrazione dell'audio
########

# funzione che inserirà nel widget testuale ciò che verrà pronunciato dall'utente
def messUtente(testo): 
    T.insert(INSERT, "\nIo: ", "yellow")    # inserisce nel widget il testo passato alla funzione
    T.insert(INSERT, testo+"\n")    # inserisce nel widget il testo passato alla funzione

# funzione che inserirà nel widget testuale ciò che verrà pronunciato dall'assistente vocale
def messElsa(testo):
    T.insert(INSERT, "\nElsa: ", "green")   # inserisce nel widget il testo passato alla funzione
    T.insert(INSERT, testo+"\n")    # inserisce nel widget il testo passato alla funzione


def callback(r, audio):
    try:
        testo = r.recognize_google(audio, language="it-IT").lower() # riconoscimento della voce tramite riconoscimento vocale google
          
        if "elsa" in testo: # se il messaggio contiene elsa, verra avviato il processo di risposta tramite risposte()
            messUtente(testo[testo.index("elsa"):]) # il mess dell'utente va da elsa in poi

            risposta = risposte.rispondi(testo[testo.index("elsa")+5:]) # riceve la risposta da risposte.py
            messElsa(risposta)  # aggiunge un testo da parte dell'assistente al widget di testo

        else:
            messUtente(testo)   # aggiunge un testo da parte dell'utente al widget di testo
            
    except sr.UnknownValueError:    # errore nel caso di un riconoscimento vocale errato
        print("errore")

def elsa():
    #print("Elsa Jean in ascolto")

    T.insert(INSERT, "Elsa Jean in ascolto\n", "orange")
    r.listen_in_background(source, callback)    # ascolto in background di ciò che viene detto dall'utente e richiamo della funzione "callback()"

    time.sleep(10000)   # il programma continuerà ad andare per 10k secondi ossia circa 3 ore
    

def inizia():
    x = t.Thread(target=elsa) # fa partire l'assistente vocale in back-end tramite thread
    x.start()
        
### prima funzione che verrà richiamata dal programma
inizia()

### creazione della finestra grafica impostata precedentemente
root.mainloop()
