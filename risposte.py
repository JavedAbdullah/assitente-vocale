import urllib.request, json, webbrowser, os, sys, requests
from json.decoder import JSONDecodeError
import win32com.client as wincl
from googletrans import Translator
import threading as t

### funzione per riproduce vocalmente il testo e viene attivata tramite un thread
def speak(testo):
    speaker = wincl.Dispatch("SAPI.SpVoice")
    speaker.speak(testo)

### funzione per tradurre da italiano alla lingua di destinazione
def traduciDaIta(testoDaTradurre, linguaDestinazione):
    risposta =""
    trans = Translator()

    if linguaDestinazione == "inglese":
        traduzione = trans.translate(testoDaTradurre, src="it", dest="en")
        risposta += traduzione.text
    
    elif linguaDestinazione == "tedesco":
        traduzione = trans.translate(testoDaTradurre, src="it", dest="de")
        risposta += traduzione.text

    elif linguaDestinazione == "francese":
        traduzione = trans.translate(testoDaTradurre, src="it", dest="fr")
        risposta += traduzione.text

    elif linguaDestinazione == "portoghese":
        traduzione = trans.translate(testoDaTradurre, src="it", dest="pt")
        risposta += traduzione.text

    elif linguaDestinazione == "spagnolo":
        traduzione = trans.translate(testoDaTradurre, src="it", dest="es")
        risposta += traduzione.text

    return risposta

### funzione per tradurre in italiano da una lingua riconosciuta automaticamente
def traduciInIta(testoDaTradurre):
    risposta = ""
    trans = Translator()

    traduzione = trans.translate(testoDaTradurre, dest="it")
    risposta+= traduzione.text
    
    return risposta

### gestione della risposta in base alla richiesta dell'utente
def rispondi(testo):

    risposta = "" # variabile risposta che verrà restituita al main per essere stampata a schermo

    parole = testo.split(" ")   # divide il testo di tipo stringa in una lista formata da n parole quanti sono gli spazi

    # richiesta di aprire il sito di spotify
    if "apri spotify" in testo:
        risposta += "apro spotify"

        x = t.Thread(target=speak, args=(risposta,)) # thread che farà partire la funzione speak la quale pronuncierà la risposta
        x.start()

        webbrowser.open("https://open.spotify.com/")

    elif "apri google" in testo:
        risposta += "apro google"

        x = t.Thread(target=speak, args=(risposta,)) # thread che farà partire la funzione speak la quale pronuncierà la risposta
        x.start()

        webbrowser.open("https://www.google.it/")

    elif "apri youtube" in testo:
        risposta += "apro youtube"

        x = t.Thread(target=speak, args=(risposta,)) # thread che farà partire la funzione speak la quale pronuncierà la risposta
        x.start()

        webbrowser.open("https://www.youtube.com/")

    ### ricerca su maps
    elif("cerca" in testo and "su google maps" in testo):
        ricerca = ""

        indexRicerca = len(parole) - parole[::-1].index("su") - 1 # trova la posizione dell'ultimo "su"

        for parola in parole[parole.index("cerca")+1:indexRicerca]:   # trova la città richiesta che sarà tra cerca e su (cerca...su)
            ricerca += parola +" "

        ricerca = ricerca.strip() # toglie lo spazio bianco alla fine della stringa

        risposta = f"cerco {ricerca} su google maps"

        x = t.Thread(target=speak, args=(risposta,)) # thread che farà partire la funzione speak la quale pronuncierà la risposta
        x.start()

        ricerca = ricerca.replace(" ", "+").strip() # per la creazione di un link google, al posto degli spazi vanno messi dei +

        link = "https://www.google.it/maps/place/"+ricerca # link completo che verrà usato per la ricerca web

        webbrowser.open(link)

    ### ricerca su youtube
    elif ("cerca" in parole and "su" in parole and "youtube" in parole):
        ricerca = ""

        indexRicerca = len(parole) - parole[::-1].index("su") - 1 # trova la posizione dell'ultimo "su"

        for parola in parole[parole.index("cerca")+1:indexRicerca]: # la ricerca è tra cerca e su (cerca ... su)
            ricerca += parola +" "
        
        ricerca = ricerca.strip() # rimuove lo spazio bianco finale

        risposta = f"cerco {ricerca} su youtube"

        x = t.Thread(target=speak, args=(risposta,)) # thread che farà partire la funzione speak la quale pronuncierà la risposta
        x.start()

        ricerca = ricerca.replace(" ", "+").strip() # per una ricerca youtube servono i "+" al posto degli spazi 

        link = "https://www.youtube.com/results?search_query="+ricerca  # link completo per la ricerca

        webbrowser.open(link)
        
    ### ricerca su google
    elif ("cerca" in parole and "su" in parole and "google" in parole):
        ricerca = ""

        indexRicerca = len(parole) - parole[::-1].index("su") - 1 # trova la posizione dell'ultimo "su"

        for parola in parole[parole.index("cerca")+1:indexRicerca]:   # la ricerca richiesta è tra cerca e su (cerca ... su)
            ricerca += parola +" "
        
        ricerca = ricerca.strip()
        risposta += f"cerco {ricerca} su google"

        x = t.Thread(target=speak, args=(risposta,)) # thread che farà partire la funzione speak la quale pronuncierà la risposta
        x.start()

        ricerca = ricerca.replace(" ", "+") # nella ricerca vanno messi i "+" al posto degli spazi

        link = "https://www.google.com.tr/search?q={}".format(ricerca)  # link completo per la ricerca

        webbrowser.open(link)

    ### ricerca su wikipedia
    elif ("cerca" in parole and "su" in parole and "wikipedia" in parole):
        ricerca = ""

        indexRicerca = len(parole) - parole[::-1].index("su") - 1 # trova la posizione dell'ultimo "su"

        for parola in parole[parole.index("cerca")+1:indexRicerca]: # prend la ricerca tra cerca e su (cerca ... su)
            ricerca += parola.title() +" "

        ricerca = ricerca.strip()

        risposta += f"cerco {ricerca} su wikipedia"

        x = t.Thread(target=speak, args=(risposta,)) # thread che farà partire la funzione speak la quale pronuncierà la risposta
        x.start()

        ricerca = ricerca.replace(" ", "_") # nella ricerca di wikipedia vanno messi gli underscore al posto degli spazi

        link = f"https://it.wikipedia.org/wiki/{ricerca}"

        webbrowser.open(link)

    elif ("traduci" in parole and "in" in parole and ("inglese" in parole or "spagnolo" in parole or "tedesco" in parole or "francese" in parole or "portoghese" in parole)):

        risposta = ""

        indexLingua = len(parole) - parole[::-1].index("in") - 1    # trova l ultima posizione della stringa "in"

        linguaDestinazione = parole[indexLingua+1]

        testoDaTradurre = parole[parole.index("traduci")+1:indexLingua] # prende la stringa che va da traduci a (traduci...in)

        testoDaTradurre = ' '.join(map(str, testoDaTradurre))   # trasforma la lista in una stringa formata da n parole quanti sono gli spazi

        risposta += traduciDaIta(testoDaTradurre, linguaDestinazione)   # restituisce come risposta il ritorno della funzione "tradiciDaIta"     
        
        x = t.Thread(target=speak, args=(risposta,)) # thread che farà partire la funzione speak la quale pronuncierà la risposta
        x.start()

    ######## inserire, stampare e azzerare le note
    elif("inserisci come nota" in testo):   # funzione per creare una nota

        f = open("note.txt", 'a')    # apertura del file di testo in modalità append contenente tutte le note
        frase = testo[testo.index("nota")+5:]   # inserisci la nota che va da inserisci come nota in poi
        f.write(frase+"\n") # aggiunge la nota al file senza cancellare le altre note
        f.close()   # chiude il file

        risposta += f"ho inserito {frase} alle tue note"

        x = t.Thread(target=speak, args=(risposta,)) # thread che farà partire la funzione speak la quale pronuncierà la risposta
        x.start()

    elif("mostra le mie note" in testo or "mostrami le mie note" in testo): # funzione per leggere tutte le tue note

        f = open("note.txt", "r")    # apre il file in modalità di lettura
        risposta += "\n"+f.read()+ "queste sono tutte le tue note"  # legge tutte le note e da la conferma di fine di tutte le note

        x = t.Thread(target=speak, args=(risposta,)) # thread che farà partire la funzione speak la quale pronuncierà la risposta
        x.start()

    elif("azzera le mie note" in testo or "azzera le note" in testo):   # funzione per azzerare tutte le note

        f = open("note.txt", "w")    # apertura del file in modalità scrittura
        f.write("") # azzera il file
        risposta += "\nNote azzerate"   # conferma l'azzeramento delle note

        x = t.Thread(target=speak, args=(risposta,)) # thread che farà partire la funzione speak la quale pronuncierà la risposta
        x.start()
    ########

    ### prende informazioni sul meteo della città richiesta tramite API
    elif "che tempo fa" in testo:

        try:
            nomeCitta = testo[15:]  # il nome della citta da va da "che tempo fa a" in poi

            indirizzoURL=f'https://api.openweathermap.org/data/2.5/weather?q={nomeCitta}&appid=5401c7c9dac626f92b0ab578ccee6135'

            data = requests.get(indirizzoURL).json()    # riceve il file JSON contente tutte le varie informazioni

            infoMeteo = data["weather"][0]["main"]  # prende dal file JSON la descrizione del meteo
            infoTemperatura = data["main"]["temp"]  # prende dal file JSON la temperatura 

            risposta_tempo = "A {} c'è {} con una temperatura di {} gradi".format(nomeCitta, traduciInIta(infoMeteo).lower(), int(infoTemperatura-273))   # essendo le informazioni in inglese, la funzione traduciInIta traduce dall'inglese all'italiano le informazioni
            risposta += risposta_tempo

            x = t.Thread(target=speak, args=(risposta,)) # thread che farà partire la funzione speak la quale pronuncierà la risposta
            x.start()

        except KeyError:    # restituisco errore nel caso la città non venga riconosciuta
            risposta += "citta non riconosciuta correttamente"

    ### ricerca film e serie tv in un database tramite API, il nome del film deve essere quello ufficiale in inglese essendo l'API inglese
    elif ("informazioni" in parole and ("film" in parole or "tv" in parole) and "del" in parole):

        nomeFilm = ""   # variabile che conterrà sia il nome del film oppure il nome della serie tv
        anno = ""

        if "film" in parole:    # richiesta in base al film
            indexRicerca = len(parole) - parole[::-1].index("del") - 1 # trova la posizione dell'ultimo "del"
            for parola in parole[parole.index("film")+1:indexRicerca]:  # il nome del film è tra (film ... del)
                nomeFilm += parola+" "
        elif "tv" in parole:    # richiesta in base alla serie tv
            indexRicerca = len(parole) - parole[::-1].index("del") - 1 # trova la posizione dell'ultimo "su"
            for parola in parole[parole.index("tv")+1:indexRicerca]:    # il nome della serie tv va da (serie tv...del)
                nomeFilm += parola+" "


        nomeFilm = nomeFilm.strip() # rimuove lo spazio vuoto finale

        rispostaVocale = f"eccoti le informazioni su {nomeFilm}"

        nomeFilm = nomeFilm.replace(" ", "+")   # per la ricerca tramite API bisogna sostituire gli spazi con dei "+"
        
        for parola in parole[parole.index("del")+1]:    # all'API serve anche una data, ossia la data di uscita del film, che verrà citata nella richiesta da parte dell'utente
            anno += parola
        
        link = f"http://www.omdbapi.com/?apikey=936d7f2c&t={nomeFilm}&y={anno}" # link completo della richiesta 

        response = urllib.request.urlopen(link)     # richiesta di informazioni dall'api tramite il passaggio in forma di link della richiesta

        with response as url:   # ricezione in un file JSON contenente le varie informazioni sul film richiesto
            data = json.loads(url.read().decode())  

        try:  

            ######## tutte le informazioni del film sono contenute in queste variabili
            infoTitolo = data["Title"]
            infoData = data["Released"]
            infoAnno = data["Year"]
            infoGenere = data["Genre"]
            infoRegista = data["Director"]
            infoAttori = data["Actors"]
            ########

            #### traduzione della trama in italiano dall'inglese tramite google translate
            trans = Translator()
            traduzione = trans.translate(data["Plot"], src="en", dest="it")
            infoTrama = traduzione.text
            ####

            risposta += "\n\ntitolo: {}\ndata di uscita: {}\ngenere: {}\nregista: {}\nattori: {}\ntrama: {}\n\n".format(infoTitolo, infoData, infoGenere, infoRegista, infoAttori, infoTrama)

            infoFilm = rispostaVocale+"\n" + infoTrama # la risposta vocale sarà formata dalla trama

            # if nel caso non ci siano informazioni sul regista, in caso contrario verranno restituite anche le informazioni sul regista
            if infoRegista == "N/A":
                rispostaVocale2 = "non sono presenti informazioni riguardo al regista"
            else:
                rispostaVocale2 = f"il regista è {infoRegista}"

            infoFilm += "\n"+rispostaVocale2    # alla risposta finale verranno aggiunte le informazioni sul regista

            x = t.Thread(target=speak, args=(infoFilm,)) # thread che farà partire la funzione speak la quale pronuncierà la risposta
            x.start()

        except: # errore che verrà richiamato nel caso di richiesta di un film non presente nel database
            risposta += "errore nella ricerca del film"
    ###

    else:   # nel caso non venga riconosciuta la funzione richiesta

        risposta += "scusa non ho capito la tua richiesta"
        
        x = t.Thread(target=speak, args=(risposta,)) # thread che farà partire la funzione speak la quale pronuncierà la risposta
        x.start()


    return risposta