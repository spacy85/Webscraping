Scraper per il Monitoraggio di Scadenze su Sito Web

Descrizione:
Questo script Python è stato creato per effettuare il scraping delle informazioni relative alle scadenze da un sito web specifico e registrarle in un file Excel. Il programma utilizza il framework Selenium per l'automazione del browser e pandas per la manipolazione dei dati.

Funzionalità Principali:

Accesso al sito web tramite autenticazione.
Scorrimento delle pagine per ottenere tutte le informazioni disponibili.
Estrazione delle informazioni relative ai clienti, escludendo dati non necessari.
Gestione degli stati di avanzamento per il ripristino in caso di interruzioni.
Creazione di un DataFrame pandas per l'analisi dei dati.
Esportazione dei dati in un file Excel.
Utilizzo:

Per utilizzare questo script, è necessario fornire le credenziali di accesso al sito web nel file "login.txt".
L'autenticazione avviene automaticamente.
I dati vengono estratti dalle pagine e salvati in un file Excel chiamato "Scadenze.xlsx".
Gli stati di avanzamento vengono registrati in un file JSON per la possibilità di riprendere l'esecuzione in caso di interruzioni.
Requisiti:

Python 3.x
Librerie Python: requests, selenium, pandas, UliPlot
WebDriver per il browser Google Chrome
