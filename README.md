# Generatore di Mail per Training (Outlook AI)

Software desktop locale sviluppato in Python appositamente per il corso **"IA per Outlook"**. 
L'applicativo genera grandi quantità di email aziendali fittizie (in formato `.eml`) utili per attività didattiche, esercizi pratici ed esperimenti con pattern e automatismi legati all'Intelligenza Artificiale applicata alla gestione della posta elettronica.

## 🚀 Funzionalità Principali

* **Interfaccia Grafica Intuitiva:** Sviluppata in `CustomTkinter` per un'esperienza d'uso semplice, fluida e moderna (modalità Dark integrata).
* **Personalizzazione Avanzata dello Scenario:** 
  * Inserimento dei dati utente (Nome e alias Email simulato).
  * **Volume Dati:** Selezione rapida del numero totale di email da generare (da 10 a 500 email).
  * **Complessità Rete:** Configurazione del numero di "colleghi" simulati (da 2 a 20) per variare le interazioni.
  * **Allegati:** Impostazione della probabilità percentuale (0-100%) che un messaggio in ingresso contenga documenti allegati (fittizi).
  * **Timeline:** Definizione di una finestra temporale (Data Inizio e Data Fine) in cui generare realisticamente le date e gli orari di ricezione.
* **Contenuti Realistici per Outlook:**
  * Mittenti e destinatari assegnati in modo casuale partendo da un vasto *pool* preconfigurato di nomi e domini (es. `azienda-finta.it`).
  * Oggetti e corpi testuali (sia HTML che Plain Text) ideati per un contesto "corporate" verosimile (es. richieste di report, circolari su smart working, inviti a webinar).
  * Supporto nativo alla generazione di allegati (File `.xlsx`, `.pdf`, `.csv`, `.png` fittizi).
* **Compatibilità Ottimizzata MS Outlook:** I messaggi (salvati come file `.eml`) sono costruiti con la libreria interna `email.mime` e sfruttano header Mimeole/Unsent specifici presi da Microsoft Outlook, assicurando un'importazione perfetta nel client, supportato anche dal forzamento dei ritorni a capo (CRLF).
* **Export "Pronto all'Uso":** Tutte le mail prodotte vengono impacchettate direttamente in RAM e salvate in un unico comodo archivio `.zip` sul disco fisso.

## 🛠 Requisiti Tecnici

Il progetto richiede **Python 3.x**.  
L'unica dipendenza esterna di terze parti necessaria al funzionamento dell'interfaccia grafica è `customtkinter`. Le altre librerie (come `zipfile`, `email`, `threading`) sono native di Python.

Per installare l'ambiente corretto esegui:
```bash
pip install customtkinter
```

## ⚙️ Istruzioni di Avvio

1. Esegui il file principale tramite Python:
   ```bash
   python "generatore con interfaccia v2.py"
   ```
2. Modifica (se necessario) le credenziali dell'utente e sposta comodamente gli slider in base allo scenario desiderato per la lezione.
3. Seleziona la cartella di destinazione tramite il tasto **"Scegli Cartella..."**.
4. Avvia il tool cliccando il tastone **"GENERA ARCHIVIO .ZIP"**.
5. Al termine della barra di caricamento, l'operazione in background produrrà un archivio `.zip` già organizzato contenente decine (o centinaia) di file `.eml` fittizi pronti da trascinare all'interno del proprio client Outlook per fare pratica.
