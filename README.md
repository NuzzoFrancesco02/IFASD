# IFASD PDF to Excel Extractor

Questo script in Python è pensato per estrarre informazioni da file PDF (nel contesto IFASD) e inserirle in un file Excel. In particolare, il programma:
- Estrae il testo dalle prime pagine di un PDF utilizzando la libreria pdfminer.
- Pulisce il testo rimuovendo spazi, interruzioni di linea e caratteri superflui.
- Utilizza espressioni regolari per identificare ed estrarre dati chiave quali:
  - Titolo
  - Autori
  - Abstract
  - Parole chiave
- Inserisce i dati estratti nelle rispettive colonne di un file Excel tramite openpyxl.
- Salva tutti i titoli estratti in un file di testo (`titles.txt`).

Questa applicazione è stata sviluppata per l’elaborazione dei documenti IFASD (ad es. documenti con nome formattato come `IFASD-2019-XXX.pdf`).

## Requisiti

- Python 3.x

### Dipendenze Python
- numpy
- openpyxl
- pdfminer.six

Puoi installare le dipendenze eseguendo:

```bash
pip install numpy openpyxl pdfminer.six
```

## Installazione

1. Clona la repository:

    ```bash
    git clone https://github.com/NuzzoFrancesco02/IFASD.git
    ```

2. Entra nella cartella del progetto:

    ```bash
    cd IFASD
    ```

3. (Opzionale) Crea e attiva un ambiente virtuale:

    ```bash
    python -m venv venv
    source venv/bin/activate   # Su Windows: venv\Scripts\activate
    ```

4. Installa le dipendenze:
   Se disponi di un file `requirements.txt` (altrimenti installa manualmente):

    ```bash
    pip install -r requirements.txt
    ```

## Configurazione

Prima di eseguire il programma, verifica e modifica i seguenti parametri nel file `pdf2excel.py`:
- **Percorso del file Excel:** Modifica la variabile `excel_path` con il percorso corretto del tuo file Excel.
- **Percorso della cartella PDF:** Aggiorna la variabile `file_path` (utilizzata per generare il percorso completo dei file PDF) in base alla posizione in cui sono salvati i tuoi file PDF.
- **Limiti delle righe da processare:** Regola le variabili `raw_begin` e `raw_end` in modo da indicare l’intervallo di righe da elaborare nel file Excel.
- **Convenzione di denominazione dei file PDF:** Se necessario, modifica la logica per generare i nomi dei file (ad es. `IFASD-2019-{indice:03}.pdf`).

Le parti da personalizzare sono evidenziate nel codice con commenti come “DA MODIFICARE”.

## Utilizzo

Una volta configurato il file, puoi eseguire lo script con:

```bash
python pdf2excel.py
```

Il programma procederà a:
- Caricare il file Excel specificato.
- Processare le righe definite e per ciascuna cercare il PDF corrispondente.
- Estrarre e pulire il testo dalle prime pagine del PDF.
- Identificare ed estrarre titolo, autori, abstract e parole chiave.
- Scrivere i dati estratti nelle rispettive celle del file Excel.
- Salvare un file `titles.txt` contenente l’elenco dei titoli estratti.
- Aggiornare e salvare il workbook Excel.

## Contributi

Se desideri migliorare o estendere il progetto, sei libero di aprire una issue o inviare una pull request. Ogni contributo è benvenuto!

## Licenza

Questo progetto è distribuito con licenza MIT. Consulta il file `LICENSE` per maggiori dettagli.
