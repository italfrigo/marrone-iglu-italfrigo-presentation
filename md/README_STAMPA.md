# Istruzioni per la Versione Stampabile della Presentazione RH Corso Venezia

Questo documento spiega come generare la versione stampabile della presentazione e l'offerta in formato PDF.

## Requisiti

Per generare i PDF, è necessario installare le seguenti librerie Python:

```bash
pip install weasyprint beautifulsoup4
```

Nota: WeasyPrint richiede alcune dipendenze di sistema. Su Ubuntu/Debian, puoi installarle con:

```bash
sudo apt-get install build-essential python3-dev python3-pip python3-setuptools python3-wheel python3-cffi libcairo2 libpango-1.0-0 libpangocairo-1.0-0 libgdk-pixbuf2.0-0 libffi-dev shared-mime-info
```

## Generazione della Presentazione Stampabile

Per generare la versione PDF della presentazione:

```bash
python3 create_printable.py
```

Questo comando creerà un file PDF nella cartella `downloads` chiamato `RH_Corso_Venezia_Presentazione.pdf`.

## Personalizzazione

È possibile specificare un nome di file diverso per l'output:

```bash
python3 create_printable.py --output NomePersonalizzato.pdf
```

## Note Importanti

- Il PDF generato è ottimizzato per la stampa in formato A4
- Tutte le slide sono formattate per garantire la leggibilità su carta
- Vengono inserite automaticamente interruzioni di pagina tra le diverse sezioni
- I loghi e le immagini sono ridimensionati per adattarsi al formato stampabile

## Creazione dell'Offerta

Per creare il file dell'offerta, è necessario preparare un documento separato e salvarlo come `RH_Corso_Venezia_Offerta.pdf` nella cartella `downloads`.

## Struttura dei File

```
Presentation/
├── create_printable.py     # Script per generare il PDF
├── downloads/              # Cartella per i file scaricabili
│   ├── RH_Corso_Venezia_Presentazione.pdf
│   └── RH_Corso_Venezia_Offerta.pdf
├── modules/                # Moduli HTML della presentazione
└── README_STAMPA.md        # Questo file
```
