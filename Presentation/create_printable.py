#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Script per generare una versione stampabile della presentazione Partnership Marrone-Iglu-Italfrigo
Questo script combina tutti i moduli HTML e genera un PDF ottimizzato per la stampa
"""

import os
import sys
import re
from pathlib import Path
import weasyprint
from bs4 import BeautifulSoup
import argparse
import time

# Colori per i messaggi
COLORI = {
    "verde": "\033[92m",
    "giallo": "\033[93m",
    "rosso": "\033[91m",
    "blu": "\033[94m",
    "magenta": "\033[95m",
    "ciano": "\033[96m",
    "bianco": "\033[97m",
    "reset": "\033[0m"
}

def stampa_colorato(testo, colore="verde"):
    """Stampa testo colorato nella console."""
    print(f"{COLORI.get(colore, COLORI['verde'])}{testo}{COLORI['reset']}")

def main():
    parser = argparse.ArgumentParser(description='Crea una versione stampabile della presentazione RH')
    parser.add_argument('--output', '-o', default='RH_Corso_Venezia_Presentazione.pdf', 
                        help='Nome del file PDF di output')
    parser.add_argument('--lang', '-l', default='it', choices=['it', 'en'],
                        help='Lingua della presentazione (it=italiano, en=inglese)')
    parser.add_argument('--both', '-b', action='store_true',
                        help='Genera entrambe le versioni (italiano e inglese)')
    args = parser.parse_args()
    
    # Directory della presentazione
    base_dir = Path(__file__).parent
    modules_dir = base_dir / "modules"
    modules_en_dir = base_dir / "modules_en"
    
    # Definisci i nomi dei file di output in base alla lingua
    if args.both:
        output_it = base_dir / "downloads" / "RH_Corso_Venezia_Presentazione_IT.pdf"
        output_en = base_dir / "downloads" / "RH_Corso_Venezia_Presentazione_EN.pdf"
    else:
        if args.lang == 'it':
            output_file = base_dir / "downloads" / "RH_Corso_Venezia_Presentazione_IT.pdf"
            if args.output != 'RH_Corso_Venezia_Presentazione.pdf':
                output_file = base_dir / "downloads" / args.output
        else:
            output_file = base_dir / "downloads" / "RH_Corso_Venezia_Presentazione_EN.pdf"
            if args.output != 'RH_Corso_Venezia_Presentazione.pdf':
                output_file = base_dir / "downloads" / args.output
    
    # Crea la directory downloads se non esiste
    os.makedirs(base_dir / "downloads", exist_ok=True)
    
    # Moduli in ordine di presentazione
    module_files = [
        "01_introduction.html",
        "02_marrone.html",
        "03_iglu.html",
        "04_italfrigo.html",
        "05_technical.html",
        "06_realization.html",
        "07_services.html",
        "08_final.html",
        "09_thanks.html"
    ]
    
    # Crea un documento HTML completo
    combined_html = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>RH Corso Venezia - Presentazione Stampabile</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                line-height: 1.6;
                color: #333;
                max-width: 100%;
            }
            h1, h2, h3, h4 {
                page-break-after: avoid;
                margin-top: 20px;
            }
            h2 {
                color: #333;
                font-size: 24pt;
                margin-bottom: 10px;
                text-align: center;
            }
            h3 {
                color: #444;
                border-bottom: 1px solid #ddd;
                padding-bottom: 10px;
            }
            p {
                margin-bottom: 10px;
            }
            img {
                max-width: 100%;
                height: auto;
            }
            ul, ol {
                margin-left: 20px;
                margin-bottom: 15px;
            }
            .page-break {
                page-break-before: always;
            }
            .section-content {
                margin-bottom: 30px;
            }
            .flex-container {
                display: flex;
                flex-wrap: wrap;
                gap: 20px;
                margin-bottom: 20px;
            }
            .flex-item {
                flex: 1;
                min-width: 45%;
                background-color: #f9f9f9;
                padding: 15px;
                border-radius: 5px;
            }
            .logo-container {
                display: flex;
                justify-content: center;
                gap: 40px;
                margin: 30px 0;
                flex-wrap: wrap;
            }
            .logo-item {
                text-align: center;
                max-width: 200px;
            }
            .cover-page {
                text-align: center;
                margin-bottom: 50px;
            }
            .cover-page h1 {
                font-size: 32pt;
                margin-bottom: 20px;
                color: #222;
            }
            .cover-page p {
                font-size: 14pt;
                color: #666;
                font-style: italic;
            }
            .footer {
                text-align: center;
                margin-top: 30px;
                font-size: 10pt;
                color: #666;
            }
            @page {
                size: A4;
                margin: 2cm;
                @bottom-center {
                    content: "RH Corso Venezia - " counter(page) " di " counter(pages);
                }
            }
        </style>
    </head>
    <body>
        <div class="cover-page">
            <h1>RH Corso Venezia</h1>
            <p>Presentazione del Progetto</p>
            <div class="logo-container">
                <div class="logo-item">
                    <img src="assets/logos/italfrigo-logo.png" alt="Italfrigo Logo">
                </div>
                <div class="logo-item">
                    <img src="assets/logos/marrone-logo.png" alt="Marrone Logo">
                </div>
                <div class="logo-item">
                    <img src="assets/logos/iglu-logo.jpg" alt="IGLU Logo">
                </div>
            </div>
        </div>
    """
    
    # Funzione per generare il PDF in una specifica lingua
    def genera_pdf_per_lingua(lingua):
        nonlocal combined_html
        
        # Reset dell'HTML combinato
        combined_html = """
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>Partnership Marrone-Iglu-Italfrigo - Presentazione Stampabile</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    line-height: 1.6;
                    color: #333;
                    max-width: 100%;
                }
                h1, h2, h3, h4 {
                    page-break-after: avoid;
                    margin-top: 20px;
                }
                h2 {
                    color: #333;
                    font-size: 24pt;
                    margin-bottom: 10px;
                    text-align: center;
                }
                h3 {
                    color: #444;
                    border-bottom: 1px solid #ddd;
                    padding-bottom: 10px;
                }
                p {
                    margin-bottom: 10px;
                }
                img {
                    max-width: 100%;
                    height: auto;
                }
                ul, ol {
                    margin-left: 20px;
                    margin-bottom: 15px;
                }
                .page-break {
                    page-break-before: always;
                }
                .section-content {
                    margin-bottom: 30px;
                }
                .flex-container {
                    display: flex;
                    flex-wrap: wrap;
                    gap: 20px;
                    margin-bottom: 20px;
                }
                .flex-item {
                    flex: 1;
                    min-width: 45%;
                    background-color: #f9f9f9;
                    padding: 15px;
                    border-radius: 5px;
                }
                .logo-container {
                    display: flex;
                    justify-content: center;
                    gap: 40px;
                    margin: 30px 0;
                    flex-wrap: wrap;
                }
                .logo-item {
                    text-align: center;
                    max-width: 200px;
                }
                .cover-page {
                    text-align: center;
                    margin-bottom: 50px;
                }
                .cover-page h1 {
                    font-size: 32pt;
                    margin-bottom: 20px;
                    color: #222;
                }
                .cover-page p {
                    font-size: 14pt;
                    color: #666;
                    font-style: italic;
                }
                .footer {
                    text-align: center;
                    margin-top: 30px;
                    font-size: 10pt;
                    color: #666;
                }
                @page {
                    size: A4;
                    margin: 2cm;
                    @bottom-center {
                        content: "Partnership Marrone-Iglu-Italfrigo - " counter(page) " di " counter(pages);
                    }
                }
            </style>
        </head>
        <body>
            <div class="cover-page">
                <h1>Partnership Marrone-Iglu-Italfrigo</h1>
        """
        
        # Aggiungi il sottotitolo in base alla lingua
        if lingua == 'it':
            combined_html += '<p>Presentazione del Progetto</p>'
            current_modules_dir = modules_dir
        else:
            combined_html += '<p>Project Presentation</p>'
            current_modules_dir = modules_en_dir
        
        combined_html += """
                <div class="logo-container">
                    <div class="logo-item">
                        <img src="assets/logos/italfrigo-logo.png" alt="Italfrigo Logo">
                    </div>
                    <div class="logo-item">
                        <img src="assets/logos/marrone-logo.png" alt="Marrone Logo">
                    </div>
                    <div class="logo-item">
                        <img src="assets/logos/iglu-logo.jpg" alt="IGLU Logo">
                    </div>
                </div>
            </div>
        """
        
        # Combina tutti i moduli
        for i, module_file in enumerate(module_files):
            module_path = current_modules_dir / module_file
            if not module_path.exists():
                stampa_colorato(f"Attenzione: File {module_file} non trovato in {current_modules_dir}", "giallo")
                # Prova a usare la versione italiana come fallback
                if lingua == 'en':
                    fallback_path = modules_dir / module_file
                    if fallback_path.exists():
                        module_path = fallback_path
                        stampa_colorato(f"Usando la versione italiana come fallback per {module_file}", "giallo")
                    else:
                        continue
                else:
                    continue
            
            stampa_colorato(f"Elaborazione di {module_file} ({lingua})...", "ciano")
            
            # Leggi il contenuto del modulo
            with open(module_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Estrai il contenuto principale
            soup = BeautifulSoup(content, 'html.parser')
            main_content = soup.find('div', style=lambda s: s and 'width: 100%' in s)
            
            if not main_content:
                stampa_colorato(f"Impossibile trovare il contenuto principale in {module_file}", "giallo")
                # Salta questo modulo e passa al prossimo
                continue
            
            # Aggiungi un'interruzione di pagina prima di ogni modulo (tranne il primo)
            if i > 0:
                combined_html += '<div class="page-break"></div>\n'
            
            # Aggiungi il contenuto al documento combinato
            combined_html += f'<div class="section-content">{str(main_content)}</div>\n'
    
        # Chiudi il documento HTML
        if lingua == 'it':
            footer_text = "© 2025 - Presentazione creata per Partnership Marrone-Iglu-Italfrigo"
        else:
            footer_text = "© 2025 - Presentation created for Partnership Marrone-Iglu-Italfrigo"
            
        combined_html += f"""
            <div class="footer">
                <p>{footer_text}</p>
            </div>
        </body>
        </html>
        """
    
        # Correggi i percorsi delle immagini
        combined_html = combined_html.replace('src="../assets/', 'src="assets/')
        
        # Determina il nome del file di output
        if args.both:
            if lingua == 'it':
                current_output = output_it
                html_output = base_dir / "downloads" / "RH_Corso_Venezia_Presentazione_IT.html"
            else:
                current_output = output_en
                html_output = base_dir / "downloads" / "RH_Corso_Venezia_Presentazione_EN.html"
        else:
            current_output = output_file
            if lingua == 'it':
                html_output = base_dir / "downloads" / "RH_Corso_Venezia_Presentazione_IT.html"
            else:
                html_output = base_dir / "downloads" / "RH_Corso_Venezia_Presentazione_EN.html"
        
        # Genera il PDF
        stampa_colorato(f"Generazione del PDF {current_output}...", "blu")
        html = weasyprint.HTML(string=combined_html, base_url=str(base_dir))
        html.write_pdf(current_output)
        
        stampa_colorato(f"PDF creato con successo: {current_output}", "verde")
        
        # Crea anche una versione HTML per debug
        with open(html_output, 'w', encoding='utf-8') as f:
            f.write(combined_html)
        
        stampa_colorato(f"Versione HTML creata: {html_output}", "verde")
        
        return current_output
    
    # Aggiorna i link di download nei moduli
    def aggiorna_link_download():
        # Percorsi dei file di download per entrambe le lingue
        modulo_it = base_dir / "modules" / "07_5_download.html"
        modulo_en = base_dir / "modules_en" / "07_5_download.html"
        
        # Nomi dei file PDF per ciascuna lingua
        pdf_it = "RH_Corso_Venezia_Presentazione_IT.pdf"
        pdf_en = "RH_Corso_Venezia_Presentazione_EN.pdf"
        
        # Aggiorna il modulo italiano
        try:
            with open(modulo_it, 'r', encoding='utf-8') as file:
                contenuto = file.read()
            
            # Sostituisci il link alla presentazione
            contenuto_modificato = contenuto.replace(
                'href="../downloads/RH_Corso_Venezia_Presentazione.pdf"',
                f'href="../downloads/{pdf_it}"'
            )
            
            with open(modulo_it, 'w', encoding='utf-8') as file:
                file.write(contenuto_modificato)
            
            stampa_colorato(f"Link nel modulo italiano aggiornato", "verde")
        except Exception as e:
            stampa_colorato(f"Errore nell'aggiornamento del modulo italiano: {e}", "rosso")
        
        # Aggiorna il modulo inglese
        try:
            with open(modulo_en, 'r', encoding='utf-8') as file:
                contenuto = file.read()
            
            # Sostituisci il link alla presentazione
            contenuto_modificato = contenuto.replace(
                'href="../downloads/RH_Corso_Venezia_Presentazione.pdf"',
                f'href="../downloads/{pdf_en}"'
            )
            
            with open(modulo_en, 'w', encoding='utf-8') as file:
                file.write(contenuto_modificato)
            
            stampa_colorato(f"Link nel modulo inglese aggiornato", "verde")
        except Exception as e:
            stampa_colorato(f"Errore nell'aggiornamento del modulo inglese: {e}", "rosso")
    
    # Tempo di inizio
    start_time = time.time()
    
    # Genera i PDF in base alle opzioni
    if args.both:
        stampa_colorato("Generazione di entrambe le versioni (IT+EN)...", "magenta")
        output_it_path = genera_pdf_per_lingua('it')
        output_en_path = genera_pdf_per_lingua('en')
        aggiorna_link_download()
        
        stampa_colorato("\nRiepilogo:", "magenta")
        stampa_colorato(f"- PDF Italiano: {output_it_path}", "verde")
        stampa_colorato(f"- PDF Inglese: {output_en_path}", "verde")
    else:
        output_path = genera_pdf_per_lingua(args.lang)
        aggiorna_link_download()
        
        stampa_colorato("\nRiepilogo:", "magenta")
        stampa_colorato(f"- PDF ({args.lang.upper()}): {output_path}", "verde")
    
    # Tempo di fine
    end_time = time.time()
    stampa_colorato(f"Tempo impiegato: {end_time - start_time:.2f} secondi", "ciano")

if __name__ == "__main__":
    stampa_colorato("=== Generazione PDF Presentazione Partnership Marrone-Iglu-Italfrigo ===", "magenta")
    main()
    stampa_colorato("\nGenerazione completata con successo!", "verde")
