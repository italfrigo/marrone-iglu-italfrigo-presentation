import pandas as pd
import os
import jinja2
import webbrowser
from datetime import datetime

def converti_excel_in_html(file_excel, output_html):
    """
    Converte un file Excel in un documento HTML elegante
    """
    print(f"Conversione di {file_excel} in HTML...")

    try:
        # Prova a leggere i fogli specifici
        xls = pd.ExcelFile(file_excel)
        if 'Informazioni' in xls.sheet_names and 'Offerta' in xls.sheet_names:
            info_df = pd.read_excel(xls, sheet_name='Informazioni')
            offerta_df = pd.read_excel(xls, sheet_name='Offerta')
        else:
            # Se non ci sono i fogli specifici, leggi il primo foglio
            print("Fogli 'Informazioni' e 'Offerta' non trovati. Lettura del primo foglio...")
            offerta_df = pd.read_excel(xls)
            # Cerca di estrarre le informazioni se le colonne esistono
            if 'Chiave' in offerta_df.columns and 'Valore' in offerta_df.columns:
                # Estrai le righe che sembrano informazioni (hanno sia Chiave che Valore)
                info_mask = offerta_df['Chiave'].notna() & offerta_df['Valore'].notna()
                info_df = offerta_df[info_mask][['Chiave', 'Valore']].copy()
                # Rimuovi le righe di info dal dataframe dell'offerta
                offerta_df = offerta_df[~info_mask].copy()
                # Se dopo la rimozione l'offerta è vuota ma le info c'erano, svuota le info
                # Questo gestisce il caso in cui l'Excel originale aveva solo le info
                if offerta_df.empty and not info_df.empty:
                     offerta_df = info_df.copy() # Metti le info nell'offerta se è l'unica cosa presente
                     info_df = pd.DataFrame(columns=['Chiave', 'Valore']) # Svuota le info
            else:
                 info_df = pd.DataFrame(columns=['Chiave', 'Valore']) # Crea un DataFrame info vuoto

    except Exception as e:
        print(f"Errore durante la lettura del file Excel: {e}")
        return None

    # Estrai le informazioni dell'offerta dal DataFrame info_df
    info_dict = {}
    if not info_df.empty:
        for _, row in info_df.iterrows():
            if pd.notna(row['Chiave']) and pd.notna(row['Valore']):
                key = str(row['Chiave']).strip(':') if isinstance(row['Chiave'], str) else row['Chiave']
                value = row['Valore']
                # Formatta le date se sono Timestamp
                if isinstance(value, pd.Timestamp):
                    value = value.strftime('%d/%m/%Y')
                info_dict[key] = value

    # --- MODIFICA: Imposta il cliente manualmente ---
    info_dict['Cliente'] = 'RH'
    # -----------------------------------------------

    # Prepara i dati dell'offerta
    offerta_data = []
    headers = []
    prezzo_netto_key = 'Prezzo netto' # Assumiamo che la colonna si chiami così
    quantita_key = 'Quantità'        # Assumiamo che la colonna si chiami così

    if not offerta_df.empty:
        # Usa le intestazioni reali dal DataFrame offerta_df
        headers = offerta_df.columns.tolist()
        # Pulisci i dati, sostituendo NaN con stringhe vuote
        offerta_df_filled = offerta_df.fillna('')
        offerta_data_list = offerta_df_filled.to_dict('records')

        # Aggiungi il flag is_category, formatta i prezzi 
        offerta_data = []
        if headers: # Assicurati che ci siano headers
            first_col_key = headers[0]
            other_col_keys = headers[1:]
            if prezzo_netto_key not in headers:
                 print(f"Attenzione: Colonna '{prezzo_netto_key}' non trovata negli headers: {headers}")

            for row_dict in offerta_data_list:
                # 1. Determina se è una riga di categoria
                is_category = False
                first_val = row_dict.get(first_col_key, '')
                if first_val != '':
                    all_others_empty = all(row_dict.get(col_key, '') == '' for col_key in other_col_keys)
                    if all_others_empty:
                        is_category = True
                row_dict['is_category'] = is_category

                # 2. Aggiungi simbolo Euro 
                original_value = row_dict.get(prezzo_netto_key, '')
                formatted_prezzo = original_value # Default al valore originale

                if not is_category and prezzo_netto_key in row_dict and original_value != '':
                    # Calcolo Totale e Formattazione Visualizzazione: Tenta una conversione accurata
                    val_float = None # Cambiato nome per chiarezza
                    try:
                        if isinstance(original_value, (int, float)):
                            val_float = float(original_value)
                        elif isinstance(original_value, str):
                            val_str = original_value.strip().replace(' ', '')
                            if val_str:
                                has_comma = ',' in val_str
                                has_dot = '.' in val_str
                                cleaned_str_for_float = val_str # Default

                                if has_comma and has_dot:
                                    cleaned_str_for_float = val_str.replace('.', '').replace(',', '.')
                                elif has_comma:
                                    cleaned_str_for_float = val_str.replace(',', '.')
                                elif has_dot:
                                     cleaned_str_for_float = val_str

                                val_float = float(cleaned_str_for_float)

                    except (ValueError, TypeError):
                        # Se la conversione fallisce, usa valore originale per display
                        print(f"Attenzione: Impossibile convertire '{original_value}' in numero. Verrà usato il valore originale per display.")
                        formatted_prezzo = f"€ {str(original_value)}" # Solo aggiunta €
                        # val_float resta None

                    # Se la conversione è riuscita, formatta 
                    if val_float is not None:
                        try:
                             # Formatta a due decimali per display
                            formatted_prezzo = f"€ {val_float:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                        except Exception as e_format:
                            print(f"Errore formattazione valore {val_float}: {e_format}. Uso valore originale.")
                            formatted_prezzo = f"€ {str(original_value)}" # Fallback a originale + €
                    # else: # Conversione fallita, formatted_prezzo è già stato impostato nel blocco except
                        pass
                # else: è categoria o valore vuoto, formatted_prezzo resta il valore originale ''

                # Imposta il valore formattato per il template
                row_dict['formatted_prezzo'] = formatted_prezzo
                offerta_data.append(row_dict)
        else:
            offerta_data = offerta_data_list # Nessun header, passa i dati così come sono

    else:
        print("Attenzione: Nessun dato trovato per la tabella dell'offerta.")

    # Template HTML Semplificato
    template_str = """
    <!DOCTYPE html>
    <html lang="it">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{{ info.get('Riferimento', 'Offerta Cucina Professionale - Partnership Marrone Iglu Italfrigo') }}</title>
        <style>
            body {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                line-height: 1.6;
                margin: 0;
                padding: 0;
                background-color: #f8f9fa;
                color: #212529;
            }
            .container {
                max-width: 1140px;
                margin: 30px auto;
                background-color: #ffffff;
                padding: 40px;
                border-radius: 8px;
                box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
            }
            h1, h2 {
                color: #0056b3; /* Blu aziendale */
                border-bottom: 2px solid #dee2e6;
                padding-bottom: 10px;
                margin-top: 0;
                margin-bottom: 25px;
            }
            h1 {
                text-align: center;
                font-size: 2em;
            }
            h2 {
                font-size: 1.5em;
            }
            .info-section {
                margin-bottom: 35px;
                padding: 25px;
                background-color: #f1f3f5;
                border: 1px solid #e9ecef;
                border-radius: 6px;
            }
            .info-item {
                margin-bottom: 12px;
                display: flex;
                flex-wrap: wrap;
            }
            .info-label {
                font-weight: 600; /* SemiBold */
                color: #495057;
                min-width: 150px;
                padding-right: 10px;
            }
            .info-value {
                 color: #343a40;
            }
            table {
                width: 100%;
                border-collapse: collapse;
                margin-top: 25px;
                box-shadow: 0 2px 5px rgba(0,0,0,0.05);
            }
            th, td {
                border: 1px solid #dee2e6;
                padding: 12px 15px;
                text-align: left;
            }
            thead {
                background-color: #0056b3; /* Blu aziendale */
                color: white;
                font-weight: 600;
            }
            tbody tr:nth-child(odd) {
                background-color: #ffffff;
            }
            tbody tr:nth-child(even) {
                background-color: #f8f9fa;
            }
            tbody tr:hover {
                 background-color: #e9ecef;
            }
            tr.category {
                background-color: #cfe2ff !important; /* Blu chiaro per categoria */
                font-weight: bold;
                color: #004085;
            }
            .footer {
                text-align: center;
                margin-top: 40px;
                padding-top: 20px;
                border-top: 1px solid #dee2e6;
                font-size: 0.9em;
                color: #6c757d;
            }
            @media print {
                body { background-color: white; margin: 0; }
                .container { box-shadow: none; border: none; padding: 0; max-width: 100%; margin: 0; }
                .info-section { background-color: #f8f9fa; border: 1px solid #eee; }
                table { box-shadow: none; }
                thead { background-color: #e9ecef; color: #333; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                tr.category { background-color: #e2e3e5 !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                .footer { border-top: 1px solid #ccc; }
                h1, h2 { border-bottom: 1px solid #ccc; }
            }

            /* --- Stili Intestazione --- */
            .header-section {
                display: flex;
                justify-content: space-between;
                align-items: flex-start; /* Allinea in alto */
                margin-bottom: 40px;
                padding-bottom: 20px;
                border-bottom: 1px solid #dee2e6;
                flex-wrap: wrap; /* Permette di andare a capo su schermi piccoli */
            }
            .logo-container {
                flex: 0 0 auto; /* Non crescere, non restringersi, base auto */
                margin-right: 30px;
                margin-bottom: 15px; /* Margine inferiore se va a capo */
            }
            .logo-container img {
                max-width: 200px; /* Larghezza massima logo */
                height: auto;
            }
            .company-details {
                flex: 1 1 auto; /* Cresce, si restringe, base auto */
                text-align: right;
                font-size: 0.9em;
                line-height: 1.4;
                color: #495057;
                margin-bottom: 15px; /* Margine inferiore se va a capo */
            }
            .company-details p {
                margin: 2px 0;
            }
            .company-details strong {
                color: #212529;
                font-weight: 600;
            }
            /* ------------------------- */

            /* --- Stili Colonne Tabella --- */
            .col-quantita {
                width: 6%; /* Larghezza ridotta */
                max-width: 50px; /* Limite massimo per PDF */
                /* text-align: center; Rimossa da qui, applicata sotto specificamente */
            }
            .col-prezzo-netto {
                width: 18%; /* Larghezza aumentata */
                text-align: right; /* Testo allineato a destra */
            }
            /* Applica centratura a header e celle quantità */
            th.col-quantita, td.col-quantita {
                text-align: center;
            }
            /* Applica allineamento a destra a header e celle prezzo */
            th.col-prezzo-netto, td.col-prezzo-netto {
                text-align: right;
            }
            /* --------------------------- */

            /* --- Stili Clausole --- */
            .clauses-section {
                margin-top: 40px;
                padding: 20px;
                border: 1px dashed #adb5bd;
                background-color: #f8f9fa;
                border-radius: 4px;
                font-size: 0.9em;
            }
            .clauses-section h3 {
                margin-top: 0;
                margin-bottom: 15px;
                color: #343a40;
                font-size: 1.1em;
                border-bottom: 1px solid #dee2e6;
                padding-bottom: 5px;
            }
            .clauses-section p, .clauses-section ul {
                margin-bottom: 10px;
                line-height: 1.5;
            }
            .clauses-section ul {
                padding-left: 20px;
                list-style: disc;
            }
            .clauses-section li {
                margin-bottom: 5px;
            }
            /* -------------------- */

        </style>
    </head>
    <body>
        <div class="container">
            <!-- === Sezione Intestazione === -->
            <div class="header-section">
                <div class="logo-container">
                    <img src="immagini italfrigo/logo_italfrigo.png" alt="Logo Italfrigo Service Srl">
                </div>
                <div class="company-details">
                    <p><strong>ITALFRIGO SERVICE SRL</strong></p>
                    <p>Via Pelizza Da Volpedo 49/A</p>
                    <p>20092 – Cinisello Balsamo (MI), Italia</p>
                    <p>Tel: +39 0266040933</p>
                    <p>P.IVA: 06825390153</p>
                    <br>
                    <p><em>Contatto:</em> Manuel Lazzaro</p>
                    <p><em>Email:</em> manuel.lazzaro@italfrigo.com</p>
                    <p><em>Mobile:</em> +39 3461587689</p>
                </div>
            </div>
            <!-- ============================ -->

            <h1>{{ info.get('Riferimento', 'Offerta Commerciale') }}</h1>

            {% if info %}
            <div class="info-section">
                <h2>Informazioni Generali</h2>
                {% for key, value in info.items() %}
                <div class="info-item">
                    <span class="info-label">{{ key }}:</span>
                    <span class="info-value">{{ value }}</span>
                </div>
                {% endfor %}
            </div>
            {% endif %}

            {% if offerta_data %}
            <h2>Dettaglio Offerta</h2>
            <table>
                <thead>
                    <tr>
                        {% for header in headers %}
                        <th {% if header == quantita_key %}class="col-quantita"{% elif header == prezzo_netto_key %}class="col-prezzo-netto"{% endif %}>{{ header }}</th>
                        {% endfor %}
                    </tr>
                </thead>
                <tbody>
                    {% for row in offerta_data %}
                    {# Usa il flag precalcolato #}
                    <tr {% if row.is_category %}class="category"{% endif %}>
                        {% for header in headers %}
                        {# Accedi ai dati usando l'header come chiave, usa prezzo formattato se è la colonna giusta #}
                        <td {% if header == quantita_key %}class="col-quantita"{% elif header == prezzo_netto_key %}class="col-prezzo-netto"{% endif %}>{% if header == prezzo_netto_key %}{{ row.formatted_prezzo }}{% else %}{{ row.get(header, '') }}{% endif %}</td>
                        {% endfor %}
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
            {% else %}
            <p>Nessun dettaglio offerta disponibile nel file Excel.</p>
            {% endif %}

            <!-- === Sezione Clausole Commerciali === -->
            <div class="clauses-section">
                <h3>Clausole Commerciali</h3>
                <p><strong>Incluso:</strong></p>
                <ul>
                    <li>Trasporto</li>
                    <li>Montaggio</li>
                    <li>Collaudo</li>
                </ul>
                <p><strong>Escluso:</strong></p>
                <ul>
                    <li>Collegamenti elettrici</li>
                    <li>Collegamenti idraulici</li>
                    <li>Eventuali opere murarie</li>
                    <li>Tutto quanto non espressamente citato nella presente offerta</li>
                </ul>
                <p><strong>Pagamento:</strong> Da concordare.</p>
                <p><strong>Data di Consegna:</strong> Da concordare.</p>
                <br>
                <p><em><strong>Nota:</strong> Tutti i prezzi sono da intendersi IVA 22% esclusa.</em></p>
            </div>
            <!-- ==================================== -->

            <div class="footer">
                Documento generato il {{ data_generazione }}
            </div>
        </div>
    </body>
    </html>
    """

    try:
        # Crea il template Jinja2
        env = jinja2.Environment(loader=jinja2.BaseLoader())
        template = env.from_string(template_str)

        # Genera l'HTML
        html_content = template.render(
            info=info_dict,
            headers=headers,
            offerta_data=offerta_data,
            prezzo_netto_key=prezzo_netto_key, # Passa la chiave per controllo nel template
            quantita_key=quantita_key,         # Passa la chiave per controllo nel template
            data_generazione=datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        )

        # Scrivi il file HTML
        with open(output_html, 'w', encoding='utf-8') as f:
            f.write(html_content)

        print(f"File HTML creato con successo: {output_html}")
        return output_html

    except jinja2.exceptions.TemplateError as e:
        print(f"Errore nel template Jinja2: {e}")
        if hasattr(e, 'lineno') and e.lineno:
             print(f"Errore alla linea: {e.lineno}")
        # Stampa una parte del template intorno all'errore se possibile
        try:
            lines = template_str.splitlines()
            start = max(0, e.lineno - 4)
            end = min(len(lines), e.lineno + 3)
            print("--- Template Vicino all'errore ---")
            for i, line in enumerate(lines[start:end], start=start + 1):
                print(f"{i:4d}{'*' if i == e.lineno else ' '}: {line}")
            print("----------------------------------")
        except Exception as inner_e:
            print(f"(Impossibile mostrare il contesto del template: {inner_e})")
        return None
    except Exception as e:
        print(f"Errore durante la generazione dell'HTML: {e}")
        import traceback
        traceback.print_exc() # Stampa il traceback completo per debug
        return None

def converti_html_in_pdf(file_html, output_pdf):
    """
    Converte un file HTML in PDF utilizzando weasyprint
    """
    try:
        # Prova ad importare weasyprint
        import importlib
        if importlib.util.find_spec("weasyprint") is None:
             print("\nATTENZIONE: La libreria 'weasyprint' non è installata.")
             print("Per generare il PDF, installala con: pip install weasyprint")
             print("Potrebbero essere necessarie dipendenze aggiuntive (pango, cairo, etc.). Vedi la documentazione di WeasyPrint.")
             # Prova ad installarla
             print("Tentativo di installazione di weasyprint...")
             import subprocess
             try:
                 subprocess.check_call(["pip", "install", "weasyprint"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                 print("WeasyPrint installato con successo.")
                 # Ricarica il modulo dopo l'installazione
                 importlib.invalidate_caches()
                 from weasyprint import HTML
             except subprocess.CalledProcessError as install_err:
                 print(f"Installazione di weasyprint fallita: {install_err}")
                 print("Controlla i messaggi di errore sopra e la documentazione di WeasyPrint per le dipendenze.")
                 return False
             except Exception as dynamic_import_err:
                 print(f"Errore durante l'importazione dinamica di WeasyPrint: {dynamic_import_err}")
                 return False
        else:
             from weasyprint import HTML

        print(f"Conversione di {file_html} in PDF...")
        # Usa un percorso assoluto per evitare problemi
        abs_html_path = os.path.abspath(file_html)
        HTML(filename=abs_html_path).write_pdf(output_pdf)
        print(f"File PDF creato con successo: {output_pdf}")
        return True

    except ImportError:
        # Questo blocco è ridondante a causa del controllo sopra, ma lo lascio per sicurezza
        print("\nATTENZIONE: La libreria 'weasyprint' non è installata.")
        return False
    except Exception as e:
        print(f"Errore durante la conversione in PDF: {e}")
        print("Assicurati che le dipendenze di WeasyPrint (come Pango, Cairo) siano installate sul tuo sistema.")
        return False

if __name__ == "__main__":
    # Percorsi dei file
    file_excel = "RH_Corso_Venezia_Offerta_Pulita.xlsx"
    output_html = "RH_Corso_Venezia_Offerta.html"
    output_pdf = "RH_Corso_Venezia_Offerta.pdf"

    if not os.path.exists(file_excel):
        print(f"Errore: Il file Excel '{file_excel}' non esiste nella directory corrente.")
        print(f"Directory corrente: {os.getcwd()}")
    else:
        # Converti Excel in HTML
        html_file_path = converti_excel_in_html(file_excel, output_html)

        if html_file_path:
            # Apri il file HTML nel browser
            try:
                webbrowser.open('file://' + os.path.abspath(html_file_path))
            except Exception as e:
                print(f"Non è stato possibile aprire automaticamente il browser: {e}")
                print(f"Puoi aprire manualmente il file: {os.path.abspath(html_file_path)}")

            # Chiedi all'utente se vuole convertire in PDF
            while True:
                try:
                    risposta = input("Vuoi convertire il file HTML in PDF? (s/n): ").lower().strip()
                    if risposta in ['s', 'n']:
                        break
                    else:
                        print("Risposta non valida. Inserisci 's' per sì o 'n' per no.")
                except EOFError:
                    print("\nInput interrotto. Uscita.")
                    risposta = 'n' # Tratta come 'no' se l'input viene interrotto
                    break

            if risposta == 's':
                pdf_success = converti_html_in_pdf(output_html, output_pdf)
                if not pdf_success:
                     print("\nLa conversione PDF non è riuscita. Puoi comunque stampare l'HTML in PDF dal tuo browser (File -> Stampa -> Salva come PDF).")
            else:
                print("Conversione PDF saltata.")
