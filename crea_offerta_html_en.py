import pandas as pd
import os
import jinja2
import webbrowser
from datetime import datetime
import sys # Aggiunto per sys.exit
from deep_translator import GoogleTranslator # Aggiunto import
import re # Per controllo eccezioni

# Importa WeasyPrint se presente (gestione errore se manca)
try:
    from weasyprint import HTML, CSS
    WEASYPRINT_AVAILABLE = True
except ImportError:
    WEASYPRINT_AVAILABLE = False
    print("Attenzione: libreria WeasyPrint non trovata. La conversione PDF non sarà disponibile.")
    print("Per abilitarla, installa WeasyPrint: pip install weasyprint")

# --- Funzione di Traduzione Dinamica (con eccezioni) ---
def translate_dynamic_text(translator, text, exceptions=None):
    """Traduci il testo usando l'istanza del traduttore fornita, saltando se contiene eccezioni."""
    if not text or not isinstance(text, str):
        return text # Ritorna il testo originale se vuoto o non stringa

    if exceptions:
        # Controlla se una qualsiasi delle eccezioni (case-insensitive) è nel testo
        if any(re.search(r'\b' + re.escape(exc) + r'\b', text, re.IGNORECASE) for exc in exceptions):
            # print(f"Skipping translation for '{text}' due to exception.") # Debug
            return text # Ritorna originale se trova eccezione

    try:
        # print(f"Translating: '{text}'") # Debug
        # Usa l'istanza del traduttore passata
        translated = translator.translate(text)
        # print(f"Translated: '{translated}'") # Debug
        print(".", end="", flush=True) # <<< INDICATORE DI PROGRESSO
        return translated if translated else text # Ritorna originale se la traduzione è vuota
    except Exception as e:
        print(f"\nError translating text '{text}': {e}. Returning original.")
        return text

# ----------------------------------------------------

def converti_excel_in_html_en(file_excel, output_html_en):
    """
    Converts an Excel file into an elegant HTML document (English Version).
    """
    print(f"Converting {file_excel} to HTML (English Version)...")

    try:
        # Try reading specific sheets
        xls = pd.ExcelFile(file_excel)
        if 'Informazioni' in xls.sheet_names and 'Offerta' in xls.sheet_names:
            info_df = pd.read_excel(xls, sheet_name='Informazioni')
            offerta_df = pd.read_excel(xls, sheet_name='Offerta')
        else:
            # If specific sheets aren't found, read the first sheet
            print("Sheets 'Informazioni' and 'Offerta' not found. Reading the first sheet...")
            offerta_df = pd.read_excel(xls)
            # Try to extract info if columns exist
            if 'Chiave' in offerta_df.columns and 'Valore' in offerta_df.columns:
                info_mask = offerta_df['Chiave'].notna() & offerta_df['Valore'].notna()
                info_df = offerta_df[info_mask][['Chiave', 'Valore']].copy()
                offerta_df = offerta_df[~info_mask].copy()
                if offerta_df.empty and not info_df.empty:
                    offerta_df = info_df.copy() # Use info as offer if it's the only thing present
                    info_df = pd.DataFrame(columns=['Chiave', 'Valore'])
            else:
                info_df = pd.DataFrame(columns=['Chiave', 'Valore']) # Create empty info DataFrame

    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

    # Extract offer information from info_df
    info_dict = {}
    if not info_df.empty:
        for _, row in info_df.iterrows():
            if pd.notna(row['Chiave']) and pd.notna(row['Valore']):
                key = str(row['Chiave']).strip(':') if isinstance(row['Chiave'], str) else row['Chiave']
                value = row['Valore']
                if isinstance(value, pd.Timestamp):
                    value = value.strftime('%d/%m/%Y') # Keep date format?
                info_dict[key] = value

    # --- EDIT: Set customer manually --- (Keep internal logic names)
    info_dict['Cliente'] = 'RH' # Or translate 'Cliente' key if needed for template?
    # Let's assume the key 'Cliente' is used in the template and should remain.
    # The value 'RH' probably doesn't need translation.
    # ------------------------------------

    # Prepare offer data
    offerta_data = []
    headers = []
    prezzo_netto_key = 'Prezzo netto' # Keep original key name from Excel
    quantita_key = 'Quantità'        # Keep original key name from Excel
    descrizione_key = 'Descrizione' # <<< CERCA QUESTA CHIAVE SPECIFICA
    exceptions_list = ['marrone']    # Parole da non tradurre
    translator = None                # Inizializza a None

    # --- ISTANZIA TRADUTTORE UNA VOLTA (PRIMA DI USARLO!) --- 
    print("Initializing translator...")
    try:
        translator = GoogleTranslator(source='it', target='en')
        print("Translator initialized.")
    except Exception as e:
        print(f"Error initializing translator: {e}. Translations might be skipped.")
    # -------------------------------------

    if not offerta_df.empty:
        headers = offerta_df.columns.tolist()
        # --- Controlla se la chiave descrizione esiste --- 
        if descrizione_key not in headers:
            print(f"Attenzione: Chiave '{descrizione_key}' non trovata negli header: {headers}. La traduzione delle descrizioni potrebbe non funzionare.")
            # Potremmo provare a dedurla o fermarci, per ora avvisiamo soltanto.
        # -----------------------------------------------

        # --- Traduci Headers --- (Solo per display, USARE IL TRADUTTORE ISTANZIATO)
        headers_map = {}
        if translator: # Traduci solo se il traduttore è stato inizializzato con successo
            # print("Translating headers...") # TEMPORANEAMENTE DISABILITATO
            # headers_map = {hdr: translate_dynamic_text(translator, hdr, exceptions=exceptions_list) for hdr in headers} # <<< PASSA translator e hdr
            # print() # A capo dopo i punti degli header
            print("Skipping header translation (temporarily disabled).")
            headers_map = {hdr: hdr for hdr in headers} # Usa header originali
        else:
            print("Skipping header translation due to translator initialization error.")
            headers_map = {hdr: hdr for hdr in headers} # Usa header originali se il traduttore non è pronto
        # -----------------------

        offerta_df_filled = offerta_df.fillna('')
        offerta_data_list = offerta_df_filled.to_dict('records')

        # --- Processa Righe --- 
        # print("Starting row processing...") # TEMPORANEAMENTE DISABILITATO
        offerta_data = []
        if headers:
            first_col_key = headers[0] # Chiave della prima colonna (usata SOLO per categoria)
            other_col_keys = headers[1:]
            if prezzo_netto_key not in headers:
                print(f"Warning: Column '{prezzo_netto_key}' not found in headers: {headers}")

            for row_dict in offerta_data_list:
                # 1. Determine if it's a category row (USA ANCORA LA PRIMA COLONNA PER QUESTO)
                is_category = False
                first_val = row_dict.get(first_col_key, '')
                if first_val != '':
                    all_others_empty = all(row_dict.get(col_key, '') == '' for col_key in other_col_keys)
                    if all_others_empty:
                        is_category = True
                row_dict['is_category'] = is_category

                # --- Traduci Descrizione Articolo (se non categoria e chiave ESISTE e traduttore OK) ---
                # if not is_category and descrizione_key in row_dict and translator: # TEMPORANEAMENTE DISABILITATO
                #     original_description = row_dict[descrizione_key]
                #     row_dict[descrizione_key] = translate_dynamic_text(translator, original_description, exceptions=exceptions_list) # <<< PASSA IL TRADUTTORE
                # else:
                #     # Mantieni la descrizione originale se la traduzione è disabilitata o fallisce
                #     pass # Non fare nulla, usa il valore originale già presente in row_dict
                pass # Traduzione disabilitata
                # -------------------------------------------------------------------

                # 2. Add Euro symbol and format price
                original_value = row_dict.get(prezzo_netto_key, '')
                formatted_prezzo = original_value
                val_float = None

                if not is_category and prezzo_netto_key in row_dict and original_value != '':
                    try:
                        if isinstance(original_value, (int, float)):
                            val_float = float(original_value)
                        elif isinstance(original_value, str):
                            val_str = original_value.strip().replace(' ', '')
                            if val_str:
                                has_comma = ',' in val_str
                                has_dot = '.' in val_str
                                cleaned_str_for_float = val_str
                                if has_comma and has_dot:
                                    cleaned_str_for_float = val_str.replace('.', '').replace(',', '.')
                                elif has_comma:
                                    cleaned_str_for_float = val_str.replace(',', '.')
                                elif has_dot:
                                    cleaned_str_for_float = val_str
                                val_float = float(cleaned_str_for_float)
                    except (ValueError, TypeError):
                        print(f"Warning: Could not convert '{original_value}' to number. Using original value for display.")
                        formatted_prezzo = f"€ {str(original_value)}"

                    if val_float is not None:
                        try:
                            # Format with two decimals for display (using dot as decimal separator for EN potentially)
                            # Let's keep the Italian formatting for numbers for now, as requested previously.
                            formatted_prezzo = f"€ {val_float:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                        except Exception as e_format:
                            print(f"Error formatting value {val_float}: {e_format}. Using original value.")
                            formatted_prezzo = f"€ {str(original_value)}"

                row_dict['formatted_prezzo'] = formatted_prezzo
                offerta_data.append(row_dict)
        else:
            offerta_data = offerta_data_list
    else:
        print("Warning: No data found for the offer table.")

    # --- Pulisce l'output dell'indicatore di progresso ---
    print() # Va a capo dopo i punti stampati
    # --------------------------------------------------

    # HTML Template (English)
    template_str = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{{ info.get('Riferimento', 'Commercial Offer') }}</title> {# Translated default #}
        <style>
            /* CSS Styles (keep as is, comments can remain in Italian or be removed/translated) */
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
                color: #0056b3; /* Corporate Blue */
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
                min-width: 150px; /* Adjust if needed for English labels */
                padding-right: 10px;
            }
            .info-value {
                 color: #343a40;
            }
            table {
                width: 100%;
                border-collapse: collapse;
                margin-bottom: 30px;
                box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
                background-color: #fff;
                border-radius: 4px;
                overflow: hidden;
            }
            th, td {
                border: 1px solid #dee2e6;
                padding: 12px 15px;
                text-align: left;
                vertical-align: top;
            }
            th {
                background-color: #e9ecef;
                color: #495057;
                font-weight: 600; /* SemiBold */
                text-transform: uppercase;
                font-size: 0.9em;
                letter-spacing: 0.5px;
            }
            tr:nth-child(even) {
                background-color: #f8f9fa;
            }
            tr:hover {
                background-color: #f1f3f5;
            }
            .category {
                font-weight: bold;
                background-color: #dbe4ff !important; /* Light blue background for category */
                color: #0056b3;
                text-align: left !important; /* Ensure left alignment */
            }
            .category td {
                 /* Apply to all cells in category row if needed */
            }
            .footer {
                text-align: center;
                margin-top: 40px;
                padding-top: 20px;
                border-top: 1px solid #dee2e6;
                font-size: 0.85em;
                color: #6c757d;
            }
            @media print {
                body { background-color: white; margin: 0; }
                .container { box-shadow: none; border: none; margin: 0; max-width: 100%; padding: 0; }
                .info-section { background-color: #f1f3f5 !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                tr:nth-child(even) { background-color: #f8f9fa !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                .category { background-color: #dbe4ff !important; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                .footer { border-top: 1px solid #ccc; }
                h1, h2 { border-bottom: 1px solid #ccc; }
                .header-section { border-bottom: 1px solid #dee2e6; }
            }

            /* --- Header Styles --- */
            .header-section {
                display: flex;
                justify-content: space-between;
                align-items: flex-start;
                margin-bottom: 40px;
                padding-bottom: 20px;
                border-bottom: 1px solid #dee2e6;
                flex-wrap: wrap;
            }
            .logo-container {
                flex: 0 0 auto;
                margin-right: 30px;
                margin-bottom: 15px;
            }
            .logo-container img {
                max-width: 200px;
                height: auto;
            }
            .company-details {
                flex: 1 1 auto;
                text-align: right;
                font-size: 0.9em;
                line-height: 1.4;
                color: #495057;
                margin-bottom: 15px;
            }
            .company-details p {
                margin: 2px 0;
            }
            .company-details strong {
                color: #212529;
                font-weight: 600;
            }
            /* ------------------------- */

            /* --- Table Column Styles --- */
            .col-quantita {
                width: 6%;
                max-width: 50px;
            }
            .col-prezzo-netto {
                width: 18%;
            }
            th.col-quantita, td.col-quantita {
                text-align: center;
            }
            th.col-prezzo-netto, td.col-prezzo-netto {
                text-align: right;
            }
            /* --------------------------- */

            /* --- Clauses Styles --- */
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
            <!-- === Header Section === -->
            <div class="header-section">
                <div class="logo-container">
                    <img src="https://italfrigo.github.io/marrone-iglu-italfrigo-presentation/Presentation/assets/logos/italfrigo-logo.png" alt="Italfrigo Service Srl Logo"> {# Translated alt text #}
                </div>
                <div class="company-details">
                    <p><strong>ITALFRIGO SERVICE SRL</strong></p>
                    <p>Via Pelizza Da Volpedo 49/A</p>
                    <p>20092 – Cinisello Balsamo (MI), Italy</p> {# Translated Country #}
                    <p>Tel: +39 0266040933</p>
                    <p>VAT ID: 06825390153</p> {# Translated P.IVA #}
                    <br>
                    <p><em>Contact:</em> Manuel Lazzaro</p>
                    <p><em>Email:</em> manuel.lazzaro@italfrigo.com</p>
                    <p><em>Mobile:</em> +39 3461587689</p>
                </div>
            </div>
            <!-- ============================ -->

            <h1>{{ info.get('Riferimento', 'Commercial Offer') }}</h1> {# Title potentially from info_dict or default translated #}

            {% if info %}
            <h2>General Information</h2> {# Translated Section Title #}
            <div class="info-section">
                {% for key, value in info.items() %}
                <div class="info-item">
                    {# Translate known keys, otherwise display original key #}
                    {% set display_key = key %}
                    {% if key == 'Cliente' %}{% set display_key = 'Customer' %}
                    {% elif key == 'Data' %}{% set display_key = 'Date' %}
                    {% elif key == 'Riferimento' %}{% set display_key = 'Reference' %}
                    {% elif key == 'Oggetto' %}{% set display_key = 'Subject' %}
                    {% endif %}
                    <span class="info-label">{{ display_key }}:</span>
                    <span class="info-value">{{ value }}</span>
                </div>
                {% endfor %}
            </div>
            {% endif %}

            <h2>Offer Details</h2> {# Translated Section Title #}
            {% if headers and offerta_data %}
            <table>
                <thead>
                    <tr>
                        {% for header in headers %}
                        {# Usa l'header tradotto dalla mappa per il display #}
                        <th {% if header == quantita_key %}class="col-quantita"{% elif header == prezzo_netto_key %}class="col-prezzo-netto"{% endif %}>{{ headers_map.get(header, header) }}</th>
                        {% endfor %}
                    </tr>
                </thead>
                <tbody>
                    {% for row in offerta_data %}
                    <tr {% if row.is_category %}class="category"{% endif %}>
                        {% for header in headers %}
                         {# Access data using header as key, use formatted price if it's the right column #}
                        <td {% if header == quantita_key %}class="col-quantita"{% elif header == prezzo_netto_key %}class="col-prezzo-netto"{% endif %}>{% if header == prezzo_netto_key %}{{ row.formatted_prezzo }}{% else %}{{ row.get(header, '') }}{% endif %}</td>
                        {% endfor %}
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
            {% else %}
            <p>No offer details available in the Excel file.</p> {# Translated message #}
            {% endif %}

            <!-- === Commercial Clauses Section === -->
            <div class="clauses-section">
                <h3>Commercial Clauses</h3> {# Translated Title #}
                <p><strong>Included:</strong></p> {# Translated Label #}
                <ul>
                    <li>Transport</li> {# Translated Item #}
                    <li>Assembly</li> {# Translated Item #}
                    <li>Testing</li> {# Translated Item #}
                </ul>
                <p><strong>Excluded:</strong></p> {# Translated Label #}
                <ul>
                    <li>Electrical connections</li> {# Translated Item #}
                    <li>Hydraulic connections</li> {# Translated Item #}
                    <li>Any masonry work</li> {# Translated Item #}
                    <li>Anything not expressly mentioned in this offer</li> {# Translated Item #}
                </ul>
                <p><strong>Payment:</strong> To be agreed.</p> {# Translated #}
                <p><strong>Delivery Date:</strong> To be agreed.</p> {# Translated #}
                <br>
                <p><em><strong>Note:</strong> All prices are excluding VAT 22%.</em></p> {# Translated #}
            </div>
            <!-- ==================================== -->

            <div class="footer">
                Document generated on {{ data_generazione }} {# Keep date format, translate text #}
            </div>
        </div>
    </body>
    </html>
    """

    # Render the template
    env = jinja2.Environment(loader=jinja2.BaseLoader())
    template = env.from_string(template_str)
    html_content = template.render(
        info=info_dict,
        headers=headers, # Passa gli header originali per la logica interna del loop dati
        headers_map=headers_map, # Passa la mappa per il display degli header tradotti
        offerta_data=offerta_data,
        prezzo_netto_key=prezzo_netto_key,
        quantita_key=quantita_key,
        data_generazione=datetime.now().strftime('%d/%m/%Y %H:%M:%S') # Keep format
    )

    # Write the HTML file
    try:
        with open(output_html_en, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print(f"English HTML file created successfully: {output_html_en}")
        return output_html_en # Return the path for potential PDF conversion
    except Exception as e:
        print(f"Error writing HTML file: {e}")
        return None

def converti_html_in_pdf(input_html, output_pdf):
    """Converts an HTML file to PDF using WeasyPrint."""
    if not WEASYPRINT_AVAILABLE:
        print("PDF conversion skipped as WeasyPrint is not available.")
        return False

    print(f"Converting {input_html} to PDF...")
    try:
        # Base URL per risolvere percorsi relativi (es. immagini)
        base_url = os.path.dirname(os.path.abspath(input_html))
        HTML(filename=input_html, base_url=base_url).write_pdf(output_pdf)
        print(f"PDF file created successfully: {output_pdf}")
        return True
    except Exception as e:
        print(f"Error during PDF conversion: {e}")
        return False

# Esecuzione principale
if __name__ == "__main__":
    file_excel_input = 'RH_Corso_Venezia_Offerta_Pulita.xlsx'
    file_html_output_en = 'RH_Corso_Venezia_Offerta_EN.html' # English output name
    file_pdf_output_en = 'RH_Corso_Venezia_Offerta_EN.pdf'   # English PDF output name

    generated_html_en = converti_excel_in_html_en(file_excel_input, file_html_output_en)

    if generated_html_en:
        # Ask if user wants to convert the ENGLISH HTML to PDF
        if WEASYPRINT_AVAILABLE:
            while True:
                convert_pdf = input(f"Do you want to convert the English HTML file ({file_html_output_en}) to PDF? (y/n): ").lower()
                if convert_pdf == 'y':
                    converti_html_in_pdf(generated_html_en, file_pdf_output_en)
                    break
                elif convert_pdf == 'n':
                    print("PDF conversion skipped.")
                    break
                else:
                    print("Invalid input. Please enter 'y' or 'n'.")
        # Optionally open the generated ENGLISH HTML file
        # try:
        #     webbrowser.open(f'file://{os.path.abspath(generated_html_en)}')
        # except Exception as e:
        #     print(f"Could not open HTML file in browser: {e}")
    else:
        print("English HTML generation failed.")
        sys.exit(1) # Exit with error code

    print("Script finished.")
