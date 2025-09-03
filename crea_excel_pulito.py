import pandas as pd
import os
import openpyxl
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def crea_excel_pulito(input_file, output_file):
    print(f"Creazione di un Excel pulito da {input_file}...")
    
    # Leggi il file Excel
    df = pd.read_excel(input_file, header=None)
    
    # Trova la riga delle intestazioni
    header_row_idx = None
    for i, row in df.iterrows():
        if pd.notna(row[0]) and isinstance(row[0], str) and "codice" in str(row[0]).strip().lower():
            header_row_idx = i
            break
        # Controlla anche se c'è una cella con "descrizione"
        for cell in row:
            if pd.notna(cell) and isinstance(cell, str) and "descrizione" in str(cell).strip().lower():
                header_row_idx = i
                break
        if header_row_idx is not None:
            break
    
    if header_row_idx is None:
        print("Errore: Non è stata trovata la riga delle intestazioni.")
        return
    
    # Estrai le intestazioni
    headers = []
    for i, value in enumerate(df.iloc[header_row_idx]):
        if pd.notna(value):
            # Rinomina "Vendita al cliente" in "Prezzo netto"
            if isinstance(value, str) and "Vendita al cliente" in value:
                headers.append("Prezzo netto")
            # Salta la colonna "Costo italfrigo"
            elif isinstance(value, str) and "Costo italfrigo" in value:
                continue
            else:
                headers.append(value.strip(':') if isinstance(value, str) and value.endswith(':') else value)
        else:
            continue  # Salta colonne vuote
    
    # Estrai i dati effettivi (righe dopo le intestazioni)
    data_rows = []
    current_category = None
    
    for i in range(header_row_idx + 1, len(df)):
        row = df.iloc[i]
        
        # Se la prima colonna ha un valore e le altre sono vuote, è una categoria
        if pd.notna(row[0]) and all(pd.isna(row[j]) for j in range(1, len(row))):
            current_category = row[0]
            new_row = [current_category]
            new_row.extend([''] * (len(headers) - 1))
            data_rows.append(new_row)
        # Altrimenti è un elemento della categoria corrente
        elif any(pd.notna(val) for val in row):
            row_data = []
            for j in range(len(df.columns)):
                # Salta la colonna "Costo italfrigo"
                if j == 3:  # Indice della colonna "Costo italfrigo"
                    continue
                if j < len(row):
                    row_data.append(row[j] if pd.notna(row[j]) else '')
                else:
                    break  # Non aggiungere colonne vuote alla fine
            
            # Assicurati che la riga abbia la lunghezza giusta
            while len(row_data) < len(headers):
                row_data.append('')
            
            # Taglia eventuali colonne in eccesso
            if len(row_data) > len(headers):
                row_data = row_data[:len(headers)]
                
            data_rows.append(row_data)
    
    # Crea un nuovo DataFrame con i dati puliti
    df_clean = pd.DataFrame(data_rows, columns=headers)
    
    # Estrai le informazioni dell'offerta (prime righe)
    info_offerta = {}
    for i in range(header_row_idx):
        row = df.iloc[i]
        if pd.notna(row[0]) and pd.notna(row[1]):
            key = str(row[0]).strip(':') if isinstance(row[0], str) and row[0].endswith(':') else row[0]
            info_offerta[key] = row[1]
    
    # Crea un foglio Excel con i dati puliti
    # Creiamo prima un file XLSX
    temp_xlsx = output_file.replace('.xls', '.xlsx')
    
    with pd.ExcelWriter(temp_xlsx, engine='openpyxl') as writer:
        # Scrivi le informazioni dell'offerta
        info_df = pd.DataFrame(list(info_offerta.items()), columns=['Chiave', 'Valore'])
        info_df.to_excel(writer, sheet_name='Informazioni', index=False)
        
        # Scrivi i dati principali
        df_clean.to_excel(writer, sheet_name='Offerta', index=False)
    
    print(f"File temporaneo XLSX creato: {temp_xlsx}")
    print(f"Per usare il formato XLS, apri il file {temp_xlsx} con Excel e salvalo come XLS manualmente.")
    print(f"Il formato XLS è un formato legacy e potrebbe non supportare tutte le funzionalità.")
    
    print(f"Excel pulito creato con successo: {output_file}")

if __name__ == "__main__":
    input_file = "rh_conti.xls"
    output_file = "RH_Corso_Venezia_Offerta_Pulita.xls"
    
    if not os.path.exists(input_file):
        print(f"Errore: Il file {input_file} non esiste.")
    else:
        crea_excel_pulito(input_file, output_file)
