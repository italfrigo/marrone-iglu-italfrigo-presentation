import pandas as pd
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import datetime
import locale
import os

# Imposta la localizzazione italiana per il formato delle date e numeri
locale.setlocale(locale.LC_ALL, 'it_IT.UTF-8')

# Percorso del file Excel originale
file_originale = 'rh_conti.xls'
output_pdf = 'RH_Corso_Venezia_Offerta.pdf'

# Percorso del logo
logo_path = 'italfrigo/logo_italfrigo.png'

# Colori per il documento
COLORE_PRIMARIO = colors.HexColor('#0066CC')  # Blu
COLORE_SECONDARIO = colors.HexColor('#FF9900')  # Arancione
COLORE_SFONDO = colors.HexColor('#F2F2F2')  # Grigio chiaro
COLORE_TESTO = colors.HexColor('#333333')  # Grigio scuro

def crea_offerta_pdf():
    print(f"Creazione dell'offerta in PDF da {file_originale}...")
    
    # Carica i dati dal file Excel
    df = pd.read_excel(file_originale)
    
    # Estrai le informazioni di intestazione
    info_offerta = {}
    for i in range(5):
        if not pd.isna(df.iloc[i, 0]) and not pd.isna(df.iloc[i, 1]):
            chiave = df.iloc[i, 0].strip(':') if isinstance(df.iloc[i, 0], str) and df.iloc[i, 0].endswith(':') else df.iloc[i, 0]
            info_offerta[chiave] = df.iloc[i, 1]
    
    # Modifica il cliente in "RH"
    if 'Cliente' in info_offerta:
        info_offerta['Cliente'] = "RH"
        
    # Formatta le date per mostrare solo il giorno
    if 'Data' in info_offerta and isinstance(info_offerta['Data'], pd.Timestamp):
        info_offerta['Data'] = info_offerta['Data'].strftime("%d/%m/%Y")
        
    if 'Validità' in info_offerta and isinstance(info_offerta['Validità'], pd.Timestamp):
        info_offerta['Validità'] = info_offerta['Validità'].strftime("%d/%m/%Y")
    
    # Trova l'indice della riga con le intestazioni delle colonne
    header_row_idx = None
    for i in range(len(df)):
        if isinstance(df.iloc[i, 0], str) and 'Codice' in df.iloc[i, 0]:
            header_row_idx = i
            break
    
    if header_row_idx is None:
        print("Non è stato possibile trovare le intestazioni delle colonne.")
        return
    
    # Estrai le intestazioni delle colonne
    headers = []
    for col in df.columns:
        value = df.iloc[header_row_idx, df.columns.get_loc(col)]
        if pd.notna(value):
            # Rinomina "Vendita al cliente" in "Prezzo netto"
            if isinstance(value, str) and "Vendita al cliente" in value:
                headers.append("Prezzo netto")
            # Salta la colonna "Costo italfrigo"
            elif isinstance(value, str) and "Costo italfrigo" in value:
                headers.append("SKIP_COLUMN")
            else:
                headers.append(value.strip(':') if isinstance(value, str) and value.endswith(':') else value)
        else:
            headers.append('')
    
    # Rimuovi l'intestazione "SKIP_COLUMN" dalla lista
    skip_column_idx = headers.index("SKIP_COLUMN") if "SKIP_COLUMN" in headers else -1
    if skip_column_idx != -1:
        headers.pop(skip_column_idx)
    
    # Estrai i dati effettivi (righe dopo le intestazioni)
    data_rows = []
    current_category = None
    
    for i in range(header_row_idx + 1, len(df)):
        row = df.iloc[i]
        
        # Se la prima colonna ha un valore e le altre sono vuote, è una categoria
        if pd.notna(row[0]) and all(pd.isna(row[j]) for j in range(1, len(row))):
            current_category = row[0]
            data_rows.append([current_category, '', '', '', ''])
        # Altrimenti è un elemento della categoria corrente
        elif any(pd.notna(val) for val in row):
            row_data = []
            for j in range(len(row)):
                row_data.append(row[j] if pd.notna(row[j]) else '')
            data_rows.append(row_data)
    
    # Crea il PDF
    doc = SimpleDocTemplate(
        output_pdf,
        pagesize=landscape(A4),
        leftMargin=1*cm,
        rightMargin=1*cm,
        topMargin=1*cm,
        bottomMargin=1*cm
    )
    
    # Stili per il testo
    styles = getSampleStyleSheet()
    
    # Crea stili personalizzati
    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor=COLORE_PRIMARIO,
        alignment=1,  # Centrato
        spaceAfter=12
    )
    
    subtitle_style = ParagraphStyle(
        'SubtitleStyle',
        parent=styles['Heading2'],
        fontSize=18,
        textColor=COLORE_SECONDARIO,
        alignment=1,  # Centrato
        spaceAfter=12
    )
    
    info_label_style = ParagraphStyle(
        'InfoLabelStyle',
        parent=styles['Normal'],
        fontSize=10,
        textColor=COLORE_TESTO,
        fontName='Helvetica-Bold'
    )
    
    info_value_style = ParagraphStyle(
        'InfoValueStyle',
        parent=styles['Normal'],
        fontSize=10,
        textColor=COLORE_TESTO
    )
    
    header_style = ParagraphStyle(
        'HeaderStyle',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.white,
        fontName='Helvetica-Bold',
        alignment=1  # Centrato
    )
    
    normal_style = ParagraphStyle(
        'NormalStyle',
        parent=styles['Normal'],
        fontSize=9,
        textColor=COLORE_TESTO
    )
    
    category_style = ParagraphStyle(
        'CategoryStyle',
        parent=styles['Normal'],
        fontSize=10,
        textColor=COLORE_PRIMARIO,
        fontName='Helvetica-Bold'
    )
    
    # Elementi del documento
    elements = []
    
    # Spacer iniziale
    elements.append(Spacer(1, 30*mm))
    
    # Sottotitolo (riferimento)
    elements.append(Paragraph("RH - CORSO VENEZIA - MILANO", subtitle_style))
    
    elements.append(Spacer(1, 20*mm))
    
    # Tabella con le informazioni dell'offerta
    info_data = []
    
    # Prima riga: Quotazione e Data
    row1 = []
    if 'Quotazione' in info_offerta:
        row1.extend([
            Paragraph("<b>Quotazione:</b>", info_label_style),
            Paragraph(str(info_offerta['Quotazione']), info_value_style)
        ])
    else:
        row1.extend([Paragraph("", normal_style), Paragraph("", normal_style)])
    
    row1.append(Paragraph("", normal_style))  # Spaziatore
    
    if 'Data' in info_offerta:
        # Converti la data in formato leggibile (solo giorno)
        data_str = info_offerta['Data']
        if isinstance(data_str, pd.Timestamp):
            data_str = data_str.strftime("%d/%m/%Y")
        row1.extend([
            Paragraph("<b>Data:</b>", info_label_style),
            Paragraph(str(data_str), info_value_style)
        ])
    else:
        row1.extend([Paragraph("", normal_style), Paragraph("", normal_style)])
    
    # Seconda riga: Cliente e Validità
    row2 = []
    if 'Cliente' in info_offerta:
        row2.extend([
            Paragraph("<b>Cliente:</b>", info_label_style),
            Paragraph(str(info_offerta['Cliente']), info_value_style)
        ])
    else:
        row2.extend([Paragraph("", normal_style), Paragraph("", normal_style)])
    
    row2.append(Paragraph("", normal_style))  # Spaziatore
    
    if 'Validità' in info_offerta:
        # Converti la data in formato leggibile (solo giorno)
        validita_str = info_offerta['Validità']
        if isinstance(validita_str, pd.Timestamp):
            validita_str = validita_str.strftime("%d/%m/%Y")
        row2.extend([
            Paragraph("<b>Validità:</b>", info_label_style),
            Paragraph(str(validita_str), info_value_style)
        ])
    else:
        row2.extend([Paragraph("", normal_style), Paragraph("", normal_style)])
    
    info_data.append(row1)
    info_data.append(row2)
    
    # Crea la tabella delle informazioni
    info_table = Table(info_data, colWidths=[3*cm, 7*cm, 2*cm, 3*cm, 7*cm])
    info_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    
    elements.append(info_table)
    elements.append(Spacer(1, 15*mm))
    
    # Nota di validità
    elements.append(Paragraph("Offerta valida per 60 giorni dalla data di emissione", ParagraphStyle(
        'ValiditaStyle',
        parent=styles['Normal'],
        fontSize=9,
        textColor=COLORE_TESTO,
        alignment=1  # Centrato
    )))
    elements.append(Spacer(1, 5*mm))
    
    # Contatti
    elements.append(Paragraph("Per informazioni:", ParagraphStyle(
        'ContattiTitleStyle',
        parent=styles['Normal'],
        fontSize=9,
        fontName='Helvetica-Bold',
        textColor=COLORE_TESTO,
        alignment=1  # Centrato
    )))
    elements.append(Paragraph("info@italfrigo.com | +39 0266040933", ParagraphStyle(
        'ContattiStyle',
        parent=styles['Normal'],
        fontSize=9,
        textColor=COLORE_TESTO,
        alignment=1  # Centrato
    )))
    elements.append(Paragraph("Referente: Manuel Lazzaro | manuel.lazzaro@italfrigo.com | +39 3461587689", ParagraphStyle(
        'ContattiStyle',
        parent=styles['Normal'],
        fontSize=9,
        textColor=COLORE_TESTO,
        alignment=1  # Centrato
    )))
    elements.append(Spacer(1, 15*mm))
    
    # Tabella con i dati dell'offerta
    # Prepara le intestazioni come Paragraphs
    header_paragraphs = []
    for h in headers:
        if h != "SKIP_COLUMN":
            header_paragraphs.append(Paragraph(h, header_style))
    
    # Prepara i dati
    table_data = [header_paragraphs]
    
    # Definisci quali colonne mantenere (0=Codice, 1=Descrizione, 2=Quantità, 4=Prezzo netto)
    # Saltiamo la colonna 3 (Costo italfrigo) e qualsiasi colonna oltre la 4
    columns_to_keep = [0, 1, 2, 4]
    
    for row in data_rows:
        # Controlla se è una riga di categoria
        is_category = pd.notna(row[0]) and all(pd.isna(row[j]) for j in range(1, len(row)))
        
        table_row = []
        
        if is_category:
            # Per le righe di categoria, aggiungi solo la prima cella e celle vuote per le altre colonne
            table_row.append(Paragraph(str(row[0]), category_style))
            for _ in range(len(columns_to_keep) - 1):
                table_row.append(Paragraph("", normal_style))
        else:
            # Per le righe normali, aggiungi solo le colonne che vogliamo mantenere
            for i in columns_to_keep:
                if i < len(row):
                    cell = row[i]
                    # Formatta i numeri
                    if i == 2 and isinstance(cell, (int, float)):  # Quantità
                        cell_text = str(int(cell)) if cell == int(cell) else f"{cell:.2f}"
                        table_row.append(Paragraph(cell_text, normal_style))
                    elif i == 4 and isinstance(cell, (int, float)):  # Prezzo
                        cell_text = f"{cell:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
                        table_row.append(Paragraph(cell_text, normal_style))
                    else:
                        table_row.append(Paragraph(str(cell) if pd.notna(cell) else "", normal_style))
                else:
                    table_row.append(Paragraph("", normal_style))
        
        table_data.append(table_row)
    
    # Crea la tabella
    col_widths = [4*cm, 14*cm, 2*cm, 4*cm]  # Larghezza colonne: Codice, Descrizione, Quantità, Prezzo
    data_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    
    # Stile della tabella
    table_style = [
        # Intestazioni
        ('BACKGROUND', (0, 0), (-1, 0), COLORE_PRIMARIO),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        
        # Bordi
        ('GRID', (0, 0), (-1, -1), 0.5, COLORE_TESTO),
        
        # Allineamento
        ('ALIGN', (2, 1), (4, -1), 'RIGHT'),  # Allinea a destra le colonne numeriche
        
        # Padding
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
    ]
    
    # Aggiungi stili per le righe di categoria
    for i, row in enumerate(data_rows):
        if pd.notna(row[0]) and all(pd.isna(val) if isinstance(val, float) else val == '' for val in row[1:]):
            table_style.append(('BACKGROUND', (0, i+1), (-1, i+1), COLORE_SFONDO))
            table_style.append(('SPAN', (0, i+1), (-1, i+1)))  # Unisci tutte le celle della riga
    
    # Aggiungi stile alternato per le righe normali
    for i in range(1, len(table_data)):
        if i % 2 == 0:
            table_style.append(('BACKGROUND', (0, i), (-1, i), colors.white))
        else:
            table_style.append(('BACKGROUND', (0, i), (-1, i), COLORE_SFONDO))
    
    data_table.setStyle(TableStyle(table_style))
    
    elements.append(data_table)
    
    # Aggiungi una pagina finale con le informazioni sui prezzi e condizioni
    elements.append(PageBreak())
    
    # Titolo della pagina finale
    elements.append(Paragraph("TERMINI E CONDIZIONI", ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=16,
        textColor=COLORE_PRIMARIO,
        alignment=1,  # Centrato
        spaceAfter=12
    )))
    elements.append(Spacer(1, 10*mm))
    
    # Informazioni sui prezzi
    elements.append(Paragraph("Informazioni sui prezzi", ParagraphStyle(
        'SubtitleStyle',
        parent=styles['Heading2'],
        fontSize=12,
        textColor=COLORE_SECONDARIO,
        fontName='Helvetica-Bold'
    )))
    elements.append(Spacer(1, 5*mm))
    
    elements.append(Paragraph("Tutti i prezzi sono da intendersi IVA esclusa. L'offerta è soggetta alle condizioni generali di vendita di Italfrigo Service SRL.", styles['Normal']))
    elements.append(Spacer(1, 10*mm))
    
    # Condizioni di pagamento
    elements.append(Paragraph("Condizioni di pagamento", ParagraphStyle(
        'SubtitleStyle',
        parent=styles['Heading2'],
        fontSize=12,
        textColor=COLORE_SECONDARIO,
        fontName='Helvetica-Bold'
    )))
    elements.append(Spacer(1, 5*mm))
    
    elements.append(Paragraph("Da concordare con il cliente", styles['Normal']))
    elements.append(Spacer(1, 10*mm))
    
    # Tempi di consegna
    elements.append(Paragraph("Tempi di consegna", ParagraphStyle(
        'SubtitleStyle',
        parent=styles['Heading2'],
        fontSize=12,
        textColor=COLORE_SECONDARIO,
        fontName='Helvetica-Bold'
    )))
    elements.append(Spacer(1, 5*mm))
    
    elements.append(Paragraph("Da concordare con il cliente.", styles['Normal']))
    elements.append(Spacer(1, 10*mm))
    
    # Inclusi ed esclusi
    elements.append(Paragraph("Inclusi ed esclusi", ParagraphStyle(
        'SubtitleStyle',
        parent=styles['Heading2'],
        fontSize=12,
        textColor=COLORE_SECONDARIO,
        fontName='Helvetica-Bold'
    )))
    elements.append(Spacer(1, 5*mm))
    
    elements.append(Paragraph("Inclusi:", ParagraphStyle('Bold', parent=styles['Normal'], fontName='Helvetica-Bold')))
    elements.append(Paragraph("- Trasporto", styles['Normal']))
    elements.append(Paragraph("- Montaggio", styles['Normal']))
    elements.append(Paragraph("- Collaudo", styles['Normal']))
    elements.append(Spacer(1, 5*mm))
    
    elements.append(Paragraph("Esclusi:", ParagraphStyle('Bold', parent=styles['Normal'], fontName='Helvetica-Bold')))
    elements.append(Paragraph("- Allacciamenti elettrici", styles['Normal']))
    elements.append(Paragraph("- Allacciamenti idraulici", styles['Normal']))
    elements.append(Paragraph("- Opere murarie", styles['Normal']))
    elements.append(Paragraph("- Tutto quanto non espressamente indicato nella presente offerta", styles['Normal']))
    elements.append(Spacer(1, 10*mm))
    
    # Validità dell'offerta
    elements.append(Paragraph("Validità dell'offerta", ParagraphStyle(
        'SubtitleStyle',
        parent=styles['Heading2'],
        fontSize=12,
        textColor=COLORE_SECONDARIO,
        fontName='Helvetica-Bold'
    )))
    elements.append(Spacer(1, 5*mm))
    
    elements.append(Paragraph("Offerta valida per 60 giorni dalla data di emissione.", styles['Normal']))
    elements.append(Spacer(1, 10*mm))
    
    # Contatti
    elements.append(Paragraph("Contatti", ParagraphStyle(
        'SubtitleStyle',
        parent=styles['Heading2'],
        fontSize=12,
        textColor=COLORE_SECONDARIO,
        fontName='Helvetica-Bold'
    )))
    elements.append(Spacer(1, 5*mm))
    
    elements.append(Paragraph("Per accettazione dell'offerta o per ulteriori informazioni, si prega di contattare:", styles['Normal']))
    elements.append(Spacer(1, 3*mm))
    elements.append(Paragraph("Italfrigo Service SRL", ParagraphStyle('Bold', parent=styles['Normal'], fontName='Helvetica-Bold')))
    elements.append(Paragraph("Via Pelizza Da Volpedo 49/A, 20092 – Cinisello Balsamo, Milano, Italia", styles['Normal']))
    elements.append(Paragraph("Tel: +39 0266040933 | Email: info@italfrigo.com", styles['Normal']))
    elements.append(Paragraph("P.IVA: 06825390153", styles['Normal']))
    elements.append(Spacer(1, 5*mm))
    elements.append(Paragraph("Referente: Manuel Lazzaro", ParagraphStyle('Bold', parent=styles['Normal'], fontName='Helvetica-Bold')))
    elements.append(Paragraph("Email: manuel.lazzaro@italfrigo.com | Tel: +39 3461587689", styles['Normal']))
    
    # Costruisci il documento
    doc.build(
        elements,
        onFirstPage=add_page_number,
        onLaterPages=add_page_number
    )
    
    print(f"Offerta PDF creata con successo: {output_pdf}")

def add_page_number(canvas, doc):
    """Aggiunge il numero di pagina e un'intestazione a ogni pagina"""
    width, height = landscape(A4)
    
    # Aggiungi un rettangolo bianco in alto per pulire l'area
    canvas.setFillColor(colors.white)
    canvas.rect(0, height-8*cm, width, 8*cm, fill=True, stroke=False)
    
    # Aggiungi titolo centrato
    canvas.setFont('Helvetica-Bold', 20)
    canvas.setFillColor(COLORE_PRIMARIO)
    canvas.drawCentredString(width/2, height - 3*cm, "OFFERTA COMMERCIALE")
    
    # Aggiungi logo
    if os.path.exists(logo_path):
        canvas.drawImage(logo_path, 1*cm, height - 2*cm, width=4*cm, height=1.5*cm, preserveAspectRatio=True)
    
    # Aggiungi la data di generazione
    canvas.setFont('Helvetica', 8)
    canvas.setFillColor(COLORE_TESTO)
    data_generazione = datetime.datetime.now().strftime("%d/%m/%Y")
    canvas.drawString(width - 5*cm, height - 2*cm, f"Generato il: {data_generazione}")
    
    # Nessun referente nell'intestazione
    
    # Aggiungi nota di riservatezza (più in basso)
    canvas.setFont('Helvetica-Bold', 8)
    canvas.drawString(1*cm, height - 6.5*cm, "DOCUMENTO RISERVATO E CONFIDENZIALE")
    canvas.setFont('Helvetica', 7)
    canvas.drawString(1*cm, height - 7*cm, "Questo documento contiene informazioni riservate. È vietata la divulgazione o riproduzione senza autorizzazione.")
    
    # Aggiungi il numero di pagina
    canvas.setFont('Helvetica', 8)
    canvas.drawString(width/2, 1*cm, f"Pagina {doc.page}")

if __name__ == "__main__":
    crea_offerta_pdf()
