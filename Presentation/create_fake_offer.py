#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.units import cm, mm
import os

# Percorso del file di output
output_file = os.path.join(os.path.dirname(__file__), "RH_Corso_Venezia_Offerta.pdf")

# Creazione del documento
doc = SimpleDocTemplate(output_file, pagesize=A4)

# Stili
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name='Titolo', 
                         fontName='Helvetica-Bold', 
                         fontSize=18, 
                         alignment=1, 
                         spaceAfter=12))
styles.add(ParagraphStyle(name='Sottotitolo', 
                         fontName='Helvetica-Bold', 
                         fontSize=14, 
                         alignment=1, 
                         spaceAfter=10))
styles.add(ParagraphStyle(name='Normale', 
                         fontName='Helvetica', 
                         fontSize=11, 
                         spaceAfter=6))
styles.add(ParagraphStyle(name='Intestazione', 
                         fontName='Helvetica-Bold', 
                         fontSize=12, 
                         spaceAfter=8))

# Contenuto del documento
elements = []

# Intestazione con logo
try:
    logo_path = os.path.join(os.path.dirname(__file__), "assets/images/logo.png")
    if os.path.exists(logo_path):
        logo = Image(logo_path, width=5*cm, height=2*cm)
        elements.append(logo)
    else:
        print(f"Logo non trovato: {logo_path}")
except Exception as e:
    print(f"Errore nel caricamento del logo: {e}")

elements.append(Spacer(1, 20))
elements.append(Paragraph("OFFERTA TECNICO-ECONOMICA", styles['Titolo']))
elements.append(Paragraph("Partnership Marrone-Iglu-Italfrigo", styles['Sottotitolo']))
elements.append(Spacer(1, 20))

# Informazioni cliente
elements.append(Paragraph("Informazioni Cliente:", styles['Intestazione']))
data = [
    ["Cliente:", "Partnership Marrone-Iglu-Italfrigo"],
    ["Indirizzo:", "Corso Venezia, Milano"],
    ["Contatto:", "info@rhcorsovenezia.it"],
    ["Data Offerta:", "1 Maggio 2025"]
]

table = Table(data, colWidths=[100, 300])
table.setStyle(TableStyle([
    ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
    ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
    ('FONTSIZE', (0, 0), (-1, -1), 11),
    ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
]))
elements.append(table)
elements.append(Spacer(1, 20))

# Descrizione del progetto
elements.append(Paragraph("Descrizione del Progetto:", styles['Intestazione']))
elements.append(Paragraph("""
Il progetto prevede la realizzazione di un impianto di refrigerazione completo per cucine professionali.
L'impianto sarà progettato secondo i più alti standard qualitativi, utilizzando componenti di prima scelta e tecnologie all'avanguardia.
""", styles['Normale']))
elements.append(Spacer(1, 10))

# Dettaglio dell'offerta
elements.append(Paragraph("Dettaglio dell'Offerta:", styles['Intestazione']))

data = [
    ["Descrizione", "Quantità", "Prezzo Unitario", "Totale"],
    ["Celle frigorifere", "2", "€ 12.000,00", "€ 24.000,00"],
    ["Abbattitore di temperatura", "1", "€ 8.500,00", "€ 8.500,00"],
    ["Banchi refrigerati", "3", "€ 3.200,00", "€ 9.600,00"],
    ["Impianto di refrigerazione", "1", "€ 15.000,00", "€ 15.000,00"],
    ["Installazione e collaudo", "1", "€ 5.000,00", "€ 5.000,00"],
    ["", "", "Totale (IVA esclusa)", "€ 62.100,00"],
    ["", "", "IVA 22%", "€ 13.662,00"],
    ["", "", "Totale (IVA inclusa)", "€ 75.762,00"]
]

table = Table(data, colWidths=[200, 70, 120, 100])
table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
    ('ALIGN', (0, 0), (0, -1), 'LEFT'),
    ('FONTNAME', (0, -3), (-1, -1), 'Helvetica-Bold'),
    ('LINEABOVE', (0, -3), (-1, -3), 1, colors.black),
    ('GRID', (0, 0), (-1, -4), 0.5, colors.grey),
]))
elements.append(table)
elements.append(Spacer(1, 20))

# Condizioni di pagamento
elements.append(Paragraph("Condizioni di Pagamento:", styles['Intestazione']))
elements.append(Paragraph("""
- 30% all'accettazione dell'offerta
- 40% all'inizio dei lavori
- 30% al collaudo finale
""", styles['Normale']))
elements.append(Spacer(1, 10))

# Tempistiche
elements.append(Paragraph("Tempistiche di Realizzazione:", styles['Intestazione']))
elements.append(Paragraph("""
- Progettazione: 2 settimane
- Fornitura materiali: 4 settimane
- Installazione: 2 settimane
- Collaudo: 1 settimana
""", styles['Normale']))
elements.append(Spacer(1, 10))

# Validità dell'offerta
elements.append(Paragraph("Validità dell'Offerta:", styles['Intestazione']))
elements.append(Paragraph("La presente offerta ha validità di 60 giorni dalla data di emissione.", styles['Normale']))
elements.append(Spacer(1, 20))

# Note finali
elements.append(Paragraph("Note:", styles['Intestazione']))
elements.append(Paragraph("""
Questa è un'offerta temporanea creata a scopo dimostrativo. I prezzi e le specifiche tecniche sono indicativi e potrebbero variare nella proposta definitiva.
""", styles['Normale']))
elements.append(Spacer(1, 30))

# Firma
elements.append(Paragraph("Per accettazione:", styles['Normale']))
elements.append(Spacer(1, 20))
elements.append(Paragraph("_________________________", styles['Normale']))
elements.append(Paragraph("Partnership Marrone-Iglu-Italfrigo", styles['Normale']))

# Generazione del PDF
doc.build(elements)

print(f"PDF creato con successo: {output_file}")
