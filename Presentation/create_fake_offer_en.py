#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.units import cm, mm
import os

# Percorso del file di output
output_file = os.path.join(os.path.dirname(__file__), "RH_Corso_Venezia_Offer_EN.pdf")

# Creazione del documento
doc = SimpleDocTemplate(output_file, pagesize=A4)

# Stili
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name='MyTitle', 
                         fontName='Helvetica-Bold', 
                         fontSize=18, 
                         alignment=1, 
                         spaceAfter=12))
styles.add(ParagraphStyle(name='MySubtitle', 
                         fontName='Helvetica-Bold', 
                         fontSize=14, 
                         alignment=1, 
                         spaceAfter=10))
styles.add(ParagraphStyle(name='MyNormal', 
                         fontName='Helvetica', 
                         fontSize=11, 
                         spaceAfter=6))
styles.add(ParagraphStyle(name='MyHeading', 
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
        print(f"Logo not found: {logo_path}")
except Exception as e:
    print(f"Error loading logo: {e}")

elements.append(Spacer(1, 20))
elements.append(Paragraph("TECHNICAL-ECONOMIC OFFER", styles['MyTitle']))
elements.append(Paragraph("Partnership Marrone-Iglu-Italfrigo", styles['MySubtitle']))
elements.append(Spacer(1, 20))

# Informazioni cliente
elements.append(Paragraph("Customer Information:", styles['MyHeading']))
data = [
    ["Customer:", "Partnership Marrone-Iglu-Italfrigo"],
    ["Address:", "Corso Venezia, Milan"],
    ["Contact:", "info@rhcorsovenezia.it"],
    ["Offer Date:", "May 1, 2025"]
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
elements.append(Paragraph("Project Description:", styles['MyHeading']))
elements.append(Paragraph("""
The project involves the installation of a complete refrigeration system for professional kitchens.
The system will be designed according to the highest quality standards, using premium components and cutting-edge technologies.
""", styles['MyNormal']))
elements.append(Spacer(1, 10))

# Dettaglio dell'offerta
elements.append(Paragraph("Offer Details:", styles['MyHeading']))

data = [
    ["Description", "Quantity", "Unit Price", "Total"],
    ["Cold rooms", "2", "€ 12,000.00", "€ 24,000.00"],
    ["Blast chiller", "1", "€ 8,500.00", "€ 8,500.00"],
    ["Refrigerated counters", "3", "€ 3,200.00", "€ 9,600.00"],
    ["Refrigeration system", "1", "€ 15,000.00", "€ 15,000.00"],
    ["Installation and testing", "1", "€ 5,000.00", "€ 5,000.00"],
    ["", "", "Total (VAT excluded)", "€ 62,100.00"],
    ["", "", "VAT 22%", "€ 13,662.00"],
    ["", "", "Total (VAT included)", "€ 75,762.00"]
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
elements.append(Paragraph("Payment Terms:", styles['MyHeading']))
elements.append(Paragraph("""
- 30% upon acceptance of the offer
- 40% at the beginning of the works
- 30% upon final testing
""", styles['MyNormal']))
elements.append(Spacer(1, 10))

# Tempistiche
elements.append(Paragraph("Implementation Timeline:", styles['MyHeading']))
elements.append(Paragraph("""
- Design: 2 weeks
- Materials supply: 4 weeks
- Installation: 2 weeks
- Testing: 1 week
""", styles['MyNormal']))
elements.append(Spacer(1, 10))

# Validità dell'offerta
elements.append(Paragraph("Offer Validity:", styles['MyHeading']))
elements.append(Paragraph("This offer is valid for 60 days from the date of issue.", styles['MyNormal']))
elements.append(Spacer(1, 20))

# Note finali
elements.append(Paragraph("Notes:", styles['MyHeading']))
elements.append(Paragraph("""
This is a temporary offer created for demonstration purposes. Prices and technical specifications are indicative and may vary in the final proposal.
""", styles['MyNormal']))
elements.append(Spacer(1, 30))

# Firma
elements.append(Paragraph("For acceptance:", styles['MyNormal']))
elements.append(Spacer(1, 20))
elements.append(Paragraph("_________________________", styles['MyNormal']))
elements.append(Paragraph("Partnership Marrone-Iglu-Italfrigo", styles['MyNormal']))

# Generazione del PDF
doc.build(elements)

print(f"PDF created successfully: {output_file}")
