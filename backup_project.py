#!/usr/bin/env python3
"""
Script per creare automaticamente un backup dell'intero progetto RH Corso Venezia.
Crea una cartella di backup con data e ora nel nome per mantenere una cronologia.
"""

import os
import sys
import shutil
import datetime
import subprocess
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
    if colore in COLORI:
        print(f"{COLORI[colore]}{testo}{COLORI['reset']}")
    else:
        print(testo)

def crea_backup():
    """Crea un backup dell'intero progetto."""
    # Ottieni il percorso della directory del progetto
    dir_progetto = os.path.dirname(os.path.abspath(__file__))
    dir_principale = os.path.dirname(dir_progetto)
    
    # Crea il nome della directory di backup con data e ora
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dir_backup = os.path.join(dir_principale, f"RH_Backup_{timestamp}")
    
    stampa_colorato(f"=== Creazione Backup Progetto RH Corso Venezia ===", "ciano")
    stampa_colorato(f"Directory progetto: {dir_progetto}", "blu")
    stampa_colorato(f"Directory backup: {dir_backup}", "blu")
    
    try:
        # Verifica se esiste già una directory di backup con lo stesso nome
        if os.path.exists(dir_backup):
            stampa_colorato(f"La directory di backup {dir_backup} esiste già.", "giallo")
            risposta = input("Vuoi sovrascriverla? (s/n): ").lower()
            if risposta != 's':
                stampa_colorato("Backup annullato dall'utente.", "rosso")
                return False
            shutil.rmtree(dir_backup)
            stampa_colorato(f"Directory di backup precedente rimossa.", "giallo")
        
        # Crea la directory di backup
        os.makedirs(dir_backup, exist_ok=True)
        stampa_colorato(f"Directory di backup creata: {dir_backup}", "verde")
        
        # Copia tutti i file e le directory del progetto
        stampa_colorato("Copia dei file in corso...", "giallo")
        
        # Escludi la directory dei backup e altri file temporanei
        escludi = [
            "RH_Backup_*",
            "__pycache__",
            "*.pyc",
            ".git",
            ".DS_Store",
            "Thumbs.db"
        ]
        
        # Usa rsync se disponibile (più veloce e con più opzioni)
        if shutil.which("rsync"):
            comando_rsync = ["rsync", "-av", "--progress"]
            for pattern in escludi:
                comando_rsync.extend(["--exclude", pattern])
            comando_rsync.extend([dir_progetto + "/", dir_backup])
            
            stampa_colorato("Utilizzo di rsync per la copia...", "blu")
            processo = subprocess.Popen(
                comando_rsync,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                universal_newlines=True
            )
            
            # Mostra l'avanzamento
            for linea in processo.stdout:
                print(linea, end='')
            
            processo.wait()
            if processo.returncode != 0:
                stampa_colorato("Errore durante la copia con rsync.", "rosso")
                return False
        else:
            # Usa shutil.copytree se rsync non è disponibile
            stampa_colorato("Utilizzo di shutil.copytree per la copia...", "blu")
            
            def ignora_file(dir, files):
                return [f for f in files for pattern in escludi if os.path.fnmatch.fnmatch(f, pattern)]
            
            shutil.copytree(
                dir_progetto, 
                dir_backup, 
                ignore=ignora_file,
                dirs_exist_ok=True
            )
        
        # Verifica che il backup sia stato creato correttamente
        if os.path.exists(dir_backup) and os.listdir(dir_backup):
            dimensione = get_dimensione_directory(dir_backup)
            stampa_colorato(f"Backup completato con successo!", "verde")
            stampa_colorato(f"Dimensione del backup: {dimensione}", "verde")
            stampa_colorato(f"Percorso del backup: {dir_backup}", "verde")
            return True
        else:
            stampa_colorato("Errore: La directory di backup è vuota o non esiste.", "rosso")
            return False
            
    except Exception as e:
        stampa_colorato(f"Errore durante la creazione del backup: {e}", "rosso")
        return False

def get_dimensione_directory(path):
    """Calcola la dimensione totale di una directory in formato leggibile."""
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(path):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            if not os.path.islink(fp):
                total_size += os.path.getsize(fp)
    
    # Converti in formato leggibile
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if total_size < 1024.0:
            return f"{total_size:.2f} {unit}"
        total_size /= 1024.0

def main():
    """Funzione principale."""
    stampa_colorato("Avvio del processo di backup...", "ciano")
    
    # Crea il backup
    inizio = time.time()
    successo = crea_backup()
    fine = time.time()
    
    if successo:
        tempo_impiegato = fine - inizio
        stampa_colorato(f"Tempo impiegato: {tempo_impiegato:.2f} secondi", "verde")
        stampa_colorato("Backup completato con successo!", "verde")
    else:
        stampa_colorato("Il backup non è stato completato correttamente.", "rosso")

if __name__ == "__main__":
    main()
