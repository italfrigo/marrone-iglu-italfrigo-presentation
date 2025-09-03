#!/usr/bin/env python3
"""
Script per avviare automaticamente la presentazione RH Corso Venezia.
Avvia un server web locale e apre il browser predefinito.
Chiude automaticamente eventuali processi che utilizzano già la porta.
"""

import http.server
import socketserver
import webbrowser
import os
import threading
import time
import sys
import subprocess
import signal
import platform
import socket
import random

# Configurazione
PORT = 8080  # Utilizziamo una porta diversa che probabilmente non è in uso
DIRECTORY = os.path.dirname(os.path.abspath(__file__))
URL = f"http://localhost:{PORT}"

def print_colored(text, color="green"):
    """Stampa testo colorato nella console."""
    colors = {
        "red": "\033[91m",
        "green": "\033[92m",
        "yellow": "\033[93m",
        "blue": "\033[94m",
        "reset": "\033[0m"
    }
    print(f"{colors.get(color, colors['green'])}{text}{colors['reset']}")

def open_browser():
    """Apre il browser dopo un breve ritardo per permettere al server di avviarsi."""
    time.sleep(1)
    print_colored(f"Apertura del browser a {URL}", "blue")
    webbrowser.open(URL)

def kill_all_python_processes():
    """Termina tutti i processi Python in esecuzione tranne quello corrente."""
    sistema = platform.system()
    current_pid = os.getpid()
    
    try:
        if sistema == "Linux" or sistema == "Darwin":  # Linux o macOS
            # Ottieni tutti i PID dei processi Python
            try:
                cmd = "pgrep -f python"
                output = subprocess.check_output(cmd, shell=True, stderr=subprocess.STDOUT, timeout=3).decode().strip()
                
                if output:
                    pids = output.split('\n')
                    for pid in pids:
                        pid = pid.strip()
                        if pid and int(pid) != current_pid:  # Non terminare il processo corrente
                            print_colored(f"Tentativo di chiusura del processo Python con PID {pid}", "yellow")
                            try:
                                # Usa SIGKILL (kill -9) per forzare la chiusura
                                subprocess.run(["kill", "-9", pid], check=True)
                                print_colored(f"Processo Python con PID {pid} terminato con successo", "green")
                            except subprocess.CalledProcessError:
                                print_colored(f"Errore nella chiusura del processo con PID {pid}", "red")
                    
                    # Attendi che tutti i processi terminino
                    time.sleep(1)
                    return True
            except subprocess.CalledProcessError:
                pass
            
        elif sistema == "Windows":
            # Per Windows, usa tasklist e taskkill
            try:
                cmd = "tasklist /FI \"IMAGENAME eq python.exe\" /FO CSV"
                output = subprocess.check_output(cmd, shell=True, stderr=subprocess.STDOUT, timeout=3).decode()
                
                if output:
                    lines = output.split('\n')[1:]  # Salta l'intestazione
                    for line in lines:
                        if line.strip():
                            parts = line.strip().replace('"', '').split(',')
                            if len(parts) > 1:
                                pid = parts[1]
                                if int(pid) != current_pid:  # Non terminare il processo corrente
                                    print_colored(f"Chiusura del processo Python con PID {pid}", "yellow")
                                    subprocess.call(f"taskkill /F /PID {pid}", shell=True)
                    
                    # Attendi che tutti i processi terminino
                    time.sleep(1)
                    return True
            except subprocess.CalledProcessError:
                pass
    except Exception as e:
        print_colored(f"Errore durante la chiusura dei processi Python: {e}", "red")
    
    return False

def find_and_kill_process_on_port(port):
    """Trova e termina qualsiasi processo che utilizza la porta specificata."""
    sistema = platform.system()
    success = False
    
    try:
        if sistema == "Linux" or sistema == "Darwin":  # Linux o macOS
            # Metodo 1: Prova con lsof
            try:
                print_colored(f"Ricerca processi sulla porta {port} con lsof...", "yellow")
                cmd = f"lsof -i :{port} -t"
                output = subprocess.check_output(cmd, shell=True, stderr=subprocess.STDOUT, timeout=3).decode().strip()
                
                if output:
                    pids = output.split('\n')
                    for pid in pids:
                        pid = pid.strip()
                        if pid:
                            print_colored(f"Tentativo di chiusura del processo con PID {pid} sulla porta {port}", "yellow")
                            try:
                                # Prima prova con SIGTERM (più gentile)
                                subprocess.run(["kill", pid], check=False)
                                time.sleep(0.5)
                                # Se ancora in esecuzione, usa SIGKILL
                                if subprocess.call(f"ps -p {pid} > /dev/null", shell=True) == 0:
                                    subprocess.run(["kill", "-9", pid], check=False)
                                    print_colored(f"Processo con PID {pid} terminato con SIGKILL", "green")
                                else:
                                    print_colored(f"Processo con PID {pid} terminato con SIGTERM", "green")
                                success = True
                            except Exception as e:
                                print_colored(f"Errore nella chiusura del processo con PID {pid}: {e}", "red")
            except Exception as e:
                print_colored(f"Errore con lsof: {e}", "red")
            
            # Metodo 2: Prova con fuser se lsof non ha avuto successo
            if not success:
                try:
                    print_colored(f"Tentativo con fuser sulla porta {port}...", "yellow")
                    cmd = f"fuser -k {port}/tcp 2>/dev/null"
                    subprocess.call(cmd, shell=True, timeout=3)
                    time.sleep(1)
                    # Verifica se la porta è stata liberata
                    if not is_port_in_use(port):
                        print_colored(f"Porta {port} liberata con fuser", "green")
                        success = True
                except Exception as e:
                    print_colored(f"Errore con fuser: {e}", "red")
            
            # Metodo 3: Prova con netstat e grep se gli altri metodi non hanno funzionato
            if not success:
                try:
                    print_colored(f"Tentativo con netstat sulla porta {port}...", "yellow")
                    cmd = f"netstat -tulpn 2>/dev/null | grep :{port}"
                    output = subprocess.check_output(cmd, shell=True, timeout=3).decode().strip()
                    
                    if output:
                        # Estrai il PID usando una regex
                        import re
                        pid_match = re.search(r'LISTEN\s+(\d+)/', output)
                        if pid_match:
                            pid = pid_match.group(1)
                            print_colored(f"Tentativo di chiusura del processo con PID {pid}", "yellow")
                            subprocess.call(f"kill -9 {pid}", shell=True)
                            time.sleep(1)
                            if not is_port_in_use(port):
                                print_colored(f"Porta {port} liberata con netstat+kill", "green")
                                success = True
                except Exception as e:
                    print_colored(f"Errore con netstat: {e}", "red")
        
        elif sistema == "Windows":
            # Per Windows, usa netstat e taskkill
            try:
                cmd = f"netstat -ano | findstr :{port}"
                output = subprocess.check_output(cmd, shell=True, stderr=subprocess.STDOUT, timeout=3).decode()
                
                if output:
                    lines = output.split('\n')
                    for line in lines:
                        if f":{port}" in line and "LISTENING" in line:
                            parts = line.strip().split()
                            if len(parts) > 4:
                                pid = parts[-1]
                                print_colored(f"Chiusura del processo con PID {pid} sulla porta {port}", "yellow")
                                subprocess.call(f"taskkill /F /PID {pid}", shell=True)
                                time.sleep(1)
                                if not is_port_in_use(port):
                                    print_colored(f"Porta {port} liberata", "green")
                                    success = True
            except Exception as e:
                print_colored(f"Errore con netstat/taskkill: {e}", "red")
    except Exception as e:
        print_colored(f"Errore durante la chiusura del processo: {e}", "red")
    
    return success

def is_port_in_use(port):
    """Verifica se la porta è già in uso."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('localhost', port)) == 0

def try_alternative_port():
    """Cerca una porta alternativa se quella principale è occupata."""
    global PORT, URL
    for alt_port in range(8001, 8020):  # Prova le porte da 8001 a 8019
        if not is_port_in_use(alt_port):
            PORT = alt_port
            URL = f"http://localhost:{PORT}"
            return True
    return False

def start_server():
    """Avvia il server HTTP."""
    global PORT, URL
    os.chdir(DIRECTORY)
    
    # Verifica se la porta è già in uso
    max_attempts = 3
    attempt = 0
    
    while is_port_in_use(PORT) and attempt < max_attempts:
        attempt += 1
        print_colored(f"Tentativo {attempt}/{max_attempts}: La porta {PORT} è già in uso.", "yellow")
        
        # Prova a chiudere il processo che occupa la porta
        if find_and_kill_process_on_port(PORT):
            print_colored(f"Porta {PORT} liberata con successo.", "green")
            # Attendi un momento per assicurarsi che la porta sia completamente liberata
            time.sleep(1)
            if not is_port_in_use(PORT):
                break
        
        # Se non siamo riusciti a liberare la porta, prova una porta alternativa
        original_port = PORT
        if try_alternative_port():
            print_colored(f"Utilizzo della porta alternativa {PORT} invece di {original_port}", "green")
            break
    
    # Se dopo tutti i tentativi la porta è ancora occupata
    if is_port_in_use(PORT):
        print_colored(f"ERRORE: Impossibile avviare il server sulla porta {PORT} dopo {max_attempts} tentativi.", "red")
        print_colored("Suggerimento: Riavvia il computer o chiudi manualmente i processi che occupano la porta.", "yellow")
        sys.exit(1)
    
    # Configura e avvia il server
    handler = http.server.SimpleHTTPRequestHandler
    
    try:
        # Prova ad avviare il server con binding solo su localhost per maggiore sicurezza
        with socketserver.TCPServer(("localhost", PORT), handler) as httpd:
            print_colored(f"Server avviato con successo alla porta {PORT}", "green")
            print_colored(f"Presentazione disponibile all'indirizzo: {URL}", "green")
            print_colored("Premi Ctrl+C per terminare il server", "yellow")
            
            # Avvia il browser in un thread separato
            threading.Thread(target=open_browser, daemon=True).start()
            
            # Mantieni il server in esecuzione
            httpd.serve_forever()
    except OSError as e:
        print_colored(f"ERRORE nell'avvio del server: {e}", "red")
        # Prova un'ultima volta con una porta completamente diversa
        PORT = 9000 + random.randint(0, 999)  # Porta casuale tra 9000 e 9999
        URL = f"http://localhost:{PORT}"
        print_colored(f"Tentativo finale con porta casuale {PORT}...", "yellow")
        try:
            with socketserver.TCPServer(("localhost", PORT), handler) as httpd:
                print_colored(f"Server avviato con successo alla porta {PORT}", "green")
                print_colored(f"Presentazione disponibile all'indirizzo: {URL}", "green")
                print_colored("Premi Ctrl+C per terminare il server", "yellow")
                
                # Avvia il browser in un thread separato
                threading.Thread(target=open_browser, daemon=True).start()
                
                # Mantieni il server in esecuzione
                httpd.serve_forever()
        except Exception as e2:
            print_colored(f"ERRORE FATALE: Impossibile avviare il server: {e2}", "red")
            sys.exit(1)
    except KeyboardInterrupt:
        print_colored("\nServer terminato dall'utente", "yellow")

if __name__ == "__main__":
    print_colored("=== Avvio Presentazione RH Corso Venezia ===", "green")
    
    # Verifica se la porta è già in uso
    if is_port_in_use(PORT):
        print_colored(f"La porta {PORT} è già in uso. Tentativo di liberarla...", "yellow")
        # Prova a liberare la porta direttamente
        find_and_kill_process_on_port(PORT)
    
    # Avvia il server con gestione degli errori
    try:
        start_server()
    except KeyboardInterrupt:
        print_colored("\nPresentazione terminata dall'utente", "yellow")
    except Exception as e:
        print_colored(f"Errore imprevisto: {e}", "red")
        print_colored("Suggerimento: Prova a riavviare il computer se il problema persiste", "yellow")
        sys.exit(1)
