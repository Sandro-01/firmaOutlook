"""
Gestore Firme Email Active Directory -> Outlook
VERSIONE 4.3 - CARTON GROUP ITALIA
- Logo personalizzabile da file locale
- Solo sedi Italia (Treviso, Perugia, Verona)
- Indirizzo differenziato per sede (Perugia ha indirizzo proprio)
- Firma NON impostata come predefinita (utente libero di modificare)
- UI moderna con progress bar
- Salvataggio locale + sync su server
- Telefono/Mobile: mostra uno solo (priorità mobile)
"""

import os
import sys
import base64
import shutil
from pathlib import Path
from datetime import datetime

# Colori console Windows
os.system('color')

class Colors:
    HEADER = '\033[95m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    WHITE = '\033[97m'
    BOLD = '\033[1m'
    END = '\033[0m'

def print_header(text):
    print(f"\n{Colors.CYAN}{'═'*70}")
    print(f"  {Colors.BOLD}{text}{Colors.END}")
    print(f"{Colors.CYAN}{'═'*70}{Colors.END}")

def print_success(text):
    print(f"{Colors.GREEN}✓ {text}{Colors.END}")

def print_error(text):
    print(f"{Colors.RED}✗ {text}{Colors.END}")

def print_warning(text):
    print(f"{Colors.YELLOW}⚠ {text}{Colors.END}")

def print_info(text):
    print(f"{Colors.BLUE}→ {text}{Colors.END}")

def print_menu_item(num, icon, text, desc=""):
    if desc:
        print(f"  {Colors.WHITE}{num}.{Colors.END} {icon} {Colors.BOLD}{text}{Colors.END} {Colors.CYAN}- {desc}{Colors.END}")
    else:
        print(f"  {Colors.WHITE}{num}.{Colors.END} {icon} {Colors.BOLD}{text}{Colors.END}")

def progress_bar(current, total, width=40):
    percent = current / total
    filled = int(width * percent)
    bar = '█' * filled + '░' * (width - filled)
    print(f"\r  {Colors.CYAN}[{bar}] {current}/{total} ({percent*100:.0f}%){Colors.END}", end='', flush=True)

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

try:
    from ldap3 import Server, Connection, ALL, SUBTREE
except ImportError:
    print_error("Modulo ldap3 non trovato!")
    print_info("Installa con: pip install ldap3")
    sys.exit(1)


class ADSignatureManager:
    def __init__(self):
        # Configurazione predefinita Carton Group
        self.ad_server = "adds.cartongrp.com"
        self.domain = "CARTONGRP"
        self.connection = None
        self.users = []

        # Paths
        self.local_path = Path(os.environ.get('TEMP', 'C:\\Temp')) / "Firme_Generate"
        self.server_path = Path(r"\\DEAZRADS101\Firme")

        # Sedi SOLO ITALIA
        self.sedi = {
            '1': ('Treviso', 'OU=Treviso,OU=rIT,OU=Client,DC=adds,DC=cartongrp,DC=com'),
            '2': ('Perugia', 'OU=Perugia,OU=rIT,OU=Client,DC=adds,DC=cartongrp,DC=com'),
            '3': ('Verona', 'OU=Verona,OU=rIT,OU=Client,DC=adds,DC=cartongrp,DC=com'),
            '4': ('TUTTE LE SEDI ITALIA', 'OU=rIT,OU=Client,DC=adds,DC=cartongrp,DC=com'),
        }

        # Indirizzi per sede
        self.indirizzi = {
            'default': 'Via Turati, 49 | Paese (TV) - Italy',
            'perugia': 'Via Guido Rossa, 5 | S. Sabina (PG) - Italy'
        }

        # Logo - default online, può essere sovrascritto con file locale
        self.logo_src = 'https://tse2.mm.bing.net/th/id/OIP.wq2AFoiw0fNYotAoONJAlwAAAA'
        self.logo_base64 = None  # Se impostato, usa questo invece dell'URL

        # Log
        self.log_file = self.local_path / f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

    def log(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_line = f"[{timestamp}] [{level}] {message}"

        try:
            self.local_path.mkdir(parents=True, exist_ok=True)
            with open(self.log_file, 'a', encoding='utf-8') as f:
                f.write(log_line + "\n")
        except:
            pass

    def load_logo_from_file(self, file_path):
        """Carica logo da file locale e converte in Base64"""
        try:
            file_path = Path(file_path.strip('"').strip("'"))

            if not file_path.exists():
                print_error(f"File non trovato: {file_path}")
                return False

            # Leggi e converti in base64
            with open(file_path, 'rb') as f:
                image_data = f.read()

            encoded = base64.b64encode(image_data).decode('utf-8')

            # Determina MIME type
            ext = file_path.suffix.lower()
            mime_types = {
                '.png': 'image/png',
                '.jpg': 'image/jpeg',
                '.jpeg': 'image/jpeg',
                '.gif': 'image/gif',
                '.bmp': 'image/bmp',
                '.svg': 'image/svg+xml'
            }
            mime_type = mime_types.get(ext, 'image/png')

            # Crea data URL
            self.logo_base64 = f"data:{mime_type};base64,{encoded}"
            self.logo_src = self.logo_base64

            size_kb = len(image_data) / 1024
            print_success(f"Logo caricato: {file_path.name} ({size_kb:.1f} KB)")
            self.log(f"Logo caricato: {file_path}")
            return True

        except Exception as e:
            print_error(f"Errore caricamento logo: {e}")
            return False

    def show_banner(self):
        clear_screen()
        print(f"""
{Colors.CYAN}╔══════════════════════════════════════════════════════════════════════╗
║                                                                      ║
║  {Colors.WHITE}{Colors.BOLD}   ██████╗ █████╗ ██████╗ ████████╗ ██████╗ ███╗   ██╗{Colors.CYAN}               ║
║  {Colors.WHITE}{Colors.BOLD}  ██╔════╝██╔══██╗██╔══██╗╚══██╔══╝██╔═══██╗████╗  ██║{Colors.CYAN}               ║
║  {Colors.WHITE}{Colors.BOLD}  ██║     ███████║██████╔╝   ██║   ██║   ██║██╔██╗ ██║{Colors.CYAN}               ║
║  {Colors.WHITE}{Colors.BOLD}  ██║     ██╔══██║██╔══██╗   ██║   ██║   ██║██║╚██╗██║{Colors.CYAN}               ║
║  {Colors.WHITE}{Colors.BOLD}  ╚██████╗██║  ██║██║  ██║   ██║   ╚██████╔╝██║ ╚████║{Colors.CYAN}               ║
║  {Colors.WHITE}{Colors.BOLD}   ╚═════╝╚═╝  ╚═╝╚═╝  ╚═╝   ╚═╝    ╚═════╝ ╚═╝  ╚═══╝{Colors.CYAN}               ║
║                                                                      ║
║  {Colors.GREEN}Gestore Firme Email v4.3 - ITALIA{Colors.CYAN}                               ║
║  {Colors.WHITE}Active Directory → Outlook Signatures{Colors.CYAN}                          ║
║  {Colors.YELLOW}Firma libera: l'utente può modificarla in Outlook{Colors.CYAN}               ║
║  {Colors.YELLOW}Indirizzo differenziato per sede (TV/PG){Colors.CYAN}                        ║
║                                                                      ║
╚══════════════════════════════════════════════════════════════════════╝{Colors.END}
""")

    def connect(self, username, password):
        """Connette ad Active Directory"""
        print_info(f"Connessione a {self.ad_server}...")

        try:
            server = Server(self.ad_server, port=389, get_info=ALL)

            user_formats = [
                f"{username}@{self.ad_server}",
                f"{self.domain}\\{username}",
                username
            ]

            for user_format in user_formats:
                try:
                    self.connection = Connection(server, user_format, password, auto_bind=True)
                    print_success("Connesso ad Active Directory")
                    self.log(f"Connesso come {username}")
                    return True
                except:
                    continue

            print_error("Credenziali non valide")
            return False

        except Exception as e:
            print_error(f"Errore connessione: {e}")
            self.log(f"Errore connessione: {e}", "ERROR")
            return False

    def search_users(self, base_dn, filter_text=""):
        """Cerca utenti in Active Directory"""
        if not self.connection:
            return []

        if filter_text:
            ldap_filter = f"(&(objectClass=user)(mail=*)(|(sAMAccountName=*{filter_text}*)(displayName=*{filter_text}*)(mail=*{filter_text}*)))"
        else:
            ldap_filter = "(&(objectClass=user)(mail=*))"

        # Aggiunto telephoneNumber per il telefono fisso
        attributes = ['sAMAccountName', 'displayName', 'givenName', 'sn', 'title', 'mail', 'mobile', 'telephoneNumber', 'department', 'distinguishedName']

        try:
            self.connection.search(
                search_base=base_dn,
                search_filter=ldap_filter,
                search_scope=SUBTREE,
                attributes=attributes
            )

            self.users = []
            for entry in self.connection.entries:
                def get_attr(attr_name):
                    if hasattr(entry, attr_name):
                        val = getattr(entry, attr_name)
                        if val and str(val) != '[]':
                            return str(val).strip('[]')
                    return ''

                user = {
                    'username': get_attr('sAMAccountName'),
                    'display_name': get_attr('displayName'),
                    'first_name': get_attr('givenName'),
                    'last_name': get_attr('sn'),
                    'title': get_attr('title') or 'Employee',
                    'email': get_attr('mail'),
                    'mobile': get_attr('mobile'),
                    'phone': get_attr('telephoneNumber'),  # Telefono fisso (Home)
                    'department': get_attr('department'),
                    'dn': get_attr('distinguishedName')
                }

                if user['email'] and '@' in user['email']:
                    self.users.append(user)

            self.users.sort(key=lambda x: x['email'])
            return self.users

        except Exception as e:
            print_error(f"Errore ricerca: {e}")
            self.log(f"Errore ricerca: {e}", "ERROR")
            return []

    def format_name(self, user):
        """Formatta nome come: Nome Cognome"""
        if user.get('first_name') and user.get('last_name'):
            return f"{user['first_name']} {user['last_name']}"

        display = user['display_name'].replace('(Europoligrafico)', '').replace('(Carton Group)', '').strip()

        if ',' in display:
            parts = [p.strip() for p in display.split(',')]
            if len(parts) == 2:
                return f"{parts[1]} {parts[0]}"

        return display

    def get_indirizzo(self, user):
        """Restituisce l'indirizzo corretto in base alla sede dell'utente"""
        dn = user.get('dn', '').lower()

        # Se l'utente è nella OU di Perugia, usa l'indirizzo di Perugia
        if 'ou=perugia' in dn:
            return self.indirizzi['perugia']

        # Altrimenti usa l'indirizzo default (Treviso)
        return self.indirizzi['default']

    def get_phone_display(self, user):
        """
        Restituisce il numero di telefono da mostrare.
        Priorità: Mobile > Telefono fisso
        Restituisce tupla (prefisso, numero) o (None, None) se nessuno disponibile
        """
        if user.get('mobile'):
            return ('M', user['mobile'])
        elif user.get('phone'):
            return ('T', user['phone'])
        return (None, None)

    def generate_html(self, user):
        """Genera firma HTML - SENZA impostazione come predefinita"""
        name = self.format_name(user)
        indirizzo = self.get_indirizzo(user)

        # Mostra solo mobile OPPURE telefono (priorità mobile)
        phone_prefix, phone_number = self.get_phone_display(user)
        phone_line = ""
        if phone_number:
            phone_line = f'<div style="margin-bottom: 8px; font-size: 9.5pt; color: #000000;">{phone_prefix}: {phone_number}</div>'

        return f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin: 0; padding: 0; font-family: Arial, Helvetica, sans-serif;">
    <table cellpadding="0" cellspacing="0" border="0" style="font-family: Arial, Helvetica, sans-serif;">
        <tr>
            <td style="padding-right: 12px; vertical-align: middle; width: 150px;">
                <img src="{self.logo_src}" alt="Carton Group" width="150" style="display: block; border: 0;">
            </td>
            <td style="width: 1px; background-color: #ADFF2F; padding: 0;"></td>
            <td style="padding-left: 12px; vertical-align: top;">
                <div style="margin-bottom: 2px;">
                    <strong style="font-size: 12pt; color: #000000;">{name}</strong>
                </div>
                <div style="margin-bottom: 10px;">
                    <span style="font-size: 9.5pt; color: #000000; font-style: italic;">{user['title']}</span>
                </div>
                {phone_line}
                <div style="margin-bottom: 8px; font-size: 9.5pt; color: #000000;">
                    {indirizzo}
                </div>
                <div style="margin-bottom: 8px; font-size: 9.5pt;">
                    <a href="https://www.carton-group.com/" style="color: #0563C1;">www.carton-group.com</a>
                </div>
                <div style="font-size: 9.5pt;">
                    <strong>Follow us:</strong>
                    <a href="https://www.linkedin.com/company/carton-group-gmbh" style="color: #0563C1;">Carton Group</a>
                </div>
            </td>
        </tr>
    </table>
</body>
</html>"""

    def generate_txt(self, user):
        """Genera firma testo"""
        name = self.format_name(user)
        indirizzo = self.get_indirizzo(user)

        # Mostra solo mobile OPPURE telefono (priorità mobile)
        phone_prefix, phone_number = self.get_phone_display(user)
        phone_line = ""
        if phone_number:
            phone_line = f"\n{phone_prefix}: {phone_number}"

        return f"""{name}
{user['title']}{phone_line}

{indirizzo}
www.carton-group.com
Follow us: Carton Group"""

    def generate_signatures(self, users=None):
        """Genera firme per gli utenti selezionati"""
        if users is None:
            users = self.users

        if not users:
            print_warning("Nessun utente da processare")
            return 0, 0

        self.local_path.mkdir(parents=True, exist_ok=True)

        print_info(f"Generazione {len(users)} firme in {self.local_path}")
        print()

        success = 0
        errors = 0

        for i, user in enumerate(users, 1):
            progress_bar(i, len(users))

            try:
                email = user['email'].lower()
                user_folder = self.local_path / email
                user_folder.mkdir(parents=True, exist_ok=True)

                (user_folder / "Firma_Aziendale.htm").write_text(self.generate_html(user), encoding='utf-8')
                (user_folder / "Firma_Aziendale.txt").write_text(self.generate_txt(user), encoding='utf-8')

                success += 1
                sede = "Perugia" if 'ou=perugia' in user.get('dn', '').lower() else "Treviso"
                phone_type = "mobile" if user.get('mobile') else ("phone" if user.get('phone') else "nessuno")
                self.log(f"Generata firma: {email} (sede: {sede}, tel: {phone_type})")

            except Exception as e:
                errors += 1
                self.log(f"Errore {user['email']}: {e}", "ERROR")

        print()
        return success, errors

    def sync_to_server(self):
        """Sincronizza firme locali sul server"""
        if not self.local_path.exists():
            print_error("Nessuna firma locale da sincronizzare")
            return 0, 0

        print_info(f"Verifica accesso a {self.server_path}...")

        if not self.server_path.exists():
            print_error(f"Server non raggiungibile: {self.server_path}")
            print_warning("Assicurati di essere connesso alla rete aziendale")
            return 0, 0

        folders = [f for f in self.local_path.iterdir() if f.is_dir() and '@' in f.name]

        if not folders:
            print_warning("Nessuna cartella firma trovata in locale")
            return 0, 0

        print_info(f"Sincronizzazione {len(folders)} firme sul server...")
        print()

        success = 0
        errors = 0

        for i, folder in enumerate(folders, 1):
            progress_bar(i, len(folders))

            try:
                dest_folder = self.server_path / folder.name

                if dest_folder.exists():
                    shutil.rmtree(dest_folder)

                shutil.copytree(folder, dest_folder)
                success += 1
                self.log(f"Sincronizzata: {folder.name}")

            except Exception as e:
                errors += 1
                self.log(f"Errore sync {folder.name}: {e}", "ERROR")

        print()
        return success, errors

    def show_users_table(self):
        """Mostra tabella utenti"""
        if not self.users:
            print_warning("Nessun utente caricato")
            return

        print(f"\n{Colors.CYAN}{'─'*110}{Colors.END}")
        print(f"{Colors.BOLD}{'#':<4} {'Email':<35} {'Nome':<22} {'Ruolo':<15} {'Sede':<8} {'Tel':<15}{Colors.END}")
        print(f"{Colors.CYAN}{'─'*110}{Colors.END}")

        for i, user in enumerate(self.users, 1):
            name = self.format_name(user)[:21]
            title = (user['title'] or 'N/A')[:14]
            email = user['email'][:34]
            sede = "Perugia" if 'ou=perugia' in user.get('dn', '').lower() else "Treviso"

            # Mostra quale telefono verrà usato
            phone_prefix, phone_number = self.get_phone_display(user)
            tel_display = f"{phone_prefix}:{phone_number[-8:]}" if phone_number else "-"

            print(f"{i:<4} {email:<35} {name:<22} {title:<15} {sede:<8} {tel_display:<15}")

        print(f"{Colors.CYAN}{'─'*110}{Colors.END}")
        print(f"{Colors.GREEN}Totale: {len(self.users)} utenti{Colors.END}")
        print(f"{Colors.CYAN}Tel: M=Mobile, T=Telefono fisso (priorità al Mobile){Colors.END}\n")

    def menu_sede(self):
        """Menu selezione sede ITALIA"""
        print_header("SELEZIONA SEDE ITALIA")

        for key, (nome, _) in self.sedi.items():
            print_menu_item(key, "🇮🇹", nome)

        print()
        choice = input(f"{Colors.WHITE}Scegli (1-4) [{Colors.CYAN}4{Colors.WHITE}]: {Colors.END}").strip() or '4'

        if choice in self.sedi:
            return self.sedi[choice]
        return self.sedi['4']

    def menu_logo(self):
        """Menu gestione logo"""
        print_header("CONFIGURAZIONE LOGO")

        print_menu_item(1, "🌐", "Usa logo online (default)", "URL attuale")
        print_menu_item(2, "📁", "Carica logo da file", "PNG, JPG, GIF, BMP")
        print_menu_item(3, "⬅️", "Torna al menu principale", "")

        print()
        if self.logo_base64:
            print_success("Logo personalizzato già caricato")
        else:
            print_info(f"Logo attuale: {self.logo_src[:50]}...")

        print()
        choice = input(f"{Colors.WHITE}Scegli (1-3): {Colors.END}").strip()

        if choice == '2':
            print()
            print_info("Inserisci il percorso completo del file logo")
            print_info("Esempio: C:\\Users\\sandro\\Desktop\\logo.png")
            print()
            file_path = input(f"{Colors.WHITE}Percorso file: {Colors.END}").strip()

            if file_path:
                self.load_logo_from_file(file_path)

            input(f"\n{Colors.CYAN}Premi Invio per continuare...{Colors.END}")

    def menu_principale(self):
        """Menu principale"""
        while True:
            print_header("MENU PRINCIPALE")

            logo_status = "✓ personalizzato" if self.logo_base64 else "default online"

            print_menu_item(1, "🔍", "Cerca utenti", "filtra per nome/email")
            print_menu_item(2, "📋", "Mostra utenti", f"{len(self.users)} caricati")
            print_menu_item(3, "🖼️", "Configura logo", logo_status)
            print_menu_item(4, "⚡", "Genera TUTTE le firme", "salva in locale")
            print_menu_item(5, "☁️", "Sincronizza su SERVER", "copia su \\\\DEAZRADS101\\Firme")
            print_menu_item(6, "🚀", "AUTOMATICO", "genera + sincronizza")
            print_menu_item(7, "👤", "Genera per utente specifico", "selezione manuale")
            print_menu_item(0, "🚪", "Esci", "")

            print()
            print_info("Indirizzi: Treviso = Via Turati | Perugia = Via Guido Rossa")
            print_info("Telefono: mostra Mobile se presente, altrimenti Tel. fisso")
            print()
            choice = input(f"{Colors.WHITE}Scegli: {Colors.END}").strip()

            if choice == '1':
                self.action_search()
            elif choice == '2':
                self.show_users_table()
                input(f"\n{Colors.CYAN}Premi Invio per continuare...{Colors.END}")
            elif choice == '3':
                self.menu_logo()
            elif choice == '4':
                self.action_generate()
            elif choice == '5':
                self.action_sync()
            elif choice == '6':
                self.action_auto()
            elif choice == '7':
                self.action_single_user()
            elif choice == '0':
                print_info("Arrivederci!")
                break

    def action_search(self):
        """Ricerca utenti con filtro"""
        print_header("RICERCA UTENTI")
        filter_text = input(f"{Colors.WHITE}Filtro (vuoto = tutti): {Colors.END}").strip()

        nome_sede, base_dn = self.current_sede
        print_info(f"Ricerca in {nome_sede}...")

        self.search_users(base_dn, filter_text)
        print_success(f"Trovati {len(self.users)} utenti")

        if self.users and len(self.users) <= 20:
            self.show_users_table()

    def action_generate(self):
        """Genera tutte le firme in locale"""
        if not self.users:
            print_warning("Prima cerca gli utenti (opzione 1)")
            return

        print_header("GENERAZIONE FIRME")
        success, errors = self.generate_signatures()

        print()
        print_success(f"Generate: {success}")
        if errors:
            print_error(f"Errori: {errors}")

        print_info(f"Salvate in: {self.local_path}")
        input(f"\n{Colors.CYAN}Premi Invio per continuare...{Colors.END}")

    def action_sync(self):
        """Sincronizza su server"""
        print_header("SINCRONIZZAZIONE SERVER")

        confirm = input(f"{Colors.YELLOW}Sincronizzare su {self.server_path}? (s/N): {Colors.END}").strip().lower()

        if confirm == 's':
            success, errors = self.sync_to_server()

            print()
            if success:
                print_success(f"Sincronizzate: {success}")
            if errors:
                print_error(f"Errori: {errors}")
        else:
            print_warning("Operazione annullata")

        input(f"\n{Colors.CYAN}Premi Invio per continuare...{Colors.END}")

    def action_auto(self):
        """Genera e sincronizza automaticamente"""
        if not self.users:
            print_warning("Prima cerca gli utenti (opzione 1)")
            return

        print_header("MODALITÀ AUTOMATICA")
        print_info(f"Verranno generate {len(self.users)} firme e sincronizzate sul server")
        print_warning("Le firme NON saranno impostate come predefinite")
        print_info("Gli utenti potranno modificarle liberamente in Outlook")
        print_info("Indirizzi: Treviso = Via Turati | Perugia = Via Guido Rossa")
        print_info("Telefono: Mobile (priorità) oppure Tel. fisso")

        confirm = input(f"\n{Colors.YELLOW}Procedere? (s/N): {Colors.END}").strip().lower()

        if confirm != 's':
            print_warning("Operazione annullata")
            return

        print()
        print_info("STEP 1/2: Generazione firme...")
        gen_ok, gen_err = self.generate_signatures()
        print()
        print_success(f"Generate: {gen_ok}")

        print()
        print_info("STEP 2/2: Sincronizzazione server...")
        sync_ok, sync_err = self.sync_to_server()
        print()

        print_header("RIEPILOGO")
        print_success(f"Firme generate: {gen_ok}")
        print_success(f"Firme sincronizzate: {sync_ok}")
        if gen_err or sync_err:
            print_error(f"Errori totali: {gen_err + sync_err}")

        print()
        print_info("Le firme sono ora disponibili per la distribuzione GPO")
        print_info("Gli utenti potranno modificarle in Outlook → File → Opzioni → Posta → Firme")
        print_info(f"Log: {self.log_file}")
        input(f"\n{Colors.CYAN}Premi Invio per continuare...{Colors.END}")

    def action_single_user(self):
        """Genera firma per utente specifico"""
        if not self.users:
            print_warning("Prima cerca gli utenti (opzione 1)")
            return

        self.show_users_table()

        try:
            num = int(input(f"{Colors.WHITE}Numero utente: {Colors.END}").strip())
            if 1 <= num <= len(self.users):
                user = self.users[num - 1]
                sede = "Perugia" if 'ou=perugia' in user.get('dn', '').lower() else "Treviso"
                phone_type = "mobile" if user.get('mobile') else ("tel. fisso" if user.get('phone') else "nessun tel.")
                print_info(f"Generazione firma per {user['email']} (sede: {sede}, {phone_type})...")

                success, _ = self.generate_signatures([user])

                if success:
                    print_success("Firma generata!")

                    sync = input(f"{Colors.YELLOW}Sincronizzare subito su server? (s/N): {Colors.END}").strip().lower()
                    if sync == 's':
                        folder = self.local_path / user['email'].lower()
                        if folder.exists():
                            try:
                                dest = self.server_path / user['email'].lower()
                                if dest.exists():
                                    shutil.rmtree(dest)
                                shutil.copytree(folder, dest)
                                print_success(f"Sincronizzata su server!")
                            except Exception as e:
                                print_error(f"Errore sync: {e}")
            else:
                print_error("Numero non valido")
        except ValueError:
            print_error("Inserisci un numero")

        input(f"\n{Colors.CYAN}Premi Invio per continuare...{Colors.END}")

    def run(self):
        """Avvia l'applicazione"""
        self.show_banner()

        # Login
        print_header("AUTENTICAZIONE ACTIVE DIRECTORY")

        username = input(f"{Colors.WHITE}Username [{Colors.CYAN}sandro.sellaro{Colors.WHITE}]: {Colors.END}").strip() or "sandro.sellaro"

        import getpass
        try:
            password = getpass.getpass(f"{Colors.WHITE}Password: {Colors.END}")
        except:
            password = input(f"{Colors.WHITE}Password: {Colors.END}")

        if not self.connect(username, password):
            input(f"\n{Colors.RED}Premi Invio per uscire...{Colors.END}")
            return

        # Selezione sede ITALIA
        self.current_sede = self.menu_sede()
        nome_sede, base_dn = self.current_sede

        print_info(f"Sede selezionata: {nome_sede}")
        print_info("Caricamento utenti...")

        self.search_users(base_dn)
        print_success(f"Caricati {len(self.users)} utenti")

        # Menu principale
        self.menu_principale()


def main():
    try:
        app = ADSignatureManager()
        app.run()
    except KeyboardInterrupt:
        print(f"\n\n{Colors.YELLOW}Interrotto dall'utente.{Colors.END}")
    except Exception as e:
        print(f"\n{Colors.RED}ERRORE CRITICO: {e}{Colors.END}")
        import traceback
        traceback.print_exc()
        input("\nPremi Invio per uscire...")


if __name__ == "__main__":
    main()
