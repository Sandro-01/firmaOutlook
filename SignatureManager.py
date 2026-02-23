"""
Email Signature Manager - Active Directory -> Outlook
VERSION 5.0 - CARTON GROUP (International)
- Customizable logo from local file
- All locations: Italy (Treviso, Perugia, Verona) + Germany (Schwabach)
- Location-based address differentiation
- Signature NOT set as default (user free to modify)
- Modern UI with progress bar
- Local save + server sync
- Phone/Mobile: shows one only (mobile priority)
- Both phone and mobile shown if both available
"""

import os
import sys
import base64
import shutil
from pathlib import Path
from datetime import datetime

# Windows console colors
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
    print(f"\n{Colors.CYAN}{'='*70}")
    print(f"  {Colors.BOLD}{text}{Colors.END}")
    print(f"{Colors.CYAN}{'='*70}{Colors.END}")


def print_success(text):
    print(f"{Colors.GREEN}+ {text}{Colors.END}")


def print_error(text):
    print(f"{Colors.RED}x {text}{Colors.END}")


def print_warning(text):
    print(f"{Colors.YELLOW}! {text}{Colors.END}")


def print_info(text):
    print(f"{Colors.BLUE}> {text}{Colors.END}")


def print_menu_item(num, icon, text, desc=""):
    if desc:
        print(f"  {Colors.WHITE}{num}.{Colors.END} {icon} {Colors.BOLD}{text}{Colors.END} {Colors.CYAN}- {desc}{Colors.END}")
    else:
        print(f"  {Colors.WHITE}{num}.{Colors.END} {icon} {Colors.BOLD}{text}{Colors.END}")


def progress_bar(current, total, width=40):
    percent = current / total
    filled = int(width * percent)
    bar = '#' * filled + '-' * (width - filled)
    print(f"\r  {Colors.CYAN}[{bar}] {current}/{total} ({percent*100:.0f}%){Colors.END}", end='', flush=True)


def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')


try:
    from ldap3 import Server, Connection, ALL, SUBTREE
except ImportError:
    print_error("Module ldap3 not found!")
    print_info("Install with: pip install ldap3")
    sys.exit(1)


class ADSignatureManager:
    def __init__(self):
        # Carton Group default configuration
        self.ad_server = "adds.cartongrp.com"
        self.domain = "CARTONGRP"
        self.connection = None
        self.users = []

        # Paths
        self.local_path = Path(os.environ.get('TEMP', 'C:\\Temp')) / "Generated_Signatures"
        self.server_path = Path(r"\\DEAZRADS101\Firme")

        # Locations - Italy + Germany
        self.locations = {
            '1': ('Treviso (IT)', 'OU=Treviso,OU=rIT,OU=Client,DC=adds,DC=cartongrp,DC=com'),
            '2': ('Perugia (IT)', 'OU=Perugia,OU=rIT,OU=Client,DC=adds,DC=cartongrp,DC=com'),
            '3': ('Verona (IT)', 'OU=Verona,OU=rIT,OU=Client,DC=adds,DC=cartongrp,DC=com'),
            '4': ('ALL ITALY', 'OU=rIT,OU=Client,DC=adds,DC=cartongrp,DC=com'),
            '5': ('Schwabach (DE)', 'OU=Schwabach,OU=rDE,OU=Client,DC=adds,DC=cartongrp,DC=com'),
            '6': ('ALL GERMANY', 'OU=rDE,OU=Client,DC=adds,DC=cartongrp,DC=com'),
            '7': ('ALL LOCATIONS', 'OU=Client,DC=adds,DC=cartongrp,DC=com'),
        }

        # Addresses by location (detected from OU in Distinguished Name)
        self.addresses = {
            'perugia': 'Via Guido Rossa, 5 | S. Sabina (PG) - Italy',
            'treviso': 'Via Turati, 49 | Paese (TV) - Italy',
            'verona': 'Via Turati, 49 | Paese (TV) - Italy',
            'schwabach': 'Berliner Str. 2 | 91126 Schwabach - Germany',
            'default': 'Via Turati, 49 | Paese (TV) - Italy',
        }

        # Logo - default online, can be overridden with local file
        self.logo_src = 'https://tse2.mm.bing.net/th/id/OIP.wq2AFoiw0fNYotAoONJAlwAAAA'
        self.logo_base64 = None

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
        """Load logo from local file and convert to Base64"""
        try:
            file_path = Path(file_path.strip('"').strip("'"))

            if not file_path.exists():
                print_error(f"File not found: {file_path}")
                return False

            with open(file_path, 'rb') as f:
                image_data = f.read()

            encoded = base64.b64encode(image_data).decode('utf-8')

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

            self.logo_base64 = f"data:{mime_type};base64,{encoded}"
            self.logo_src = self.logo_base64

            size_kb = len(image_data) / 1024
            print_success(f"Logo loaded: {file_path.name} ({size_kb:.1f} KB)")
            self.log(f"Logo loaded: {file_path}")
            return True

        except Exception as e:
            print_error(f"Error loading logo: {e}")
            return False

    def show_banner(self):
        clear_screen()
        print(f"""
{Colors.CYAN}+======================================================================+
|                                                                      |
|  {Colors.WHITE}{Colors.BOLD}    CCCC   AAA   RRRR  TTTTT  OOO  N   N{Colors.CYAN}                              |
|  {Colors.WHITE}{Colors.BOLD}   C      A   A  R   R   T   O   O NN  N{Colors.CYAN}                              |
|  {Colors.WHITE}{Colors.BOLD}   C      AAAAA  RRRR    T   O   O N N N{Colors.CYAN}                              |
|  {Colors.WHITE}{Colors.BOLD}   C      A   A  R  R    T   O   O N  NN{Colors.CYAN}                              |
|  {Colors.WHITE}{Colors.BOLD}    CCCC  A   A  R   R   T    OOO  N   N{Colors.CYAN}                              |
|                                                                      |
|  {Colors.GREEN}Email Signature Manager v5.0 - INTERNATIONAL{Colors.CYAN}                    |
|  {Colors.WHITE}Active Directory -> Outlook Signatures{Colors.CYAN}                          |
|  {Colors.YELLOW}Free signature: users can modify it in Outlook{Colors.CYAN}                  |
|  {Colors.YELLOW}Location-based address (IT: TV/PG | DE: Schwabach){Colors.CYAN}              |
|                                                                      |
+======================================================================+{Colors.END}
""")

    def connect(self, username, password):
        """Connect to Active Directory"""
        print_info(f"Connecting to {self.ad_server}...")

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
                    print_success("Connected to Active Directory")
                    self.log(f"Connected as {username}")
                    return True
                except:
                    continue

            print_error("Invalid credentials")
            return False

        except Exception as e:
            print_error(f"Connection error: {e}")
            self.log(f"Connection error: {e}", "ERROR")
            return False

    def search_users(self, base_dn, filter_text=""):
        """Search users in Active Directory"""
        if not self.connection:
            return []

        if filter_text:
            ldap_filter = f"(&(objectClass=user)(mail=*)(|(sAMAccountName=*{filter_text}*)(displayName=*{filter_text}*)(mail=*{filter_text}*)))"
        else:
            ldap_filter = "(&(objectClass=user)(mail=*))"

        attributes = [
            'sAMAccountName', 'displayName', 'givenName', 'sn',
            'title', 'mail', 'mobile', 'telephoneNumber',
            'department', 'distinguishedName'
        ]

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
                    try:
                        if hasattr(entry, attr_name):
                            val = getattr(entry, attr_name)
                            if val:
                                if hasattr(val, 'value') and val.value:
                                    return str(val.value).strip()
                                elif hasattr(val, 'values') and val.values:
                                    return str(val.values[0]).strip()
                                else:
                                    val_str = str(val).strip()
                                    if val_str.startswith('[') and val_str.endswith(']'):
                                        val_str = val_str[1:-1].strip("'\"")
                                    if val_str and val_str != '[]':
                                        return val_str
                    except:
                        pass
                    return ''

                user = {
                    'username': get_attr('sAMAccountName'),
                    'display_name': get_attr('displayName'),
                    'first_name': get_attr('givenName'),
                    'last_name': get_attr('sn'),
                    'title': get_attr('title') or 'Employee',
                    'email': get_attr('mail'),
                    'mobile': get_attr('mobile'),
                    'phone': get_attr('telephoneNumber'),
                    'department': get_attr('department'),
                    'dn': get_attr('distinguishedName')
                }

                if user['email'] and '@' in user['email']:
                    self.users.append(user)

            self.users.sort(key=lambda x: x['email'])
            return self.users

        except Exception as e:
            print_error(f"Search error: {e}")
            self.log(f"Search error: {e}", "ERROR")
            return []

    def format_name(self, user):
        """Format name as: First Last"""
        if user.get('first_name') and user.get('last_name'):
            return f"{user['first_name']} {user['last_name']}"

        display = user['display_name']
        # Remove common suffixes
        for suffix in ['(Europoligrafico)', '(Carton Group)', '(Leupold)', '(NoblePac)']:
            display = display.replace(suffix, '').strip()

        if ',' in display:
            parts = [p.strip() for p in display.split(',')]
            if len(parts) == 2:
                return f"{parts[1]} {parts[0]}"

        return display

    def get_location(self, user):
        """Detect user location from Distinguished Name"""
        dn = user.get('dn', '').lower()

        if 'ou=perugia' in dn:
            return 'perugia'
        elif 'ou=treviso' in dn:
            return 'treviso'
        elif 'ou=verona' in dn:
            return 'verona'
        elif 'ou=schwabach' in dn:
            return 'schwabach'
        elif 'ou=rit' in dn:
            return 'treviso'  # Default for Italy
        elif 'ou=rde' in dn:
            return 'schwabach'  # Default for Germany
        return 'default'

    def get_address(self, user):
        """Return correct address based on user location"""
        location = self.get_location(user)
        return self.addresses.get(location, self.addresses['default'])

    def get_location_display(self, user):
        """Return display name for location"""
        location = self.get_location(user)
        location_names = {
            'perugia': 'Perugia',
            'treviso': 'Treviso',
            'verona': 'Verona',
            'schwabach': 'Schwabach',
            'default': 'N/A'
        }
        return location_names.get(location, 'N/A')

    def get_phone_display(self, user):
        """
        Return phone number to display.
        Priority: Mobile > Landline
        Returns tuple (prefix, number) or (None, None) if none available
        """
        if user.get('mobile'):
            return ('M', user['mobile'])
        elif user.get('phone'):
            return ('T', user['phone'])
        return (None, None)

    def generate_html(self, user):
        """Generate HTML signature"""
        name = self.format_name(user)
        address = self.get_address(user)

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
                    {address}
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
        """Generate text signature"""
        name = self.format_name(user)
        address = self.get_address(user)

        phone_prefix, phone_number = self.get_phone_display(user)
        phone_line = ""
        if phone_number:
            phone_line = f"\n{phone_prefix}: {phone_number}"

        return f"""{name}
{user['title']}{phone_line}

{address}
www.carton-group.com
Follow us: Carton Group"""

    def generate_signatures(self, users=None):
        """Generate signatures for selected users"""
        if users is None:
            users = self.users

        if not users:
            print_warning("No users to process")
            return 0, 0

        self.local_path.mkdir(parents=True, exist_ok=True)

        print_info(f"Generating {len(users)} signatures in {self.local_path}")
        print()

        success = 0
        errors = 0

        for i, user in enumerate(users, 1):
            progress_bar(i, len(users))

            try:
                email = user['email'].lower()
                user_folder = self.local_path / email
                user_folder.mkdir(parents=True, exist_ok=True)

                (user_folder / "Firma_Aziendale.htm").write_text(
                    self.generate_html(user), encoding='utf-8')
                (user_folder / "Firma_Aziendale.txt").write_text(
                    self.generate_txt(user), encoding='utf-8')

                success += 1
                location = self.get_location_display(user)
                phone_type = "mobile" if user.get('mobile') else (
                    "phone" if user.get('phone') else "none")
                self.log(f"Generated: {email} (location: {location}, tel: {phone_type})")

            except Exception as e:
                errors += 1
                self.log(f"Error {user['email']}: {e}", "ERROR")

        print()
        return success, errors

    def sync_to_server(self):
        """Sync local signatures to server"""
        if not self.local_path.exists():
            print_error("No local signatures to sync")
            return 0, 0

        print_info(f"Checking access to {self.server_path}...")

        if not self.server_path.exists():
            print_error(f"Server unreachable: {self.server_path}")
            print_warning("Make sure you are connected to the corporate network")
            return 0, 0

        folders = [f for f in self.local_path.iterdir()
                   if f.is_dir() and '@' in f.name]

        if not folders:
            print_warning("No signature folders found locally")
            return 0, 0

        print_info(f"Syncing {len(folders)} signatures to server...")
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
                self.log(f"Synced: {folder.name}")

            except Exception as e:
                errors += 1
                self.log(f"Sync error {folder.name}: {e}", "ERROR")

        print()
        return success, errors

    def show_users_table(self):
        """Display users table"""
        if not self.users:
            print_warning("No users loaded")
            return

        print(f"\n{Colors.CYAN}{'-'*115}{Colors.END}")
        print(f"{Colors.BOLD}{'#':<4} {'Email':<35} {'Name':<22} {'Title':<15} {'Location':<12} {'Phone':<15}{Colors.END}")
        print(f"{Colors.CYAN}{'-'*115}{Colors.END}")

        for i, user in enumerate(self.users, 1):
            name = self.format_name(user)[:21]
            title = (user['title'] or 'N/A')[:14]
            email = user['email'][:34]
            location = self.get_location_display(user)

            phone_prefix, phone_number = self.get_phone_display(user)
            tel_display = f"{phone_prefix}:{phone_number[-8:]}" if phone_number else "-"

            print(f"{i:<4} {email:<35} {name:<22} {title:<15} {location:<12} {tel_display:<15}")

        print(f"{Colors.CYAN}{'-'*115}{Colors.END}")
        print(f"{Colors.GREEN}Total: {len(self.users)} users{Colors.END}")
        print(f"{Colors.CYAN}Phone: M=Mobile, T=Landline (Mobile takes priority){Colors.END}\n")

    def menu_location(self):
        """Location selection menu"""
        print_header("SELECT LOCATION")

        for key, (name, _) in self.locations.items():
            icon = "[IT]" if 'IT' in name or 'ITALY' in name else (
                "[DE]" if 'DE' in name or 'GERMANY' in name else "[ALL]")
            print_menu_item(key, icon, name)

        print()
        choice = input(f"{Colors.WHITE}Choose (1-7) [{Colors.CYAN}7{Colors.WHITE}]: {Colors.END}").strip() or '7'

        if choice in self.locations:
            return self.locations[choice]
        return self.locations['7']

    def menu_logo(self):
        """Logo configuration menu"""
        print_header("LOGO CONFIGURATION")

        print_menu_item(1, "[WEB]", "Use online logo (default)", "Current URL")
        print_menu_item(2, "[FILE]", "Load logo from file", "PNG, JPG, GIF, BMP")
        print_menu_item(3, "[BACK]", "Back to main menu", "")

        print()
        if self.logo_base64:
            print_success("Custom logo already loaded")
        else:
            print_info(f"Current logo: {self.logo_src[:50]}...")

        print()
        choice = input(f"{Colors.WHITE}Choose (1-3): {Colors.END}").strip()

        if choice == '2':
            print()
            print_info("Enter the full path to the logo file")
            print_info("Example: C:\\Users\\user\\Desktop\\logo.png")
            print()
            file_path = input(f"{Colors.WHITE}File path: {Colors.END}").strip()

            if file_path:
                self.load_logo_from_file(file_path)

            input(f"\n{Colors.CYAN}Press Enter to continue...{Colors.END}")

    def menu_main(self):
        """Main menu"""
        while True:
            print_header("MAIN MENU")

            logo_status = "+ custom" if self.logo_base64 else "default online"

            print_menu_item(1, "[SEARCH]", "Search users", "filter by name/email")
            print_menu_item(2, "[LIST]", "Show users", f"{len(self.users)} loaded")
            print_menu_item(3, "[LOGO]", "Configure logo", logo_status)
            print_menu_item(4, "[GEN]", "Generate ALL signatures", "save locally")
            print_menu_item(5, "[SYNC]", "Sync to SERVER", "copy to \\\\DEAZRADS101\\Firme")
            print_menu_item(6, "[AUTO]", "AUTOMATIC", "generate + sync")
            print_menu_item(7, "[USER]", "Generate for specific user", "manual selection")
            print_menu_item(0, "[EXIT]", "Exit", "")

            print()
            print_info("Addresses: IT = Via Turati (TV) / Via Guido Rossa (PG)")
            print_info("           DE = Berliner Str. 2 (Schwabach)")
            print_info("Phone: shows Mobile if available, otherwise Landline")
            print()
            choice = input(f"{Colors.WHITE}Choose: {Colors.END}").strip()

            if choice == '1':
                self.action_search()
            elif choice == '2':
                self.show_users_table()
                input(f"\n{Colors.CYAN}Press Enter to continue...{Colors.END}")
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
                print_info("Goodbye!")
                break

    def action_search(self):
        """Search users with filter"""
        print_header("USER SEARCH")
        filter_text = input(f"{Colors.WHITE}Filter (empty = all): {Colors.END}").strip()

        location_name, base_dn = self.current_location
        print_info(f"Searching in {location_name}...")

        self.search_users(base_dn, filter_text)
        print_success(f"Found {len(self.users)} users")

        if self.users and len(self.users) <= 20:
            self.show_users_table()

    def action_generate(self):
        """Generate all signatures locally"""
        if not self.users:
            print_warning("Search for users first (option 1)")
            return

        print_header("SIGNATURE GENERATION")
        success, errors = self.generate_signatures()

        print()
        print_success(f"Generated: {success}")
        if errors:
            print_error(f"Errors: {errors}")

        print_info(f"Saved in: {self.local_path}")
        input(f"\n{Colors.CYAN}Press Enter to continue...{Colors.END}")

    def action_sync(self):
        """Sync to server"""
        print_header("SERVER SYNC")

        confirm = input(f"{Colors.YELLOW}Sync to {self.server_path}? (y/N): {Colors.END}").strip().lower()

        if confirm == 'y':
            success, errors = self.sync_to_server()

            print()
            if success:
                print_success(f"Synced: {success}")
            if errors:
                print_error(f"Errors: {errors}")
        else:
            print_warning("Operation cancelled")

        input(f"\n{Colors.CYAN}Press Enter to continue...{Colors.END}")

    def action_auto(self):
        """Generate and sync automatically"""
        if not self.users:
            print_warning("Search for users first (option 1)")
            return

        print_header("AUTOMATIC MODE")
        print_info(f"{len(self.users)} signatures will be generated and synced to server")
        print_warning("Signatures will NOT be set as default")
        print_info("Users can freely modify them in Outlook")
        print_info("Addresses: IT = Via Turati (TV) / Via Guido Rossa (PG)")
        print_info("           DE = Berliner Str. 2 (Schwabach)")
        print_info("Phone: Mobile (priority) or Landline")

        confirm = input(f"\n{Colors.YELLOW}Proceed? (y/N): {Colors.END}").strip().lower()

        if confirm != 'y':
            print_warning("Operation cancelled")
            return

        print()
        print_info("STEP 1/2: Generating signatures...")
        gen_ok, gen_err = self.generate_signatures()
        print()
        print_success(f"Generated: {gen_ok}")

        print()
        print_info("STEP 2/2: Syncing to server...")
        sync_ok, sync_err = self.sync_to_server()
        print()

        print_header("SUMMARY")
        print_success(f"Signatures generated: {gen_ok}")
        print_success(f"Signatures synced: {sync_ok}")
        if gen_err or sync_err:
            print_error(f"Total errors: {gen_err + sync_err}")

        print()
        print_info("Signatures are now available for GPO distribution")
        print_info("Users can modify them in Outlook > File > Options > Mail > Signatures")
        print_info(f"Log: {self.log_file}")
        input(f"\n{Colors.CYAN}Press Enter to continue...{Colors.END}")

    def action_single_user(self):
        """Generate signature for specific user"""
        if not self.users:
            print_warning("Search for users first (option 1)")
            return

        self.show_users_table()

        try:
            num = int(input(f"{Colors.WHITE}User number: {Colors.END}").strip())
            if 1 <= num <= len(self.users):
                user = self.users[num - 1]
                location = self.get_location_display(user)
                phone_type = "mobile" if user.get('mobile') else (
                    "landline" if user.get('phone') else "no phone")
                print_info(
                    f"Generating signature for {user['email']} "
                    f"(location: {location}, {phone_type})...")

                success, _ = self.generate_signatures([user])

                if success:
                    print_success("Signature generated!")

                    sync = input(
                        f"{Colors.YELLOW}Sync to server now? (y/N): {Colors.END}"
                    ).strip().lower()
                    if sync == 'y':
                        folder = self.local_path / user['email'].lower()
                        if folder.exists():
                            try:
                                dest = self.server_path / user['email'].lower()
                                if dest.exists():
                                    shutil.rmtree(dest)
                                shutil.copytree(folder, dest)
                                print_success("Synced to server!")
                            except Exception as e:
                                print_error(f"Sync error: {e}")
            else:
                print_error("Invalid number")
        except ValueError:
            print_error("Enter a number")

        input(f"\n{Colors.CYAN}Press Enter to continue...{Colors.END}")

    def run(self):
        """Start the application"""
        self.show_banner()

        # Login
        print_header("ACTIVE DIRECTORY AUTHENTICATION")

        username = input(
            f"{Colors.WHITE}Username [{Colors.CYAN}sandro.sellaro{Colors.WHITE}]: {Colors.END}"
        ).strip() or "sandro.sellaro"

        import getpass
        try:
            password = getpass.getpass(f"{Colors.WHITE}Password: {Colors.END}")
        except:
            password = input(f"{Colors.WHITE}Password: {Colors.END}")

        if not self.connect(username, password):
            input(f"\n{Colors.RED}Press Enter to exit...{Colors.END}")
            return

        # Location selection
        self.current_location = self.menu_location()
        location_name, base_dn = self.current_location

        print_info(f"Selected location: {location_name}")
        print_info("Loading users...")

        self.search_users(base_dn)
        print_success(f"Loaded {len(self.users)} users")

        # Main menu
        self.menu_main()


def main():
    try:
        app = ADSignatureManager()
        app.run()
    except KeyboardInterrupt:
        print(f"\n\n{Colors.YELLOW}Interrupted by user.{Colors.END}")
    except Exception as e:
        print(f"\n{Colors.RED}CRITICAL ERROR: {e}{Colors.END}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()
