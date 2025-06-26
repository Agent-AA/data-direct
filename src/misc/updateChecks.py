import requests
import sys
import os
from misc import ui
import zipfile
import tempfile

GITHUB_API = f"https://api.github.com/repositories/1002552720/releases/latest"

def check_for_update(current_version):
    
    try:
        print('Checking for latest version...')
        response = requests.get(GITHUB_API, timeout=10)
        response.raise_for_status()
        data = response.json()
        latest_version = data['tag_name'].lstrip('v')
        if latest_version == current_version:
            ui.print_success('This software is up to date.')
            return  # No update needed

        ui.print_warning(f"A new version of this software is available.")
        ui.prompt_user('Press "U" to update now. Otherwise, the program will continue.')
        keypress = ui.pause()
        if keypress != b'u' and keypress != b'U':
            return

        # Find the first asset (assuming it's a zip file)
        asset = next((a for a in data['assets'] if a['name'].endswith('.zip')), None)
        if not asset:
            ui.print_error("No downloadable asset found for the latest release.")
            return

        download_url = asset['browser_download_url']
        print(f"Downloading update from {download_url} ...")
        with requests.get(download_url, stream=True) as r:
            r.raise_for_status()
            with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp_file:
                for chunk in r.iter_content(chunk_size=8192):
                    tmp_file.write(chunk)
                tmp_zip_path = tmp_file.name

        print("Extracting update...")
        with zipfile.ZipFile(tmp_zip_path, 'r') as zip_ref:
            # Extract the executable to the current working directory
            for member in zip_ref.namelist():
                if member.endswith('.exe'):
                    exe_path = os.path.join(os.getcwd(), os.path.basename(member))
                    with open(exe_path, 'wb') as out_file:
                        out_file.write(zip_ref.read(member))
                    print(f"Extracted new executable to {exe_path}")
                    break
            else:
                ui.print_error("No executable found in the update package.")
                os.remove(tmp_zip_path)
                return

        print("Launching installer...")
        os.remove(tmp_zip_path)
        os.startfile(exe_path)
        sys.exit(0)

    except Exception as e:
        ui.print_error(f"Update check failed: {e}")