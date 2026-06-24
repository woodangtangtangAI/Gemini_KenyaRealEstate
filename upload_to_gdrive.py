import os
import sys
import json
import re
import requests

# Set stdout encoding to UTF-8
try:
    sys.stdout.reconfigure(encoding='utf-8')
except:
    pass

PARENT_FOLDER_ID = "1CZ6bJC8e7AZZvvEej3hVR_4i0p9UCsYJ"

def get_latest_run_folder(base_path):
    run_folders = [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d)) and re.match(r'^\d{8}\([a-zA-Z]{3}\)$', d)]
    if not run_folders:
        return None
    # Sort folders by name (since it starts with YYYYMMDD, sorting alphabetically finds the latest)
    run_folders.sort(reverse=True)
    return run_folders[0]

def refresh_access_token(token_data):
    refresh_token = token_data.get("refresh_token")
    client_id = token_data.get("client_id")
    client_secret = token_data.get("client_secret")
    token_uri = token_data.get("token_uri", "https://oauth2.googleapis.com/token")

    if not refresh_token or not client_id or not client_secret:
        print("❌ Error: Missing refresh_token, client_id, or client_secret in token.json")
        return None

    print("Refreshing Google Drive access token...")
    res = requests.post(token_uri, data={
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "client_id": client_id,
        "client_secret": client_secret
    })

    if res.status_code != 200:
        print(f"❌ Error refreshing token: {res.status_code} - {res.text}")
        return None

    return res.json().get("access_token")

def get_or_create_gdrive_folder(folder_name, headers):
    # Check if folder already exists in the parent folder
    q = f"name = '{folder_name}' and '{PARENT_FOLDER_ID}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    url = f"https://www.googleapis.com/drive/v3/files?q={requests.utils.quote(q)}&fields=files(id)"
    
    r = requests.get(url, headers=headers)
    if r.status_code == 200:
        files = r.json().get("files", [])
        if files:
            print(f"📁 Google Drive folder '{folder_name}' already exists (ID: {files[0]['id']})")
            return files[0]["id"]
    
    # Create the folder
    print(f"📁 Creating new Google Drive folder '{folder_name}'...")
    create_url = "https://www.googleapis.com/drive/v3/files"
    metadata = {
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [PARENT_FOLDER_ID]
    }
    r = requests.post(create_url, headers=headers, json=metadata)
    if r.status_code in [200, 201]:
        folder_id = r.json().get("id")
        print(f"📁 Created folder successfully (ID: {folder_id})")
        return folder_id
    else:
        print(f"❌ Failed to create folder: {r.status_code} - {r.text}")
        return None

def upload_file_to_gdrive(local_file_path, gdrive_folder_id, headers):
    filename = os.path.basename(local_file_path)
    
    # Check if file already exists in the folder
    q = f"name = '{filename}' and '{gdrive_folder_id}' in parents and trashed = false"
    url = f"https://www.googleapis.com/drive/v3/files?q={requests.utils.quote(q)}&fields=files(id)"
    r = requests.get(url, headers=headers)
    
    if r.status_code == 200:
        files = r.json().get("files", [])
        for f in files:
            # Delete existing file to avoid duplicate copies
            print(f"   🗑️ Deleting old version of '{filename}' (ID: {f['id']})...")
            requests.delete(f"https://www.googleapis.com/drive/v3/files/{f['id']}", headers=headers)

    # Upload file using multipart upload
    print(f"   📤 Uploading '{filename}'...")
    upload_url = "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart"
    metadata = {
        "name": filename,
        "parents": [gdrive_folder_id]
    }
    
    files_payload = {
        'metadata': (None, json.dumps(metadata), 'application/json; charset=UTF-8'),
        'file': (filename, open(local_file_path, 'rb'))
    }
    
    # Do not pass Content-Type header manually, requests will set multipart boundary
    upload_headers = headers.copy()
    r = requests.post(upload_url, headers=upload_headers, files=files_payload)
    
    if r.status_code in [200, 201]:
        print(f"   ✅ Successfully uploaded '{filename}' (ID: {r.json().get('id')})")
        return True
    else:
        print(f"   ❌ Failed to upload '{filename}': {r.status_code} - {r.text}")
        return False

def main():
    base_path = os.path.dirname(os.path.abspath(__file__))
    
    # 1. Look for token.json in base path, parent path, or repository root
    token_file = os.path.join(base_path, "token.json")
    if not os.path.exists(token_file):
        # Check parent directory
        token_file = os.path.join(os.path.dirname(base_path), "token.json")
    if not os.path.exists(token_file):
        # Fallback to current working directory
        token_file = "token.json"
        
    if not os.path.exists(token_file):
        print(f"❌ Error: token.json not found")
        sys.exit(1)
        
    with open(token_file, "r", encoding="utf-8") as f:
        token_data = json.load(f)
        
    # 2. Get Access Token
    access_token = refresh_access_token(token_data)
    if not access_token:
        print("❌ Error: Failed to obtain access token.")
        sys.exit(1)
        
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    # 3. Find latest weekly run folder
    latest_folder = get_latest_run_folder(base_path)
    if not latest_folder:
        print("❌ Error: No run folders (e.g. YYYYMMDD(Day of week)) found to upload.")
        sys.exit(1)
        
    local_folder_path = os.path.join(base_path, latest_folder)
    print(f"🚀 Found local run folder: '{latest_folder}'")

    # 4. Get or Create target folder in Google Drive
    gdrive_folder_id = get_or_create_gdrive_folder(latest_folder, headers)
    if not gdrive_folder_id:
        sys.exit(1)

    # 5. Upload all files inside the local folder
    for filename in os.listdir(local_folder_path):
        local_file_path = os.path.join(local_folder_path, filename)
        if os.path.isfile(local_file_path):
            upload_file_to_gdrive(local_file_path, gdrive_folder_id, headers)

    print("\n🎉 Google Drive upload complete!")

if __name__ == "__main__":
    main()
