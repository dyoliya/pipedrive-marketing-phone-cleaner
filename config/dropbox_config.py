import os
import time
import requests
import dropbox
import sys
from dotenv import load_dotenv


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

dotenv_path = resource_path(os.path.join("config", ".env"))

load_dotenv(dotenv_path)

DROPBOX_CONVERSION_APP_KEY = os.getenv("DROPBOX_CONVERSION_APP_KEY")
DROPBOX_CONVERSION_APP_SECRET = os.getenv("DROPBOX_CONVERSION_APP_SECRET")
DROPBOX_CONVERSION_REFRESH_TOKEN = os.getenv("DROPBOX_CONVERSION_REFRESH_TOKEN")

if not DROPBOX_CONVERSION_REFRESH_TOKEN:
    raise Exception("Missing DROPBOX_CONVERSION_REFRESH_TOKEN in environment")
if not DROPBOX_CONVERSION_APP_KEY:
    raise Exception("Missing DROPBOX_CONVERSION_APP_KEY in environment")
if not DROPBOX_CONVERSION_APP_SECRET:
    raise Exception("Missing DROPBOX_CONVERSION_APP_SECRET in environment")

_access_token = None
_token_expiry = 0

def refresh_access_token():
    global _access_token, _token_expiry

    if _access_token and time.time() < _token_expiry - 60:
        return _access_token

    url = "https://api.dropbox.com/oauth2/token"
    data = {
        "grant_type": "refresh_token",
        "refresh_token": DROPBOX_CONVERSION_REFRESH_TOKEN,
        "client_id": DROPBOX_CONVERSION_APP_KEY,
        "client_secret": DROPBOX_CONVERSION_APP_SECRET,
    }

    resp = requests.post(url, data=data)
    if resp.status_code != 200:
        raise Exception(f"Failed to refresh Dropbox token: {resp.text}")

    token_data = resp.json()
    _access_token = token_data["access_token"]
    expires_in = token_data.get("expires_in", 14400)
    _token_expiry = time.time() + expires_in
    return _access_token

def get_dropbox_client():
    access_token = refresh_access_token()
    return dropbox.Dropbox(access_token)
