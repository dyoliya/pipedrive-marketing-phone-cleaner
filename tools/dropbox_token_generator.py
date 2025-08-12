'''
1) Builds Dropbox OAuth authorization URL with your app key and redirect URI.
2) Opens the URL in your default web browser for you to authorize the app.
3) Starts a temporary local HTTP server on your machine to catch Dropbox’s redirect with the code.
4) Extracts the code from the request URL sent by Dropbox after you authorize.
5) Uses the code to request access and refresh tokens from Dropbox’s API.
6) Saves the tokens to dropbox_tokens.json.
'''

import http.server
import socketserver
import webbrowser
import requests
import urllib.parse
import json

APP_KEY = "" # dropbox app key
APP_SECRET = "" # dropbox secret key
REDIRECT_URI = "http://localhost:8080"
TOKEN_FILE = "dropbox_tokens.json"

# Global place to store code once received
auth_code = None

class OAuthHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        global auth_code
        parsed_path = urllib.parse.urlparse(self.path)
        params = urllib.parse.parse_qs(parsed_path.query)
        if 'code' in params:
            auth_code = params['code'][0]
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"<html><body><h1>Authorization complete!</h1>You can close this window.</body></html>")
        else:
            self.send_response(400)
            self.end_headers()
            self.wfile.write(b"Missing code parameter")

def run_local_server():
    with socketserver.TCPServer(("localhost", 8080), OAuthHandler) as httpd:
        httpd.handle_request()

def main():
    global auth_code

    # Step 1: Build the authorization URL
    auth_url = (
        f"https://www.dropbox.com/oauth2/authorize?"
        f"client_id={APP_KEY}&"
        f"response_type=code&"
        f"redirect_uri={urllib.parse.quote(REDIRECT_URI)}&"
        f"token_access_type=offline"
    )

    print("Opening browser for Dropbox authorization...")
    webbrowser.open(auth_url)

    # Step 2: Start local server to catch the redirect with code
    print("Waiting for authorization code...")
    run_local_server()

    if not auth_code:
        print("Failed to get authorization code.")
        return

    print(f"Received code: {auth_code}")

    # Step 3: Exchange code for tokens
    token_url = "https://api.dropbox.com/oauth2/token"
    data = {
        "code": auth_code,
        "grant_type": "authorization_code",
        "client_id": APP_KEY,
        "client_secret": APP_SECRET,
        "redirect_uri": REDIRECT_URI,
    }

    response = requests.post(token_url, data=data)
    if response.status_code != 200:
        print("Failed to get tokens:")
        print(response.text)
        return

    tokens = response.json()
    print("Access token:", tokens["access_token"])
    print("Refresh token:", tokens.get("refresh_token", "No refresh token provided"))

    # Save tokens to a file
    with open(TOKEN_FILE, "w") as f:
        json.dump(tokens, f, indent=2)

    print(f"Tokens saved to {TOKEN_FILE}")

if __name__ == "__main__":
    main()
