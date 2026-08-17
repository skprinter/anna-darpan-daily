import os
import requests


LOGIN_URL = (
    "https://adbackend.annadarpan.in/"
    "prdannadarpan.in/security/passwordLogin"
)

username = os.getenv("ANNA_DARPAN_USERNAME", "")
password = os.getenv("ANNA_DARPAN_PASSWORD", "")

if not username:
    raise RuntimeError("ANNA_DARPAN_USERNAME missing")

if not password:
    raise RuntimeError("ANNA_DARPAN_PASSWORD missing")


payload = {
    "username": username,
    "password": password,
    "clientId": "ad-fci",

    # IMPORTANT:
    # Do NOT invent or reuse an old CAPTCHA token here.
    # We first need to determine whether the endpoint
    # accepts authentication without a browser-generated token.
}

headers = {
    "accept": "application/json, text/plain, */*",
    "content-type": "application/json",
    "origin": "https://www.annadarpan.in",
    "referer": "https://www.annadarpan.in/",
    "user-agent": "Mozilla/5.0",
}

print("=" * 60)
print("ANNA DARPAN LOGIN TEST")
print("=" * 60)

response = requests.post(
    LOGIN_URL,
    headers=headers,
    json=payload,
    timeout=60,
)

print("HTTP STATUS:", response.status_code)

try:
    data = response.json()
except Exception:
    print("Response is not JSON")
    print(response.text[:1000])
    raise SystemExit(1)

# Never print access_token / refresh_token.
print("Response keys:", list(data.keys()))

if response.status_code == 200 and data.get("access_token"):
    print("LOGIN SUCCESS")
    print("Fresh access token received.")
    print("expires_in:", data.get("expires_in"))
    print("refresh_expires_in:", data.get("refresh_expires_in"))
else:
    print("LOGIN NOT SUCCESSFUL")
    print("Response:", {
        k: v for k, v in data.items()
        if k not in ["access_token", "refresh_token"]
    })
