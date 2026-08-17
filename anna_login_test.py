import os
import sys
import json
import base64
import requests

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric import padding


# ============================================================
# ANNA DARPAN LOGIN TEST
# ============================================================

LOGIN_URL = (
    "https://adbackend.annadarpan.in"
    "/prdannadarpan.in/security/passwordLogin"
)

CLIENT_ID = "ad-fci"
ORIGIN = "https://www.annadarpan.in"
REFERER = "https://www.annadarpan.in/"

# Public key taken from Anna Darpan frontend configuration.
PUBLIC_KEY_B64 = (
    "MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEArzBjfgBHAMquAmXWt5Ue0kg3tir6Q/"
    "WizQJrJSX2JKoQduBVEzGS7Ly9/AG9m/NpJMhmEMW6Az3KXKcKnnkgaLYqgIVSAqysekpz3U7m"
    "LO5TsAA9kEWz6edLTaZoI4qFFeY5wHafuBQvWmAQKqICx7ZN04CwyrSk7jLilAX+IGMASBB7Qv"
    "35jiDvzPWr2VKVwfjUDyELgNi0eY5WIFmNvEDB8R+6Gsz8+PY4rmWYmoEjiCpKpDSMz2j/AzMH"
    "dh8+qdIdGU1qZJLpPGHgWFjp+HomTU864w8J2GCt2OJYswQZNFQoUQuzrDki9g5r8CYfkW2Wrq"
    "XRLRXHvEzrydpPawIDAQAB"
)


def get_required_env(name):
    value = os.environ.get(name)

    if not value:
        print(f"ERROR: GitHub Secret/Environment variable missing: {name}")
        sys.exit(1)

    return value


def build_public_key():
    """
    Convert Base64 DER RSA public key into PEM format.
    """

    der = base64.b64decode(PUBLIC_KEY_B64)

    pem = (
        b"-----BEGIN PUBLIC KEY-----\n"
        + base64.encodebytes(der)
        + b"-----END PUBLIC KEY-----\n"
    )

    return serialization.load_pem_public_key(pem)


def encrypt_password(password):
    """
    Anna Darpan frontend uses RSA encryption.
    JSEncrypt-style encryption uses RSA PKCS#1 v1.5 padding.
    """

    public_key = build_public_key()

    encrypted = public_key.encrypt(
        password.encode("utf-8"),
        padding.PKCS1v15()
    )

    return base64.b64encode(encrypted).decode("utf-8")


def write_github_env(name, value):
    """
    Make token available to later GitHub Actions steps
    without printing the secret.
    """

    github_env = os.environ.get("GITHUB_ENV")

    if not github_env:
        return

    with open(github_env, "a", encoding="utf-8") as f:
        f.write(f"{name}<<EOF\n")
        f.write(value)
        f.write("\nEOF\n")


def write_github_output(name, value):
    """
    Write non-sensitive status/output values to GitHub Actions.
    """

    github_output = os.environ.get("GITHUB_OUTPUT")

    if not github_output:
        return

    with open(github_output, "a", encoding="utf-8") as f:
        f.write(f"{name}={value}\n")


def main():

    print("=" * 60)
    print("ANNA DARPAN LOGIN TEST")
    print("=" * 60)

    username = get_required_env("ANNA_DARPAN_USERNAME")
    password = get_required_env("ANNA_DARPAN_PASSWORD")

    # CAPTCHA token must come from the legitimate Anna Darpan
    # authentication flow. This script does NOT bypass CAPTCHA.
    captcha_token = get_required_env("ANNA_DARPAN_CAPTCHA_TOKEN")

    print(f"Username: {username}")
    print("Password: [HIDDEN]")
    print("CAPTCHA token: [HIDDEN]")

    # --------------------------------------------------------
    # Encrypt password exactly before sending it
    # --------------------------------------------------------

    try:
        encrypted_password = encrypt_password(password)

    except Exception as e:
        print("ERROR: Password encryption failed")
        print(type(e).__name__, str(e))
        sys.exit(1)

    print("Password encryption: SUCCESS")
    print(
        "Encrypted password length:",
        len(encrypted_password)
    )

    # --------------------------------------------------------
    # Request payload
    # --------------------------------------------------------

    payload = {
        "username": username,
        "password": encrypted_password,
        "clientId": CLIENT_ID,
        "tokenCaptcha": captcha_token,
    }

    headers = {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "Origin": ORIGIN,
        "Referer": REFERER,
        "User-Agent": (
            "Mozilla/5.0 "
            "(Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/151.0.0.0 Safari/537.36"
        ),
    }

    print()
    print("Sending login request...")
    print("URL:", LOGIN_URL)

    # --------------------------------------------------------
    # Login
    # --------------------------------------------------------

    try:

        response = requests.post(
            LOGIN_URL,
            json=payload,
            headers=headers,
            timeout=30,
        )

    except requests.RequestException as e:
        print()
        print("LOGIN REQUEST FAILED")
        print(type(e).__name__, str(e))
        sys.exit(1)

    print()
    print("HTTP STATUS:", response.status_code)

    # --------------------------------------------------------
    # Parse response
    # --------------------------------------------------------

    try:
        data = response.json()

    except ValueError:
        print("Response is not JSON")
        print(response.text[:1000])
        sys.exit(1)

    if response.status_code != 200:

        print()
        print("LOGIN FAILED")

        # Do not print credentials/tokens.
        if isinstance(data, dict):

            for key, value in data.items():

                if key.lower() in (
                    "access_token",
                    "refresh_token",
                    "token",
                    "password",
                ):
                    print(f"{key}: [HIDDEN]")
                else:
                    print(f"{key}: {value}")

        else:
            print(data)

        sys.exit(1)

    # --------------------------------------------------------
    # Successful login
    # --------------------------------------------------------

    access_token = data.get("access_token")
    refresh_token = data.get("refresh_token")

    if not access_token:

        print()
        print("HTTP 200 received, but access_token is missing.")
        print("Response keys:", list(data.keys()))
        sys.exit(1)

    print()
    print("=" * 60)
    print("LOGIN SUCCESS")
    print("=" * 60)

    print("Token type:", data.get("token_type"))
    print("Expires in:", data.get("expires_in"), "seconds")
    print("Refresh expires in:", data.get("refresh_expires_in"), "seconds")
    print("Access token: [RECEIVED]")
    print("Refresh token:", "[RECEIVED]" if refresh_token else "[NOT RECEIVED]")

    # --------------------------------------------------------
    # Send tokens to following GitHub Actions steps.
    # They are NOT printed.
    # --------------------------------------------------------

    write_github_env(
        "ANNA_DARPAN_ACCESS_TOKEN",
        access_token
    )

    if refresh_token:
        write_github_env(
            "ANNA_DARPAN_REFRESH_TOKEN",
            refresh_token
        )

    write_github_output("login_success", "true")

    print()
    print("GitHub Actions environment updated.")
    print("ANNA_DARPAN_ACCESS_TOKEN is ready for next step.")
    print()
    print("SUCCESS")


if __name__ == "__main__":
    main()
