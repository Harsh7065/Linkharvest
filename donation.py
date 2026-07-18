"""
donation.py

Generates a UPI "Donate" QR code at runtime. No personal details (name,
bank, UPI ID) are ever displayed as text anywhere in the app's UI —
only the QR image itself is shown, with a generic label ("Support
Development") instead of the real account name.

Note on privacy limits: UPI is designed so that whoever actually scans
the code and opens it in their payment app WILL see the verified
account holder's name before paying — that's a bank-level anti-fraud
check built into UPI itself, and no app (including this one) can turn
it off or hide it at that stage. This module only controls what is
shown inside this app's own window.
"""
import os
import io

import qrcode
from PIL import Image

# Your UPI ID — only used to build the payment QR, never shown as text in the UI.
UPI_ID = "devendrasharma41788-2@okicici"
DISPLAY_NAME = "Support Development"  # generic label sent in the UPI intent


def build_upi_uri(amount: str = "") -> str:
    """Builds a upi://pay deep link. Leaving amount blank lets the payer
    choose how much to donate."""
    uri = f"upi://pay?pa={UPI_ID}&pn={DISPLAY_NAME.replace(' ', '%20')}&cu=INR"
    if amount:
        uri += f"&am={amount}"
    return uri


def generate_qr_image(save_path: str, amount: str = "") -> str:
    """Generates the donation QR PNG and returns the file path."""
    uri = build_upi_uri(amount)
    qr = qrcode.QRCode(border=2, box_size=8)
    qr.add_data(uri)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    os.makedirs(os.path.dirname(save_path), exist_ok=True)
    img.save(save_path)
    return save_path
