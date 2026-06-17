import re
from .models import AgreementData, FieldState
from .defaults import CRITICAL_FIELDS

def get_nested_attr(obj, path):
    for p in path.split("."):
        if not obj: return None
        obj = getattr(obj, p, None) if not isinstance(obj, dict) else obj.get(p)
    return obj

def validate_agreement(data, verified):
    missing = [f for f in CRITICAL_FIELDS if not get_nested_attr(data, f)]
    warns = []
    if data.client.trade_name == "EchowPanda" and not data.client.legal_name:
        warns.append("EchowPanda trade name detected. Legal name required.")
    if not data.contractor.vat_number: warns.append("VAT missing.")
    if missing: warns.append("Placeholders remain.")
    return missing, warns

def get_placeholder(path): return f"[REQUIRED: {path.split('.')[-1].upper()}]"
def get_demo_value(path): return f"SAMPLE-{path.split('.')[-1].upper()}"

def protect_hallucinations(text, data):
    if data.mode != "production": return text
    # Simplified IBAN protection
    for iban in re.findall(r'[A-Z]{2}\d{2}[A-Z0-9]{11,30}', text):
        if iban != data.contractor.iban: text = text.replace(iban, "[REQUIRED: IBAN]")
    return text
