"""مدیریت مخاطبین و لیست ایمیل‌ها"""
import json
from pathlib import Path

BASE_DIR = Path(__file__).parent
CONTACTS_FILE = BASE_DIR / "contacts.json"
EMAILS_FILE = BASE_DIR / "recipients.json"


def load_contacts():
    if CONTACTS_FILE.exists():
        try:
            data = json.loads(CONTACTS_FILE.read_text(encoding="utf-8"))
            return data.get("contacts", [])
        except Exception:
            pass
    return []


def save_contacts(contacts):
    data = {"contacts": list(contacts)}
    CONTACTS_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def load_emails():
    if EMAILS_FILE.exists():
        try:
            data = json.loads(EMAILS_FILE.read_text(encoding="utf-8"))
            return {
                "recipients": data.get("recipients", []),
                "subject": data.get("subject", "گزارش هفتگی — کالای راکد افق کوروش"),
            }
        except Exception:
            pass
    return {
        "recipients": [],
        "subject": "گزارش هفتگی — کالای راکد افق کوروش",
    }


def save_emails(recipients, subject):
    data = {"recipients": list(recipients), "subject": subject}
    EMAILS_FILE.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def parse_contacts_file(file_bytes, filename="contacts.xlsx"):
    import io
    import pandas as pd

    buf = io.BytesIO(file_bytes)
    name = (filename or "").lower()
    if name.endswith(".csv"):
        df = pd.read_csv(buf)
    else:
        df = pd.read_excel(buf)

    df.columns = df.columns.astype(str).str.strip()

    def find_col(keys):
        for c in df.columns:
            cl = c.lower()
            if any(k in cl or k in c for k in keys):
                return c
        return None

    email_col = find_col(["email", "ایمیل", "mail"])
    name_col = find_col(["نام", "name"])
    type_col = find_col(["نوع", "type", "role", "نقش"])
    chat_col = find_col(["chat_id", "chat", "bale", "بله"])

    if email_col is None and chat_col is None:
        raise ValueError("ستون ایمیل یا chat_id در فایل پیدا نشد.")

    contacts = []
    for _, row in df.iterrows():
        email = str(row.get(email_col, "")).strip() if email_col else ""
        chat_id = None
        if chat_col:
            raw_chat = row.get(chat_col)
            s = str(raw_chat).strip()
            if s not in ("", "nan", "None"):
                try:
                    chat_id = int(float(s))
                except (ValueError, TypeError):
                    chat_id = None

        if not email and not chat_id:
            continue
        if email and "@" not in email:
            email = ""

        name_v = str(row.get(name_col, "")).strip() if name_col else ""
        type_v = str(row.get(type_col, "")).strip().lower() if type_col else "regular"

        if "supervisor" in type_v or "سوپروایزر" in type_v or "سرپرست" in type_v:
            t = "supervisor"
        elif "manager" in type_v or "مدیر" in type_v:
            t = "manager"
        else:
            t = "regular"

        entry = {"name": name_v, "email": email, "type": t}
        if chat_id:
            entry["bale_chat_id"] = chat_id
        contacts.append(entry)

    return contacts


def update_contact_chat_id(email_or_name, chat_id):
    contacts = load_contacts()
    updated = False
    for c in contacts:
        if c.get("email") == email_or_name or c.get("name") == email_or_name:
            c["bale_chat_id"] = int(chat_id)
            updated = True
            break
    if updated:
        save_contacts(contacts)
    return updated