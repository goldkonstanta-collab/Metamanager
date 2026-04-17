import os
import re
import sys
import shutil
import tempfile
import importlib.util
from pathlib import Path
from typing import Any, Dict
from urllib.parse import quote

import requests
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware


ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))


TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")
CHECKO_API_KEY = os.getenv("CHECKO_API_KEY", "")

app = FastAPI(title="MetaManager Backend")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


def _telegram_configured() -> bool:
    return bool(TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID)


def _send_telegram_message(text: str) -> None:
    if not _telegram_configured():
        return
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    resp = requests.post(url, json={"chat_id": TELEGRAM_CHAT_ID, "text": text}, timeout=20)
    if not resp.ok:
        raise RuntimeError(f"Telegram sendMessage failed: {resp.status_code} {resp.text}")


def _send_telegram_document(file_path: str, caption: str = "") -> None:
    if not _telegram_configured():
        return
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    with open(file_path, "rb") as f:
        files = {"document": (os.path.basename(file_path), f)}
        data = {"chat_id": TELEGRAM_CHAT_ID}
        if caption:
            data["caption"] = caption
        resp = requests.post(url, data=data, files=files, timeout=90)
    if not resp.ok:
        raise RuntimeError(f"Telegram sendDocument failed: {resp.status_code} {resp.text}")


def _as_bool(value: Any, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    return str(value).strip().lower() in {"1", "true", "yes", "on", "да"}


def _checko_key_or_http() -> str:
    key = (CHECKO_API_KEY or "").strip()
    if not key:
        raise HTTPException(
            status_code=503,
            detail="CHECKO_API_KEY не задан на backend (нужен для запросов к api.checko.ru)",
        )
    return key


def _format_company_name(name: str, d: Dict[str, Any]) -> str:
    """Форматирует полное наименование организации (как в desktop main.py)."""
    if not name:
        return name
    LOWERCASE_WORDS = {
        "с",
        "о",
        "в",
        "и",
        "а",
        "на",
        "по",
        "из",
        "до",
        "за",
        "от",
        "об",
        "при",
        "для",
        "под",
        "над",
        "без",
        "про",
        "или",
        "но",
        "да",
        "не",
    }
    if name == name.upper():
        parts = re.split(r'(")', name)
        result = []
        in_quotes = False
        for part in parts:
            if part == '"':
                if not in_quotes:
                    result.append("«")
                else:
                    result.append("»")
                in_quotes = not in_quotes
            elif in_quotes:
                result.append(part.title())
            else:
                words = part.split()
                formatted = []
                for wi, w in enumerate(words):
                    w_lower = w.lower()
                    if wi == 0:
                        formatted.append(w.capitalize())
                    elif w_lower in LOWERCASE_WORDS:
                        formatted.append(w_lower)
                    elif len(w) <= 3 and w.isalpha() and w.isupper():
                        formatted.append(w)
                    else:
                        formatted.append(w.capitalize())
                result.append(" ".join(formatted))
        return "".join(result)
    name = re.sub(
        r"Общество С Ограниченной Ответственностью",
        "Общество с ограниченной ответственностью",
        name,
    )
    return name


def _format_short_name(name: str) -> str:
    if not name:
        return name
    name = name.replace('"', "«", 1).replace('"', "»", 1)
    parts = re.split(r"([«»])", name)
    result = []
    in_quotes = False
    for part in parts:
        if part in ("«", "»"):
            if part == "«":
                in_quotes = True
            else:
                in_quotes = False
            result.append(part)
        elif in_quotes and part == part.upper() and any(c.isalpha() for c in part):
            result.append(part.title())
        else:
            result.append(part)
    return "".join(result)


def _title_bank(s: str) -> str:
    if not s:
        return s
    s = re.sub(r'"([^"]+)"', lambda m: "«" + m.group(1) + "»", s)
    parts = re.split(r"(«[^»]+»)", s)
    result = []
    for part in parts:
        if part.startswith("«"):
            result.append(part)
        else:
            words = part.split()
            titled = []
            for wi, w in enumerate(words):
                if wi == 0 or w.lower() not in (
                    "в",
                    "на",
                    "по",
                    "для",
                    "и",
                    "или",
                    "с",
                    "о",
                    "об",
                    "от",
                    "до",
                    "из",
                    "за",
                    "при",
                    "под",
                    "над",
                    "без",
                    "через",
                ):
                    titled.append(w.capitalize())
                else:
                    titled.append(w.lower())
            result.append(" ".join(titled))
    return "".join(result)


def _parse_company_payload(d: Dict[str, Any]) -> Dict[str, Any]:
    full_name = d.get("НаимПолн", "") or ""
    full_name_fmt = _format_company_name(str(full_name), d)
    short_name = d.get("НаимСокр", "") or ""
    short_name_fmt = _format_short_name(str(short_name))

    addr = ""
    if "ЮрАдрес" in d and d["ЮрАдрес"]:
        addr = (d["ЮрАдрес"] or {}).get("АдресРФ", "") or ""

    contacts = d.get("Контакты", {}) or {}
    phones = contacts.get("Тел", []) or []
    emails = contacts.get("Емэйл", []) or []
    phone_str = ", ".join(phones[:2]) if phones else ""
    email_str = emails[0] if emails else ""

    director_title = ""
    director_name = ""
    basis = ""
    rukovod = d.get("Руковод", []) or []
    if rukovod:
        r = rukovod[0]
        director_name = str(r.get("ФИО", "") or "")
        dolzhn_raw = str(r.get("НаимДолжн", "") or "")
        director_title = dolzhn_raw.lower() if dolzhn_raw else ""

        okopf_code = ""
        if "ОКОПФ" in d and d["ОКОПФ"]:
            okopf_code = str((d["ОКОПФ"] or {}).get("Код", "") or "")
        if okopf_code.startswith("501"):
            basis = "свидетельства о государственной регистрации"
        else:
            basis = "Устава"

    return {
        "customerFullname": full_name_fmt,
        "customerShortname": short_name_fmt,
        "customerAddress": addr,
        "customerOgrn": str(d.get("ОГРН", "") or ""),
        "customerInn": str(d.get("ИНН", "") or ""),
        "customerKpp": str(d.get("КПП", "") or ""),
        "customerPhone": phone_str,
        "customerEmail": email_str,
        "customerDirectorTitle": director_title,
        "customerDirectorName": director_name,
        "customerBasis": basis,
    }


def _parse_bank_payload(d: Dict[str, Any]) -> Dict[str, Any]:
    bank_name_raw = (
        d.get("Наим", "")
        or d.get("НаимКред", "")
        or d.get("НаимПолн", "")
        or d.get("Наименование", "")
        or ""
    )
    bank_name_raw = str(bank_name_raw)
    bank_name = (
        _title_bank(bank_name_raw) if bank_name_raw == bank_name_raw.upper() else bank_name_raw
    )

    ks_raw = d.get("КорСчет", "") or d.get("КС", "") or ""
    ks_num = ""
    if isinstance(ks_raw, dict):
        ks_num = str(ks_raw.get("Номер", "") or "")
    elif isinstance(ks_raw, list) and ks_raw:
        first = ks_raw[0]
        ks_num = str(first.get("Номер", "")) if isinstance(first, dict) else str(first)
    else:
        ks_num = str(ks_raw) if ks_raw else ""

    return {"customerBank": bank_name, "customerKs": ks_num}


def _load_attr_from_file(file_path: Path, attr_name: str):
    """
    Загружает атрибут (класс/функцию) из python-файла по абсолютному пути.
    Используем для устойчивой работы на хостингах, где PYTHONPATH может отличаться.
    """
    if not file_path.exists():
        raise RuntimeError(f"File not found: {file_path}")

    module_name = f"_dynamic_{file_path.stem}"
    spec = importlib.util.spec_from_file_location(module_name, str(file_path))
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot create import spec for {file_path}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    if not hasattr(module, attr_name):
        raise RuntimeError(f"{attr_name} not found in {file_path.name}")
    return getattr(module, attr_name)


@app.get("/health")
def health() -> Dict[str, Any]:
    generator_path = ROOT_DIR / "generator.py"
    contract_path = ROOT_DIR / "contract_generator.py"
    templates_path = ROOT_DIR / "templates"
    return {
        "ok": True,
        "telegram_configured": _telegram_configured(),
        "checko_configured": bool((CHECKO_API_KEY or "").strip()),
        "paths": {
            "root_dir": str(ROOT_DIR),
            "generator_exists": generator_path.exists(),
            "contract_generator_exists": contract_path.exists(),
            "templates_exists": templates_path.exists(),
        },
    }


@app.get("/")
def root() -> Dict[str, Any]:
    return {"ok": True, "service": "metamanager-backend"}


@app.get("/lookup/company")
def lookup_company(inn: str) -> Dict[str, Any]:
    inn = (inn or "").strip()
    if not inn.isdigit() or len(inn) not in (10, 12):
        raise HTTPException(status_code=400, detail="ИНН должен содержать 10 или 12 цифр")

    key = _checko_key_or_http()
    url = f"https://api.checko.ru/v2/company?key={quote(key, safe='')}&inn={quote(inn, safe='')}"
    try:
        resp = requests.get(url, timeout=20, headers={"User-Agent": "MetaManagerBackend/1.0"})
        data = resp.json()
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Checko request failed: {e}")

    if not resp.ok:
        raise HTTPException(status_code=502, detail=f"Checko HTTP {resp.status_code}: {data}")

    if "data" not in data:
        raise HTTPException(status_code=404, detail="Данные не найдены. Проверьте ИНН.")

    d = data.get("data") or {}
    if not isinstance(d, dict) or not d:
        raise HTTPException(status_code=404, detail="Данные не найдены. Проверьте ИНН.")

    return {"ok": True, "inn": inn, "company": _parse_company_payload(d)}


@app.get("/lookup/bank")
def lookup_bank(bic: str) -> Dict[str, Any]:
    bik = (bic or "").strip()
    if not bik.isdigit() or len(bik) != 9:
        raise HTTPException(status_code=400, detail="БИК должен содержать 9 цифр")

    key = _checko_key_or_http()
    url = f"https://api.checko.ru/v2/bank?key={quote(key, safe='')}&bic={quote(bik, safe='')}"
    try:
        resp = requests.get(url, timeout=20, headers={"User-Agent": "MetaManagerBackend/1.0"})
        data = resp.json()
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Checko request failed: {e}")

    if not resp.ok:
        raise HTTPException(status_code=502, detail=f"Checko HTTP {resp.status_code}: {data}")

    if "data" not in data or not data["data"]:
        raise HTTPException(status_code=404, detail="Банк не найден")

    d = data.get("data") or {}
    if not isinstance(d, dict):
        raise HTTPException(status_code=404, detail="Банк не найден")

    return {"ok": True, "bic": bik, "bank": _parse_bank_payload(d)}


@app.post("/generate/kp")
def generate_kp(payload: Dict[str, Any]) -> Dict[str, Any]:
    if not payload.get("kpName") or not payload.get("kpTitle"):
        raise HTTPException(status_code=400, detail="kpName и kpTitle обязательны")

    temp_dir = tempfile.mkdtemp(prefix="metamanager_kp_")
    try:
        try:
            KPGenerator = _load_attr_from_file(ROOT_DIR / "generator.py", "KPGenerator")
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Import KPGenerator failed: {e}")

        smr_type = payload.get("smrType") or "без смр"
        kp_data = {
            "kp_name": str(payload.get("kpName", "")).strip(),
            "kp_title": str(payload.get("kpTitle", "")).strip(),
            "branch": payload.get("branch") or "хоз.пит",
            "volume": payload.get("volume") or "до 100",
            "smr_type": smr_type,
            "include_wells": _as_bool(payload.get("includeWells"), default=True),
            "include_pump": _as_bool(payload.get("includePump"), default=True),
            "include_bmz": _as_bool(payload.get("includeBmz"), default=True),
            "save_dir": temp_dir,
            "include_pir": _as_bool(payload.get("includePir"), default=False),
            "pir_count": payload.get("pirCount") or 1,
            "pir_price": payload.get("pirPrice") or 0,
        }

        if smr_type == "с смр":
            kp_data.update(
                {
                    "wells_count": payload.get("wellsCountSmr") or 1,
                    "wells_design": str(payload.get("wellsDesign", "")).strip(),
                    "wells_depth": str(payload.get("wellsDepth", "")).strip(),
                    "wells_price": str(payload.get("wellsPrice", "")).strip(),
                    "pump_price": str(payload.get("pumpPrice", "")).strip(),
                    "bmz_size": str(payload.get("bmzSize", "")).strip(),
                    "bmz_price": str(payload.get("bmzPrice", "")).strip(),
                }
            )
        else:
            kp_data["wells_count"] = payload.get("wellsCount") or 1

        generator = KPGenerator()
        docx_path, pdf_path = generator.create_kp(kp_data)

        _send_telegram_message(f"Новое КП: {kp_data['kp_name']}")
        _send_telegram_document(docx_path, caption="КП (Word)")
        if pdf_path and os.path.exists(pdf_path):
            _send_telegram_document(pdf_path, caption="КП (PDF)")

        return {
            "ok": True,
            "type": "kp",
            "telegram": {
                "configured": _telegram_configured(),
                "sent": _telegram_configured(),
                "targetChatId": TELEGRAM_CHAT_ID if TELEGRAM_CHAT_ID else None,
            },
            "files": {
                "docx": os.path.basename(docx_path),
                "pdf": os.path.basename(pdf_path) if pdf_path and os.path.exists(pdf_path) else None,
            },
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


@app.post("/generate/contract")
def generate_contract(payload: Dict[str, Any]) -> Dict[str, Any]:
    if not payload.get("contractNumber") or not payload.get("customerShortname"):
        raise HTTPException(status_code=400, detail="contractNumber и customerShortname обязательны")

    if _as_bool(payload.get("includeWorkAddress")) and not payload.get("workAddress"):
        raise HTTPException(status_code=400, detail="Укажите адрес проведения работ")

    temp_dir = tempfile.mkdtemp(prefix="metamanager_contract_")
    try:
        try:
            ContractGenerator = _load_attr_from_file(
                ROOT_DIR / "contract_generator.py",
                "ContractGenerator"
            )
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Import ContractGenerator failed: {e}")

        contract_data = {
            "contract_number": str(payload.get("contractNumber", "")).strip(),
            "save_dir": temp_dir,
            "customer_fullname": str(payload.get("customerFullname", "")).strip(),
            "customer_shortname": str(payload.get("customerShortname", "")).strip(),
            "customer_address": str(payload.get("customerAddress", "")).strip(),
            "customer_ogrn": str(payload.get("customerOgrn", "")).strip(),
            "customer_inn": str(payload.get("customerInn", "")).strip(),
            "customer_kpp": str(payload.get("customerKpp", "")).strip(),
            "customer_bank": str(payload.get("customerBank", "")).strip(),
            "customer_bik": str(payload.get("customerBik", "")).strip(),
            "customer_rs": str(payload.get("customerRs", "")).strip(),
            "customer_ks": str(payload.get("customerKs", "")).strip(),
            "customer_phone": str(payload.get("customerPhone", "")).strip(),
            "customer_email": str(payload.get("customerEmail", "")).strip(),
            "customer_director_title": str(payload.get("customerDirectorTitle", "")).strip(),
            "customer_director_name": str(payload.get("customerDirectorName", "")).strip(),
            "customer_basis": str(payload.get("customerBasis", "")).strip(),
            "advance_percent": str(payload.get("advancePercent", "30")).strip() or "30",
            "kp_file": "",  # В MVP веб-версии КП-файл пока не загружается.
            "include_work_address": _as_bool(payload.get("includeWorkAddress"), default=False),
            "work_address": str(payload.get("workAddress", "")).strip(),
        }

        generator = ContractGenerator()
        docx_path = generator.create_contract(contract_data)

        _send_telegram_message(f"Новый договор: 0ЦЦБ-{contract_data['contract_number']}")
        _send_telegram_document(docx_path, caption="Договор (Word)")

        return {
            "ok": True,
            "type": "contract",
            "telegram": {
                "configured": _telegram_configured(),
                "sent": _telegram_configured(),
                "targetChatId": TELEGRAM_CHAT_ID if TELEGRAM_CHAT_ID else None,
            },
            "files": {
                "docx": os.path.basename(docx_path),
            },
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
