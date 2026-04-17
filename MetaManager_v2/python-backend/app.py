import os
import sys
import shutil
import tempfile
import importlib.util
from pathlib import Path
from typing import Any, Dict

import requests
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware


ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))


TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")

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
