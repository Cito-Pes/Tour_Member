# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import re
import sqlite3
import sys
from pathlib import Path
from typing import Any

import pyodbc
import requests


SOURCE_DIR = Path(__file__).resolve().parent
BUNDLE_DIR = Path(getattr(sys, "_MEIPASS", SOURCE_DIR))
BASE_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else SOURCE_DIR
DB_DIR = BASE_DIR / "DB"
DB_FILE = "Config_DB.db"
CONFIG_NAME = "HD_MSSQL"
CONFIG_TXT = BASE_DIR / "config.txt" if (BASE_DIR / "config.txt").exists() else BUNDLE_DIR / "config.txt"


def load_config_txt(filepath: str | os.PathLike[str] = CONFIG_TXT) -> dict[str, str]:
    config = {"Not_Charge_IDP": "", "Not_PlaceofDuty": "", "google_doc_key": ""}
    path = Path(filepath)
    if not path.exists():
        return config
    with path.open("r", encoding="utf-8") as file:
        for line in file:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            if key.strip() in config:
                config[key.strip()] = value.strip()
    return config


def download_db(
    gdrive_url: str | None = None,
    db_dir: str | os.PathLike[str] = DB_DIR,
    db_file: str = DB_FILE,
) -> tuple[bool, str]:
    db_path = Path(db_dir) / db_file
    db_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        config = load_config_txt()
        doc_key = re.search(r"/d/([a-zA-Z0-9_-]+)", gdrive_url or "")
        if not doc_key and config.get("google_doc_key"):
            doc_key = re.search(r"([a-zA-Z0-9_-]{20,})", config["google_doc_key"])
        if not doc_key:
            raise ValueError("config.txt에 google_doc_key가 없습니다.")
        file_id = doc_key.group(1)
        response = requests.get(
            f"https://drive.google.com/uc?export=download&id={file_id}",
            timeout=60,
        )
        response.raise_for_status()
        db_path.write_bytes(response.content)
        return True, str(db_path)
    except Exception as exc:
        return False, str(exc)


def ensure_config_db(
    db_dir: str | os.PathLike[str] = DB_DIR,
    db_file: str = DB_FILE,
) -> Path:
    db_path = Path(db_dir) / db_file
    if db_path.exists():
        return db_path
    ok, result = download_db(db_dir=db_dir, db_file=db_file)
    if not ok:
        raise RuntimeError(f"설정 DB 다운로드 실패: {result}")
    return Path(result)


def load_db_config(
    config_name: str = CONFIG_NAME,
    db_dir: str | os.PathLike[str] = DB_DIR,
    db_file: str = DB_FILE,
    auto_download: bool = True,
) -> dict[str, Any]:
    db_path = ensure_config_db(db_dir, db_file) if auto_download else Path(db_dir) / db_file
    if not db_path.exists():
        raise FileNotFoundError(f"설정 DB 파일이 없습니다: {db_path}")
    with sqlite3.connect(db_path) as conn:
        row = conn.execute(
            "SELECT DB_Type, Host, Port, DB_Name, DB_ID, DB_PW FROM DBCON WHERE Name = ?",
            (config_name,),
        ).fetchone()
    if not row:
        raise LookupError(f"DBCON Name='{config_name}' 레코드가 없습니다.")
    return {
        "DB_Type": row[0], "Host": row[1], "Port": row[2],
        "DB_Name": row[3], "DB_ID": row[4], "DB_PW": row[5],
    }


def build_mssql_connection_string(config: dict[str, Any]) -> str:
    return (
        f"DRIVER={{{config['DB_Type']}}};"
        f"SERVER={config['Host']},{config['Port']};"
        f"DATABASE={config['DB_Name']};"
        f"UID={config['DB_ID']};PWD={config['DB_PW']}"
    )


def get_mssql_connection(config: dict[str, Any], timeout: int = 30) -> pyodbc.Connection:
    return pyodbc.connect(build_mssql_connection_string(config), timeout=timeout)


def connect(config_name: str = CONFIG_NAME, timeout: int = 30) -> pyodbc.Connection:
    return get_mssql_connection(load_db_config(config_name=config_name), timeout=timeout)
