import json
import sqlite3
from pathlib import Path
from typing import Any, Dict, List, Optional


def connect(db_path: Path) -> sqlite3.Connection:
    con = sqlite3.connect(str(db_path))
    con.row_factory = sqlite3.Row
    return con


def init_db(db_path: Path, seed_bot_db: Optional[Path] = None) -> None:
    """Create tables if missing. Optionally import clients from bot DB."""
    db_path.parent.mkdir(parents=True, exist_ok=True)
    con = connect(db_path)
    cur = con.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS clients (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fio TEXT NOT NULL,
            passport TEXT DEFAULT '',
            organ TEXT DEFAULT '',
            vydan TEXT DEFAULT '',
            phone TEXT DEFAULT '',
            address TEXT DEFAULT '',
            extra_json TEXT DEFAULT '{}',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
        """
    )

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS contracts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            client_id INTEGER NOT NULL,
            template_id TEXT NOT NULL,
            company TEXT DEFAULT '',
            template_name TEXT DEFAULT '',
            contract_no TEXT DEFAULT '',
            contract_date TEXT DEFAULT '',
            car_model TEXT DEFAULT '',
            dkp_amount TEXT DEFAULT '',
            duty_amount TEXT DEFAULT '',
            mapping_json TEXT DEFAULT '{}',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (client_id) REFERENCES clients(id)
        )
        """
    )

    con.commit()

    # optional seed from telegram bot DB (clients.db)
    if seed_bot_db and seed_bot_db.exists():
        try:
            bot_con = sqlite3.connect(str(seed_bot_db))
            bot_con.row_factory = sqlite3.Row
            bot_cur = bot_con.cursor()
            bot_cur.execute(
                "SELECT fio, contract_no, contract_date, car_model FROM clients ORDER BY id ASC"
            )
            rows = bot_cur.fetchall()
            bot_con.close()

            # If current db already has clients, do not import.
            cur.execute("SELECT COUNT(1) AS c FROM clients")
            if int(cur.fetchone()[0]) == 0 and rows:
                for r in rows:
                    cur.execute(
                        "INSERT INTO clients (fio) VALUES (?)",
                        (r["fio"],),
                    )
                    cid = cur.lastrowid
                    cur.execute(
                        """
                        INSERT INTO contracts (
                          client_id, template_id, company, template_name,
                          contract_no, contract_date, car_model
                        ) VALUES (?, '', '', 'Imported', ?, ?, ?)
                        """,
                        (cid, r["contract_no"], r["contract_date"], r["car_model"]),
                    )
                con.commit()
        except Exception:
            # silent seed failure
            pass

    con.close()


def list_clients(db_path: Path, limit: int = 200) -> List[Dict[str, Any]]:
    con = connect(db_path)
    cur = con.cursor()
    cur.execute(
        "SELECT id, fio, passport, organ, vydan, phone, address, extra_json, created_at FROM clients ORDER BY id DESC LIMIT ?",
        (limit,),
    )
    out = []
    for r in cur.fetchall():
        d = dict(r)
        try:
            d["extra"] = json.loads(d.get("extra_json") or "{}")
        except Exception:
            d["extra"] = {}
        out.append(d)
    con.close()
    return out


def get_client(db_path: Path, client_id: int) -> Optional[Dict[str, Any]]:
    con = connect(db_path)
    cur = con.cursor()
    cur.execute(
        "SELECT id, fio, passport, organ, vydan, phone, address, extra_json, created_at FROM clients WHERE id=?",
        (client_id,),
    )
    r = cur.fetchone()
    con.close()
    if not r:
        return None
    d = dict(r)
    try:
        d["extra"] = json.loads(d.get("extra_json") or "{}")
    except Exception:
        d["extra"] = {}
    return d


def upsert_client(
    db_path: Path,
    *,
    client_id: Optional[int],
    fio: str,
    passport: str = "",
    organ: str = "",
    vydan: str = "",
    phone: str = "",
    address: str = "",
    extra: Optional[Dict[str, Any]] = None,
) -> int:
    con = connect(db_path)
    cur = con.cursor()
    extra_json = json.dumps(extra or {}, ensure_ascii=False)

    if client_id:
        cur.execute(
            """
            UPDATE clients
            SET fio=?, passport=?, organ=?, vydan=?, phone=?, address=?, extra_json=?
            WHERE id=?
            """,
            (fio, passport, organ, vydan, phone, address, extra_json, client_id),
        )
        con.commit()
        con.close()
        return int(client_id)

    cur.execute(
        """
        INSERT INTO clients (fio, passport, organ, vydan, phone, address, extra_json)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (fio, passport, organ, vydan, phone, address, extra_json),
    )
    con.commit()
    cid = int(cur.lastrowid)
    con.close()
    return cid


def add_contract(
    db_path: Path,
    *,
    client_id: int,
    template_id: str,
    company: str,
    template_name: str,
    contract_no: str,
    contract_date: str,
    car_model: str,
    dkp_amount: str,
    duty_amount: str,
    mapping: Dict[str, str],
) -> int:
    con = connect(db_path)
    cur = con.cursor()
    cur.execute(
        """
        INSERT INTO contracts (
          client_id, template_id, company, template_name,
          contract_no, contract_date, car_model,
          dkp_amount, duty_amount, mapping_json
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            client_id,
            template_id,
            company,
            template_name,
            contract_no,
            contract_date,
            car_model,
            dkp_amount,
            duty_amount,
            json.dumps(mapping, ensure_ascii=False),
        ),
    )
    con.commit()
    cid = int(cur.lastrowid)
    con.close()
    return cid


def list_contracts_for_client(db_path: Path, client_id: int, limit: int = 50) -> List[Dict[str, Any]]:
    con = connect(db_path)
    cur = con.cursor()
    cur.execute(
        """
        SELECT id, template_id, company, template_name, contract_no, contract_date,
               car_model, dkp_amount, duty_amount, mapping_json, created_at
        FROM contracts
        WHERE client_id=?
        ORDER BY id DESC
        LIMIT ?
        """,
        (client_id, limit),
    )
    out = []
    for r in cur.fetchall():
        d = dict(r)
        try:
            d["mapping"] = json.loads(d.get("mapping_json") or "{}")
        except Exception:
            d["mapping"] = {}
        out.append(d)
    con.close()
    return out


def get_contract(db_path: Path, contract_id: int) -> Optional[Dict[str, Any]]:
    con = connect(db_path)
    cur = con.cursor()
    cur.execute(
        """
        SELECT id, client_id, template_id, company, template_name, contract_no, contract_date,
               car_model, dkp_amount, duty_amount, mapping_json, created_at
        FROM contracts WHERE id=?
        """,
        (contract_id,),
    )
    r = cur.fetchone()
    con.close()
    if not r:
        return None
    d = dict(r)
    try:
        d["mapping"] = json.loads(d.get("mapping_json") or "{}")
    except Exception:
        d["mapping"] = {}
    return d
