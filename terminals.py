import sqlite3
from typing import List


DB_ADDRESS = r'C:\Max\slip_db\slipDB_single.db'


def get(object_code: str) -> List[str]:
    conn = sqlite3.connect(DB_ADDRESS)
    cur = conn.execute(f"SELECT DISTINCT pos_id FROM slips WHERE object_code = ?", (object_code,))
    rows = cur.fetchall()
    conn.close()
    return [item[0] for item in rows]
