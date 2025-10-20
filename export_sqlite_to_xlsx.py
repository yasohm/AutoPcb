import sqlite3
from pathlib import Path
import pandas as pd

DB_PATH = Path("data.db")
OUTPUT_DIR = Path("input")
TABLE_TO_FILE = {
    "abc": OUTPUT_DIR / "ABC.xlsx",
    "fb": OUTPUT_DIR / "FB.xlsx",
    "pcb": OUTPUT_DIR / "PCB.xlsx",
}


def export_table(conn: sqlite3.Connection, table: str, out_path: Path) -> None:
    df = pd.read_sql_query(f'SELECT * FROM "{table}"', conn)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=table, index=False)
    print(f"Exported table {table} -> {out_path}")


def main() -> None:
    if not DB_PATH.exists():
        raise FileNotFoundError(f"Database not found: {DB_PATH}")
    with sqlite3.connect(DB_PATH.as_posix()) as conn:
        for table, out_path in TABLE_TO_FILE.items():
            cursor = conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
                (table,),
            )
            if cursor.fetchone() is None:
                print(f"Skip: table {table} not found in {DB_PATH}")
                continue
            export_table(conn, table, out_path)
    print("Done exporting all available tables.")


if __name__ == "__main__":
    main()
