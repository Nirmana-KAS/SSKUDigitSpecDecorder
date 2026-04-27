import sqlite3
import os
import sys


def get_db_path():
    if getattr(sys, 'frozen', False):
        app_dir = os.path.dirname(sys.executable)
    else:
        app_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(app_dir, 'data')
    os.makedirs(data_dir, exist_ok=True)
    return os.path.join(data_dir, 'libraries.db')


TABLES = {
    'stone_names': ('code', 'name'),
    'polishing_types': ('code', 'name'),
    'shapes': ('code', 'name'),
    'colours': ('code', 'name'),
    'origins': ('code', 'name'),
}


class Database:
    def __init__(self):
        self.conn = sqlite3.connect(get_db_path())
        self._create_tables()

    def _create_tables(self):
        c = self.conn.cursor()
        for table in TABLES:
            c.execute(f'''CREATE TABLE IF NOT EXISTS {table} (
                code TEXT PRIMARY KEY COLLATE NOCASE,
                name TEXT NOT NULL
            )''')
        self.conn.commit()

    def add_entries(self, table, entries):
        c = self.conn.cursor()
        added, skipped = 0, 0
        for code, name in entries:
            code_str = str(code).strip()
            name_str = str(name).strip()
            if not code_str or not name_str:
                continue
            try:
                c.execute(f'INSERT INTO {table} (code, name) VALUES (?, ?)',
                          (code_str, name_str))
                added += 1
            except sqlite3.IntegrityError:
                skipped += 1
        self.conn.commit()
        return added, skipped

    def get_count(self, table):
        c = self.conn.cursor()
        c.execute(f'SELECT COUNT(*) FROM {table}')
        return c.fetchone()[0]

    def get_lookup(self, table):
        c = self.conn.cursor()
        c.execute(f'SELECT code, name FROM {table}')
        return {row[0].upper(): row[1] for row in c.fetchall()}

    def get_all_entries(self, table):
        c = self.conn.cursor()
        c.execute(f'SELECT code, name FROM {table} ORDER BY code')
        return c.fetchall()

    def update_entry(self, table, code, new_name):
        c = self.conn.cursor()
        c.execute(f'UPDATE {table} SET name = ? WHERE code = ?',
                  (str(new_name).strip(), str(code).strip()))
        self.conn.commit()

    def delete_entry(self, table, code):
        c = self.conn.cursor()
        c.execute(f'DELETE FROM {table} WHERE code = ?', (str(code).strip(),))
        self.conn.commit()

    def delete_all(self, table):
        c = self.conn.cursor()
        c.execute(f'DELETE FROM {table}')
        self.conn.commit()

    def close(self):
        self.conn.close()
