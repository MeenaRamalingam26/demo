import sqlite3, os
from werkzeug.security import generate_password_hash

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, "users.db")

conn = sqlite3.connect(DB_PATH)
c = conn.cursor()

# Ensure table exists
c.execute('''CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL)''')

users = [
    ('alice', generate_password_hash('password123')),
    ('bob', generate_password_hash('securepass')),
    ('charlie', generate_password_hash('mypassword')),
]

# Insert users, ignore if they already exist
c.executemany(
    "INSERT OR IGNORE INTO users (username, password) VALUES (?, ?)",
    users
)

conn.commit()
conn.close()

print("Sample users added successfully!")