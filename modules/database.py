import sqlite3
import logging
import os

# Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

DB_FILE = "subjects.db"

def get_db_connection():
    """Establishes a connection to the SQLite database."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """Initializes the database table if it doesn't exist."""
    conn = get_db_connection()
    try:
        conn.execute('''
            CREATE TABLE IF NOT EXISTS subject_mappings (
                raw_name TEXT PRIMARY KEY,
                clean_name TEXT NOT NULL
            )
        ''')
        # We ignore subject_codes table as per latest request
        conn.commit()
        logger.info("Database initialized successfully.")
    except Exception as e:
        logger.error(f"Error initializing database: {e}")
    finally:
        conn.close()

def get_clean_name(raw_name: str) -> str:
    """
    Retrieves the clean name for a given raw name.
    If not found, inserts the raw_name as the clean_name (auto-discovery) and returns it.
    """
    conn = get_db_connection()
    clean_name = raw_name # Default fallback
    try:
        cursor = conn.execute('SELECT clean_name FROM subject_mappings WHERE raw_name = ?', (raw_name,))
        row = cursor.fetchone()
        
        if row:
            clean_name = row['clean_name']
        else:
            # Auto-insert new raw name with itself as the default clean name
            conn.execute('INSERT INTO subject_mappings (raw_name, clean_name) VALUES (?, ?)', (raw_name, raw_name))
            conn.commit()
            logger.info(f"New subject discovered and added: {raw_name}")
            
    except Exception as e:
        logger.error(f"Error in get_clean_name: {e}")
    finally:
        conn.close()
        
    return clean_name

def update_mapping(raw_name: str, new_clean_name: str):
    """Updates the clean name for a specific raw name."""
    conn = get_db_connection()
    try:
        conn.execute('UPDATE subject_mappings SET clean_name = ? WHERE raw_name = ?', (new_clean_name, raw_name))
        conn.commit()
        logger.info(f"Updated mapping: {raw_name} -> {new_clean_name}")
    except Exception as e:
        logger.error(f"Error updating mapping: {e}")
    finally:
        conn.close()

def get_all_mappings():
    """Returns all mappings as a list of dictionaries."""
    conn = get_db_connection()
    mappings = []
    try:
        cursor = conn.execute('SELECT * FROM subject_mappings ORDER BY raw_name')
        mappings = [dict(row) for row in cursor.fetchall()]
    except Exception as e:
        logger.error(f"Error fetching mappings: {e}")
    finally:
        conn.close()
    return mappings

# No longer using save_subject_code or get_subject_codes_matrix
