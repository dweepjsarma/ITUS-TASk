"""
ebitda_margins_data_udf.py

Excel UDFs to query a local SQLite database of financial data.

Supports:
- get_quarterly_data(accord_code, field, date)
- get_series(accord_code, field, start_date, end_date)
- get_quarterly_matrix(date, field)
- get_all_ebitda_margins(accord_code, field)

Integration:
- Primary integration via xlwings (recommended)
- Can be used as standalone CLI for testing

Configuration:
- Reads config.ini in the same directory
"""

import sqlite3
import os
import configparser
from functools import lru_cache
from datetime import datetime
import time
import logging
from logging.handlers import RotatingFileHandler
import xlwings as xw


# Optional Excel integration
try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except Exception:
    XLWINGS_AVAILABLE = False

# -------------------------
# Configuration management
# -------------------------

def load_config(config_path=None):
    """Load configuration from config.ini file"""
    config = configparser.ConfigParser(interpolation=None)

    if config_path and os.path.exists(config_path):
        config.read(config_path)
    else:
        # Try script directory
        base_dir = os.path.dirname(os.path.abspath(__file__))
        local_path = os.path.join(base_dir, "config.ini")

        if os.path.exists(local_path):
            config.read(local_path)
        else:
            # Fallback defaults
            config['database'] = {
                'path': os.path.join(base_dir, 'ebitda_margins.db'),
                'table': 'ebitda_margins',
                'date_format': '%Y-%m-%d'
            }

    return config

CONFIG = load_config()
DB_FILE = CONFIG['database'].get('path', 'ebitda_margins.db')
TABLE_NAME = CONFIG['database'].get('table', 'ebitda_margins')
DATE_FORMAT = CONFIG['database'].get('date_format', '%Y-%m-%d')

# -------------------------
# Logging setup
# -------------------------

LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "query_log.txt")
logger = logging.getLogger("ebitda_udf_logger")
logger.setLevel(logging.INFO)

if not logger.handlers:
    handler = RotatingFileHandler(LOG_FILE, maxBytes=1_000_000, backupCount=3, encoding='utf-8')
    formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)

def log_call(func_name, params, start_time, success=True, error_msg=None):
    """Log function calls with timing information"""
    elapsed_ms = (time.perf_counter() - start_time) * 1000
    msg = f"{func_name} | params={params} | time_ms={elapsed_ms:.2f} | success={success}"
    if error_msg:
        msg += f" | error={error_msg}"
    logger.info(msg)

# -------------------------
# Database helper
# -------------------------

def get_connection():
    """Create and return a database connection"""
    if not os.path.exists(DB_FILE):
        raise FileNotFoundError(f"Database file not found at {DB_FILE}")

    conn = sqlite3.connect(DB_FILE, detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
    conn.row_factory = sqlite3.Row
    return conn

def validate_date(date_text):
    """Validate date format"""
    try:
        return datetime.strptime(date_text, DATE_FORMAT).date()
    except Exception as e:
        raise ValueError(f"Date '{date_text}' does not match format {DATE_FORMAT}") from e

# -------------------------
# Core query functions
# -------------------------

@lru_cache(maxsize=4096)
def _get_quarterly_data_cached(accord_code, field, date_text):
    """Cached version of get_quarterly_data"""
    start = time.perf_counter()
    try:
        validate_date(date_text)
        sql = f"""SELECT {field} FROM {TABLE_NAME}
                  WHERE accord_code = ? AND date = ? LIMIT 1"""

        with get_connection() as conn:
            cur = conn.execute(sql, (accord_code, date_text))
            row = cur.fetchone()
            value = row[field] if row and field in row.keys() else None

        log_call("get_quarterly_data", 
                {"accord_code": accord_code, "field": field, "date": date_text}, 
                start, success=True)
        return value

    except Exception as e:
        log_call("get_quarterly_data", 
                {"accord_code": accord_code, "field": field, "date": date_text}, 
                start, success=False, error_msg=str(e))
        raise

def get_quarterly_data(accord_code, field, date_text):
    """
    Return a scalar value for accord_code on the exact date.

    Args:
        accord_code: Company identifier
        field: Database column name
        date_text: Date in format YYYY-MM-DD

    Returns:
        Single value or None
    """
    return _get_quarterly_data_cached(int(accord_code), field, date_text)

def get_series(accord_code, field, start_date, end_date):
    """
    Returns list of (date, value) between start_date and end_date inclusive.

    Args:
        accord_code: Company identifier
        field: Database column name
        start_date: Start date in format YYYY-MM-DD
        end_date: End date in format YYYY-MM-DD

    Returns:
        List of tuples (date, value)
    """
    start = time.perf_counter()
    try:
        validate_date(start_date)
        validate_date(end_date)

        sql = f"""SELECT date, {field} FROM {TABLE_NAME}
                  WHERE accord_code = ? AND date BETWEEN ? AND ?
                  ORDER BY date"""

        with get_connection() as conn:
            cur = conn.execute(sql, (int(accord_code), start_date, end_date))
            rows = cur.fetchall()
            result = [(r["date"], r[field]) for r in rows]

        log_call("get_series", 
                {"accord_code": accord_code, "field": field, 
                 "start_date": start_date, "end_date": end_date}, 
                start, success=True)
        return result

    except Exception as e:
        log_call("get_series", 
                {"accord_code": accord_code, "field": field, 
                 "start_date": start_date, "end_date": end_date}, 
                start, success=False, error_msg=str(e))
        raise

def get_quarterly_matrix(date_text, field):
    """
    Returns list of rows: (accord_code, company_name, sector, mcap_category, field_value)
    for the specified date.

    Args:
        date_text: Date in format YYYY-MM-DD
        field: Database column name

    Returns:
        List of tuples with company data
    """
    start = time.perf_counter()
    try:
        validate_date(date_text)

        cols = "accord_code, company_name, sector, mcap_category"
        sql = f"""SELECT {cols}, {field} FROM {TABLE_NAME}
                  WHERE date = ?
                  ORDER BY accord_code"""

        with get_connection() as conn:
            cur = conn.execute(sql, (date_text,))
            rows = cur.fetchall()
            result = [(r["accord_code"], r["company_name"], r["sector"], 
                      r["mcap_category"], r[field]) for r in rows]

        log_call("get_quarterly_matrix", 
                {"date": date_text, "field": field}, 
                start, success=True)
        return result

    except Exception as e:
        log_call("get_quarterly_matrix", 
                {"date": date_text, "field": field}, 
                start, success=False, error_msg=str(e))
        raise

def get_all_ebitda_margins(accord_code, field):
    """
    Returns list of (date, value) for all dates available for a company.

    Args:
        accord_code: Company identifier
        field: Database column name

    Returns:
        List of tuples (date, value)
    """
    start = time.perf_counter()
    try:
        sql = f"""SELECT date, {field} FROM {TABLE_NAME}
                  WHERE accord_code = ?
                  ORDER BY date"""

        with get_connection() as conn:
            cur = conn.execute(sql, (int(accord_code),))
            rows = cur.fetchall()
            result = [(r["date"], r[field]) for r in rows]

        log_call("get_all_ebitda_margins", 
                {"accord_code": accord_code, "field": field}, 
                start, success=True)
        return result

    except Exception as e:
        log_call("get_all_ebitda_margins", 
                {"accord_code": accord_code, "field": field}, 
                start, success=False, error_msg=str(e))
        raise

# -------------------------
# Excel (xlwings) decorators
# -------------------------

if XLWINGS_AVAILABLE:

    @xw.func
    @xw.arg('accord_code', numbers=int)
    @xw.arg('field', doc='Database field name')
    @xw.arg('date_text', doc='Date in YYYY-MM-DD format')
    def xw_get_quarterly_data(accord_code, field, date_text):
        """
        Excel UDF: Get quarterly data for a specific company and date.

        Usage in Excel: =xw_get_quarterly_data(100186, "quarterly_ebitda_margins", "2025-06-30")
        """
        try:
            return get_quarterly_data(accord_code, field, date_text)
        except Exception as e:
            return f"#ERROR: {e}"

    @xw.func
    @xw.arg('accord_code', numbers=int)
    @xw.arg('field', doc='Database field name')
    @xw.arg('start_date', doc='Start date in YYYY-MM-DD format')
    @xw.arg('end_date', doc='End date in YYYY-MM-DD format')
    def xw_get_series(accord_code, field, start_date, end_date):
        """
        Excel UDF: Get time series data for a specific company.

        Usage in Excel: =xw_get_series(100186, "quarterly_ebitda_margins", "2022-03-31", "2025-09-30")
        """
        try:
            return get_series(accord_code, field, start_date, end_date)
        except Exception as e:
            return f"#ERROR: {e}"

    @xw.func
    @xw.arg('date_text', doc='Date in YYYY-MM-DD format')
    @xw.arg('field', doc='Database field name')
    def xw_get_quarterly_matrix(date_text, field):
        """
        Excel UDF: Get all companies' data for a specific date.

        Usage in Excel: =xw_get_quarterly_matrix("2025-06-30", "quarterly_ebitda_margins")
        """
        try:
            data = get_quarterly_matrix(date_text, field)
            # xlwings can return a 2D list which spills in Excel
            header = ["accord_code", "company_name", "sector", "mcap_category", field]
            return [header] + [list(r) for r in data]
        except Exception as e:
            return f"#ERROR: {e}"

    @xw.func
    @xw.arg('accord_code', numbers=int)
    @xw.arg('field', doc='Database field name')
    def xw_get_all_ebitda_margins(accord_code, field):
        """
        Excel UDF: Get all historical data for a specific company.

        Usage in Excel: =xw_get_all_ebitda_margins(100186, "quarterly_ebitda_margins")
        """
        try:
            data = get_all_ebitda_margins(accord_code, field)
            header = ["date", field]
            return [header] + [list(r) for r in data]
        except Exception as e:
            return f"#ERROR: {e}"

# -------------------------
# CLI helpers for testing
# -------------------------

def _cli_demo():
    """Run CLI demo to test database connectivity"""
    print("Running CLI demo (sample outputs):")
    print(f"Database file: {DB_FILE}")
    print(f"Table name: {TABLE_NAME}")
    print()

    try:
        with get_connection() as conn:
            # Test 1: Show first 5 records
            print("First 5 records from database:")
            cur = conn.execute(f"SELECT accord_code, company_name, date FROM {TABLE_NAME} LIMIT 5")
            rows = cur.fetchall()
            for r in rows:
                print(f"  {r['accord_code']}: {r['company_name']} - {r['date']}")

            # Test 2: Count records
            print()
            cur = conn.execute(f"SELECT COUNT(*) as count FROM {TABLE_NAME}")
            count = cur.fetchone()['count']
            print(f"Total records in database: {count}")

            # Test 3: Show available fields
            print()
            cur = conn.execute(f"PRAGMA table_info({TABLE_NAME})")
            columns = cur.fetchall()
            print("Available fields:")
            for col in columns:
                print(f"  - {col['name']} ({col['type']})")

            print(xw_get_quarterly_data(100186, "quarterly_ebitda_margins", "2025-06-30"))
            print(xw_get_series(100186, "quarterly_ebitda_margins", "2022-03-31", "2025-09-30"))

    except Exception as e:
        print(f"Error accessing database: {e}")
        print()
        print("Please ensure:")
        print(f"1. Database file exists at: {DB_FILE}")
        print(f"2. Table '{TABLE_NAME}' exists in the database")
        print("3. config.ini is properly configured")

if __name__ == "__main__":
    _cli_demo()
