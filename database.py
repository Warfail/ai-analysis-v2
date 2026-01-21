# database.py - Fixed version
import os
import sys
from datetime import datetime
import time

# Get connection string from environment or use local SQLite for dev
DATABASE_URL = os.getenv("DATABASE_URL", "")
USE_SQLITE = os.getenv("USE_SQLITE", "true").lower() == "true"

def get_connection():
    """Get database connection (PostgreSQL for production, SQLite for local dev)"""
    
    if USE_SQLITE:
        # Local development with SQLite
        try:
            import sqlite3
            print("🔧 Using SQLite for local development")
            conn = sqlite3.connect('reports.db', check_same_thread=False)
            conn.row_factory = sqlite3.Row
            return conn
        except ImportError:
            print("❌ SQLite not available")
            raise
    
    # Production with PostgreSQL
    if not DATABASE_URL:
        print("⚠️ DATABASE_URL not set. Falling back to SQLite for development.")
        # Fall back to SQLite
        try:
            import sqlite3
            conn = sqlite3.connect('reports.db', check_same_thread=False)
            conn.row_factory = sqlite3.Row
            return conn
        except:
            raise ValueError("Neither DATABASE_URL set nor SQLite available")
    
    # Production - try PostgreSQL
    try:
        import psycopg2
        from psycopg2.extras import RealDictCursor
        
        # Retry logic for production
        max_retries = 3
        for attempt in range(max_retries):
            try:
                conn = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
                print(f"✅ Connected to PostgreSQL (attempt {attempt + 1}/{max_retries})")
                return conn
            except Exception as e:
                print(f"⚠️ PostgreSQL connection failed (attempt {attempt + 1}): {e}")
                if attempt < max_retries - 1:
                    time.sleep(2)
                else:
                    print("❌ All PostgreSQL connection attempts failed")
                    raise
    except ImportError:
        print("❌ psycopg2 not installed. Install with: pip install psycopg2-binary")
        raise

def setup_database():
    """Create tables if they don't exist"""
    print("🔧 Setting up database...")
    print(f"📊 Mode: {'SQLite' if USE_SQLITE else 'PostgreSQL'}")
    print(f"📁 DATABASE_URL: {'Set' if DATABASE_URL else 'Not set'}")
    
    conn = None
    try:
        conn = get_connection()
        
        if USE_SQLITE or not DATABASE_URL:
            # SQLite setup
            import sqlite3
            cursor = conn.cursor()
            
            # Check if table exists
            cursor.execute("""
                SELECT name FROM sqlite_master 
                WHERE type='table' AND name='reports'
            """)
            
            if not cursor.fetchone():
                print("📦 Creating SQLite table...")
                cursor.execute("""
                    CREATE TABLE reports (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        department TEXT,
                        report_date TEXT,
                        filename TEXT,
                        file_type TEXT,
                        file_size INTEGER,
                        content_preview TEXT,
                        word_count INTEGER,
                        upload_date TEXT,
                        ai_analysis TEXT,
                        ai_conclusion TEXT,
                        total_reports INTEGER DEFAULT 0,
                        total_problems INTEGER DEFAULT 0,
                        total_achievements INTEGER DEFAULT 0,
                        sentiment_label TEXT DEFAULT 'neutral'
                    )
                """)
                
                # Create indexes
                cursor.execute("CREATE INDEX idx_reports_department ON reports(department)")
                cursor.execute("CREATE INDEX idx_reports_upload_date ON reports(upload_date)")
                cursor.execute("CREATE INDEX idx_reports_sentiment ON reports(sentiment_label)")
                
                print("✅ SQLite database created")
            else:
                print("✅ SQLite database already exists")
            
            conn.commit()
            
        else:
            # PostgreSQL setup for Supabase
            import psycopg2
            cursor = conn.cursor()
            
            # Check if table exists
            cursor.execute("""
                SELECT EXISTS (
                    SELECT FROM information_schema.tables 
                    WHERE table_name = 'reports'
                )
            """)
            
            table_exists = cursor.fetchone()["exists"]
            
            if not table_exists:
                print("📦 Creating PostgreSQL table...")
                cursor.execute("""
                    CREATE TABLE reports (
                        id SERIAL PRIMARY KEY,
                        department TEXT,
                        report_date TEXT,
                        filename TEXT,
                        file_type TEXT,
                        file_size INTEGER,
                        content_preview TEXT,
                        word_count INTEGER,
                        upload_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        ai_analysis JSONB,
                        ai_conclusion JSONB,
                        total_reports INTEGER DEFAULT 0,
                        total_problems INTEGER DEFAULT 0,
                        total_achievements INTEGER DEFAULT 0,
                        sentiment_label TEXT DEFAULT 'neutral'
                    )
                """)
                
                # Create indexes
                cursor.execute("""
                    CREATE INDEX idx_reports_department 
                    ON reports(department)
                """)
                cursor.execute("""
                    CREATE INDEX idx_reports_upload_date 
                    ON reports(upload_date)
                """)
                cursor.execute("""
                    CREATE INDEX idx_reports_sentiment 
                    ON reports(sentiment_label)
                """)
                
                print("✅ PostgreSQL database created")
            else:
                print("✅ PostgreSQL database already exists")
            
            conn.commit()
        
    except Exception as e:
        print(f"❌ Database setup error: {e}")
        if conn:
            conn.rollback()
        raise
    finally:
        if conn:
            cursor.close()
            conn.close()

# Initialize database
if __name__ != "__main__":
    # Don't crash on import, wait for actual use
    pass

def init_db():
    """Call this function to initialize the database"""
    setup_database()

# Add these functions to your existing database.py file

def execute_query(query, params=None, fetch_one=False, fetch_all=False, commit=True):
    """Helper function to execute queries safely"""
    conn = None
    cursor = None
    
    try:
        conn = get_connection()
        cursor = conn.cursor()
        
        # Execute query
        cursor.execute(query, params or ())
        
        # Fetch results if needed
        if fetch_one:
            result = cursor.fetchone()
            if result and hasattr(result, '_asdict'):
                result = dict(result)
        elif fetch_all:
            result = cursor.fetchall()
            if result and hasattr(result[0], '_asdict'):
                result = [dict(row) for row in result]
        else:
            result = None
        
        # Commit if needed
        if commit:
            conn.commit()
        
        return result
        
    except Exception as e:
        print(f"❌ Query error: {e}")
        print(f"   Query: {query}")
        print(f"   Params: {params}")
        if conn:
            conn.rollback()
        raise
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

def execute_many(query, params_list):
    """Execute the same query with multiple parameter sets"""
    conn = None
    cursor = None
    
    try:
        conn = get_connection()
        cursor = conn.cursor()
        
        cursor.executemany(query, params_list)
        conn.commit()
        
        return cursor.rowcount
        
    except Exception as e:
        print(f"❌ execute_many error: {e}")
        if conn:
            conn.rollback()
        raise
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

# Helper functions for common operations
def fetch_one(query, params=None):
    """Fetch a single row"""
    return execute_query(query, params, fetch_one=True)

def fetch_all(query, params=None):
    """Fetch all rows"""
    return execute_query(query, params, fetch_all=True)

def insert(query, params=None):
    """Insert a row and return the ID"""
    conn = None
    cursor = None
    
    try:
        conn = get_connection()
        cursor = conn.cursor()
        
        cursor.execute(query, params or ())
        
        # Try to get the inserted ID
        if "RETURNING" in query.upper():
            result = cursor.fetchone()
            inserted_id = result["id"] if result else None
        else:
            inserted_id = cursor.lastrowid if hasattr(cursor, 'lastrowid') else None
        
        conn.commit()
        return inserted_id
        
    except Exception as e:
        print(f"❌ Insert error: {e}")
        if conn:
            conn.rollback()
        raise
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

def update(query, params=None):
    """Update rows and return count of affected rows"""
    return execute_query(query, params, commit=True)    