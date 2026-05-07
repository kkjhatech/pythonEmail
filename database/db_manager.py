import pyodbc
import pandas as pd
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime

from config.settings import Settings
from utils.logger import get_logger


class DatabaseManager:
    """Manager for SQL Server database operations."""
    
    def __init__(self, settings: Settings):
        self.settings = settings
        self.logger = get_logger('DatabaseManager', settings.log_folder)
        self.connection: Optional[pyodbc.Connection] = None
    
    def connect(self) -> bool:
        """Establish database connection."""
        try:
            connection_string = self.settings.get_db_connection_string()
            self.connection = pyodbc.connect(connection_string, timeout=30)
            self.logger.info("Database connection established")
            return True
        except Exception as e:
            self.logger.error(f"Failed to connect to database: {str(e)}")
            return False
    
    def disconnect(self):
        """Close database connection."""
        if self.connection:
            try:
                self.connection.close()
                self.logger.info("Database connection closed")
            except Exception as e:
                self.logger.warning(f"Error closing database connection: {str(e)}")
            finally:
                self.connection = None
    
    def __enter__(self):
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()
    
    def test_connection(self) -> bool:
        """Test database connectivity."""
        try:
            with self.connection.cursor() as cursor:
                cursor.execute("SELECT 1")
                cursor.fetchone()
            return True
        except Exception as e:
            self.logger.error(f"Connection test failed: {str(e)}")
            return False
    
    def table_exists(self, table_name: str) -> bool:
        """Check if table exists in database."""
        try:
            with self.connection.cursor() as cursor:
                cursor.execute("""
                    SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES 
                    WHERE TABLE_NAME = ?
                """, (table_name,))
                result = cursor.fetchone()
                return result[0] > 0
        except Exception as e:
            self.logger.error(f"Error checking table existence: {str(e)}")
            return False
    
    def create_table(self, sql: str) -> bool:
        """Execute CREATE TABLE statement."""
        try:
            with self.connection.cursor() as cursor:
                cursor.execute(sql)
                self.connection.commit()
            self.logger.info("Table created successfully")
            return True
        except Exception as e:
            self.logger.error(f"Failed to create table: {str(e)}")
            self.connection.rollback()
            return False
    
    def get_table_columns(self, table_name: str) -> List[str]:
        """Get list of column names for a table."""
        try:
            with self.connection.cursor() as cursor:
                cursor.execute("""
                    SELECT COLUMN_NAME 
                    FROM INFORMATION_SCHEMA.COLUMNS 
                    WHERE TABLE_NAME = ?
                    ORDER BY ORDINAL_POSITION
                """, (table_name,))
                # Strip brackets from column names if present (for proper matching)
                cols = [row[0] for row in cursor.fetchall()]
                return [col.strip('[]') for col in cols]
        except Exception as e:
            self.logger.error(f"Error getting table columns: {str(e)}")
            return []
    
    def check_duplicate(
        self,
        table_name: str,
        email_id: str,
        data_hash: Optional[str] = None
    ) -> bool:
        """
        Check if data from this email has already been processed.
        
        Args:
            table_name: Target table
            email_id: Email identifier
            data_hash: Optional hash of data content
        
        Returns:
            True if duplicate exists
        """
        try:
            with self.connection.cursor() as cursor:
                if data_hash:
                    cursor.execute(f"""
                        SELECT COUNT(*) FROM {table_name} 
                        WHERE email_id = ? OR data_hash = ?
                    """, (email_id, data_hash))
                else:
                    cursor.execute(f"""
                        SELECT COUNT(*) FROM {table_name} 
                        WHERE email_id = ?
                    """, (email_id,))
                
                result = cursor.fetchone()
                return result[0] > 0
        except Exception as e:
            self.logger.error(f"Error checking duplicates: {str(e)}")
            return False
    
    def insert_dataframe(
        self,
        df: pd.DataFrame,
        table_name: str,
        email_id: Optional[str] = None,
        batch_size: int = 1000
    ) -> Tuple[bool, int, str]:
        """
        Insert DataFrame into SQL Server table.
        
        Args:
            df: DataFrame to insert
            table_name: Target table name
            email_id: Optional email identifier for tracking
            batch_size: Number of rows per batch
        
        Returns:
            Tuple of (success, rows_inserted, message)
        """
        if df.empty:
            return True, 0, "No data to insert"
        
        try:
            # Get table columns
            table_columns = self.get_table_columns(table_name)
            self.logger.debug(f"Table columns: {table_columns}")
            self.logger.debug(f"DataFrame columns: {list(df.columns)}")
            
            # Filter DataFrame columns to match table
            df_columns = [col for col in df.columns if col in table_columns]
            
            if not df_columns:
                error_msg = f"No matching columns. Table has: {table_columns}, DataFrame has: {list(df.columns)}"
                self.logger.error(error_msg)
                return False, 0, error_msg
            
            # Build INSERT statement with bracket-wrapped column names
            columns_bracketed = [f"[{col}]" for col in df_columns]
            columns_str = ', '.join(columns_bracketed)
            placeholders = ', '.join(['?' for _ in df_columns])
            
            insert_sql = f"INSERT INTO {table_name} ({columns_str}) VALUES ({placeholders})"
            
            rows_inserted = 0
            self.logger.info(f"Inserting {len(df)} rows into {table_name} using columns: {df_columns}")
            
            with self.connection.cursor() as cursor:
                # Insert in batches
                for start_idx in range(0, len(df), batch_size):
                    batch = df.iloc[start_idx:start_idx + batch_size]
                    
                    for _, row in batch.iterrows():
                        values = [row[col] for col in df_columns]
                        cursor.execute(insert_sql, values)
                    
                    rows_inserted += len(batch)
                
                self.connection.commit()
            
            message = f"Successfully inserted {rows_inserted} rows into {table_name}"
            self.logger.info(message)
            return True, rows_inserted, message
            
        except Exception as e:
            error_msg = f"Failed to insert data: {str(e)}"
            self.logger.error(error_msg)
            self.logger.error(f"SQL: {insert_sql if 'insert_sql' in locals() else 'N/A'}")
            if df_columns:
                sample_values = [str(df.iloc[0][col])[:50] for col in df_columns[:3]]
                self.logger.error(f"Sample values (first row, first 3 cols): {sample_values}")
            self.connection.rollback()
            return False, 0, error_msg
    
    def insert_with_tracking(
        self,
        df: pd.DataFrame,
        table_name: str,
        email_id: str,
        check_duplicates: bool = True
    ) -> Tuple[bool, int, str]:
        """
        Insert data with email tracking and duplicate checking.
        
        Args:
            df: DataFrame to insert
            table_name: Target table name
            email_id: Email identifier
            check_duplicates: Whether to check for existing data
        
        Returns:
            Tuple of (success, rows_inserted, message)
        """
        if check_duplicates:
            if self.check_duplicate(table_name, email_id):
                msg = f"Data from email {email_id} already exists. Skipping."
                self.logger.info(msg)
                return True, 0, msg
        
        # Add tracking column if table has it
        table_columns = self.get_table_columns(table_name)
        
        if 'email_id' in table_columns and email_id:
            df = df.copy()
            df['email_id'] = email_id
        
        return self.insert_dataframe(df, table_name, email_id)
    
    def execute_query(self, sql: str, params: Optional[tuple] = None) -> List[tuple]:
        """Execute a custom SQL query."""
        try:
            with self.connection.cursor() as cursor:
                if params:
                    cursor.execute(sql, params)
                else:
                    cursor.execute(sql)
                
                if sql.strip().upper().startswith('SELECT'):
                    return cursor.fetchall()
                else:
                    self.connection.commit()
                    return []
        except Exception as e:
            self.logger.error(f"Query execution failed: {str(e)}")
            self.connection.rollback()
            raise
    
    def get_processing_stats(self) -> Dict[str, Any]:
        """Get statistics about processed data."""
        try:
            with self.connection.cursor() as cursor:
                # Get all tables with email_id column
                cursor.execute("""
                    SELECT TABLE_NAME 
                    FROM INFORMATION_SCHEMA.COLUMNS 
                    WHERE COLUMN_NAME = 'email_id'
                """)
                tables = [row[0] for row in cursor.fetchall()]
                
                stats = {}
                for table in tables:
                    cursor.execute(f"""
                        SELECT 
                            COUNT(*) as total_rows,
                            COUNT(DISTINCT email_id) as unique_emails,
                            MAX(processed_date) as last_processed
                        FROM {table}
                    """)
                    row = cursor.fetchone()
                    stats[table] = {
                        'total_rows': row[0],
                        'unique_emails': row[1],
                        'last_processed': row[2]
                    }
                
                return stats
        except Exception as e:
            self.logger.error(f"Error getting stats: {str(e)}")
            return {}
