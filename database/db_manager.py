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
        sender_email: Optional[str] = None,
        email_details_a: Optional[int] = None,
        batch_size: int = 10000
    ) -> Tuple[bool, int, str]:
        """
        Insert DataFrame into SQL Server table.
        
        Args:
            df: DataFrame to insert
            table_name: Target table name
            sender_email: Sender email to include
            email_details_a: Email_Details_A ID for prefixed tables
            batch_size: Number of rows to insert at once
            
        Returns:
            Tuple of (success, rows_inserted, message)
        """
        try:
            with self.connection.cursor() as cursor:
                # Enable fast executemany for bulk inserts
                cursor.fast_executemany = True

                # Get table columns
                table_columns = self.get_table_columns(table_name)

                # Check if this is a prefixed table (starts with PY_ followed by numbers)
                is_prefixed_table = table_name.startswith("PY_") and "_" in table_name[3:]
                
                # Add columns based on table type
                if is_prefixed_table and email_details_a and 'Email_Details_A' in table_columns:
                    # For prefixed tables, add Email_Details_A column
                    df = df.copy()
                    df['Email_Details_A'] = email_details_a
                elif 'sender_email' in table_columns and sender_email and not is_prefixed_table:
                    # For regular tables, add sender_email column
                    df = df.copy()
                    df['sender_email'] = sender_email
                
                # Prepare column names for SQL
                columns = []
                for col in df.columns:
                    if col in table_columns:
                        # Don't double-wrap if column already has brackets
                        if col.startswith('[') and col.endswith(']'):
                            columns.append(col)
                        else:
                            columns.append(f"[{col}]")
                
                if not columns:
                    return False, 0, "No matching columns found"
                
                # Prepare INSERT statement
                placeholders = ", ".join(["?" for _ in columns])
                insert_sql = f"INSERT INTO {table_name} ({', '.join(columns)}) VALUES ({placeholders})"
                
                # Convert DataFrame to list of tuples
                # Use original column names from DataFrame
                original_columns = [col.strip('[]') for col in columns] if all(c.startswith('[') and c.endswith(']') for c in columns) else columns
                data = df[[col for col in df.columns if col in table_columns]].values.tolist()
                
                # Insert in batches
                rows_inserted = 0
                for i in range(0, len(data), batch_size):
                    batch = data[i:i + batch_size]
                    cursor.executemany(insert_sql, batch)
                    rows_inserted += len(batch)
                    self.logger.info(f"Inserted batch of {len(batch)} rows")
                
                self.connection.commit()
                message = f"Successfully inserted {rows_inserted} rows into {table_name}"
                self.logger.info(message)
                return True, rows_inserted, message
                
        except Exception as e:
            error_msg = f"Failed to insert data: {str(e)}"
            self.logger.error(error_msg)
            self.logger.error(f"SQL: {insert_sql if 'insert_sql' in locals() else 'N/A'}")
            if df is not None and len(df) > 0:
                df_columns = df.columns.tolist()
                sample_values = [str(df.iloc[0][col])[:50] for col in df_columns[:3]]
                self.logger.error(f"Sample values (first row, first 3 cols): {sample_values}")
            self.connection.rollback()
            return False, 0, error_msg
    
    def insert_with_tracking(
        self,
        df: pd.DataFrame,
        table_name: str,
        sender_email: str,
        check_duplicates: bool = True
    ) -> Tuple[bool, int, str]:
        """
        Insert data with email tracking and duplicate checking.
        
        Args:
            df: DataFrame to insert
            table_name: Target table name
            sender_email: Sender email address
            check_duplicates: Whether to check for existing data
        
        Returns:
            Tuple of (success, rows_inserted, message)
        """
        self.logger.info(f"Insert with tracking: table={table_name}, sender={sender_email}, check_dup={check_duplicates}")
        
        if check_duplicates:
            is_dup = self.check_duplicate(table_name, sender_email)
            self.logger.info(f"Duplicate check result: {is_dup}")
            if is_dup:
                msg = f"Data from email {sender_email} already exists. Skipping."
                self.logger.info(msg)
                return True, 0, msg
        
        # Add tracking column if table has it
        table_columns = self.get_table_columns(table_name)
        self.logger.info(f"Table columns: {table_columns}")
        
        if 'sender_email' in table_columns and sender_email:
            df = df.copy()
            df['sender_email'] = sender_email
        
        return self.insert_dataframe(df, table_name, sender_email)
    
    def insert_email_master(self, actual_email: str, created_by: str = "System") -> Tuple[bool, int, str]:
        """
        Insert email into Email_Master table if not exists.
        
        Returns:
            Tuple of (success, email_master_a_id, message)
        """
        try:
            cursor = self.connection.cursor()
            
            # Check if email already exists
            cursor.execute("SELECT Email_Master_A FROM Email_Master WHERE EmailID = ?", (actual_email,))
            existing = cursor.fetchone()
            
            if existing:
                email_master_a = existing[0]
                self.logger.info(f"Email {actual_email} already exists in Email_Master with ID {email_master_a}")
                return True, email_master_a, "Email already exists"
            
            # Check if actual_email is valid
            if not actual_email:
                self.logger.error("Email cannot be empty")
                return False, 0, "Email cannot be empty"
            
            # Insert new email using stored procedure
            self.logger.info(f"Inserting email: {actual_email}")
            cursor.execute("EXEC usp_insert_email @Email_ID = ?, @CreatedBy = ?", (actual_email, created_by))
            
            # Get the inserted ID - try multiple methods
            try:
                # Method 1: Try to get result from stored procedure
                result = cursor.fetchone()
                if result and result[0] is not None:
                    email_master_a = int(result[0])
                    self.connection.commit()
                    self.logger.info(f"Inserted email {actual_email} into Email_Master with ID {email_master_a}")
                    return True, email_master_a, "Email inserted successfully"
            except:
                pass
            
            # Method 2: Use SCOPE_IDENTITY() directly
            try:
                cursor.execute("SELECT SCOPE_IDENTITY()")
                result = cursor.fetchone()
                if result and result[0] is not None:
                    email_master_a = int(result[0])
                    self.connection.commit()
                    self.logger.info(f"Inserted email {actual_email} into Email_Master with ID {email_master_a} (using SCOPE_IDENTITY)")
                    return True, email_master_a, "Email inserted successfully"
            except:
                pass
            
            # Method 3: Use IDENT_CURRENT as last resort
            try:
                cursor.execute("SELECT IDENT_CURRENT('Email_Master')")
                result = cursor.fetchone()
                if result and result[0] is not None:
                    email_master_a = int(result[0])
                    self.connection.commit()
                    self.logger.info(f"Inserted email {actual_email} into Email_Master with ID {email_master_a} (using IDENT_CURRENT)")
                    return True, email_master_a, "Email inserted successfully"
            except:
                pass
            
            self.logger.error("Failed to get inserted ID from all methods")
            return False, 0, "Failed to get inserted ID"
            
        except Exception as e:
            self.logger.error(f"Failed to insert into Email_Master: {str(e)}")
            # Try to rollback if possible
            try:
                self.connection.rollback()
            except:
                pass
            return False, 0, str(e)
    
    def insert_email_details(self, email_master_a: int, subject: str, sheet_name: str, 
                           total_rows: int, received_date: datetime) -> Tuple[bool, int, str]:
        """
        Insert email details into Email_Details table.
        
        Returns:
            Tuple of (success, email_details_a_id, message)
        """
        try:
            cursor = self.connection.cursor()
            
            # Validate required parameters
            if not subject:
                subject = "No Subject"
            if not sheet_name:
                sheet_name = "Sheet1"
            if total_rows is None or total_rows <= 0:
                total_rows = 0
            if received_date is None:
                received_date = datetime.now()
            
            # Insert email details - try stored procedure first, then direct SQL
            self.logger.info(f"Inserting email details for Email_Master_A {email_master_a}")
            
            try:
                # Try stored procedure first
                cursor.execute(
                    "EXEC usp_insert_email_details @EmailID_N = ?, @Subject_Name = ?, @SheetName = ?, @TotalRows = ?, @ReceivedDate = ?",
                    (email_master_a, subject, sheet_name, total_rows, received_date)
                )
            except Exception as proc_error:
                self.logger.warning(f"Stored procedure failed: {proc_error}. Trying direct SQL insert.")
                # Fallback to direct SQL
                cursor.execute(
                    "INSERT INTO Email_Details (EmailID_N, Subject_Name, SheetName, TotalRows, ReceivedDate) VALUES (?, ?, ?, ?, ?)",
                    (email_master_a, subject, sheet_name, total_rows, received_date)
                )
            
            # Get the inserted ID - try multiple methods
            try:
                # Method 1: Try to get result from stored procedure
                result = cursor.fetchone()
                if result and result[0] is not None:
                    email_details_a = int(result[0])
                    self.connection.commit()
                    self.logger.info(f"Inserted email details for Email_Master_A {email_master_a} with ID {email_details_a}")
                    return True, email_details_a, "Email details inserted successfully"
            except:
                pass
            
            # Method 2: Use SCOPE_IDENTITY() directly
            try:
                cursor.execute("SELECT SCOPE_IDENTITY()")
                result = cursor.fetchone()
                if result and result[0] is not None:
                    email_details_a = int(result[0])
                    self.connection.commit()
                    self.logger.info(f"Inserted email details for Email_Master_A {email_master_a} with ID {email_details_a} (using SCOPE_IDENTITY)")
                    return True, email_details_a, "Email details inserted successfully"
            except:
                pass
            
            # Method 3: Use IDENT_CURRENT as last resort
            try:
                cursor.execute("SELECT IDENT_CURRENT('Email_Details')")
                result = cursor.fetchone()
                if result and result[0] is not None:
                    email_details_a = int(result[0])
                    self.connection.commit()
                    self.logger.info(f"Inserted email details for Email_Master_A {email_master_a} with ID {email_details_a} (using IDENT_CURRENT)")
                    return True, email_details_a, "Email details inserted successfully"
            except:
                pass
            
            self.logger.error("Failed to get inserted ID from all methods")
            return False, 0, "Failed to get inserted ID"
            
        except Exception as e:
            self.logger.error(f"Failed to insert into Email_Details: {str(e)}")
            # Try to rollback if possible
            try:
                self.connection.rollback()
            except:
                pass
            return False, 0, str(e)
    
    def execute_query(self, sql: str, params: Optional[tuple] = None) -> List[tuple]:
        """Execute a custom SQL query."""
        try:
            cursor = self.connection.cursor()
            if params:
                cursor.execute(sql, params)
            else:
                cursor.execute(sql)
            
            if cursor.description:
                return cursor.fetchall()
            return []
        except Exception as e:
            self.logger.error(f"Query execution failed: {str(e)}")
            return []
    
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
