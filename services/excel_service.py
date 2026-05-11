import pandas as pd
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple
import openpyxl

from utils.logger import get_logger
from utils.validators import ExcelValidator


class ExcelService:
    """Service for handling Excel file operations."""
    
    def __init__(self, log_folder: str):
        self.logger = get_logger('ExcelService', log_folder)
        self.validator = ExcelValidator()
    
    def read_excel(
        self,
        file_path: str,
        sheet_name: Optional[str] = None,
        header_row: int = 0
    ) -> Optional[pd.DataFrame]:
        """
        Read Excel file into pandas DataFrame.
        
        Args:
            file_path: Path to Excel file
            sheet_name: Sheet to read (None for first sheet)
            header_row: Row to use as header (0-indexed)
        
        Returns:
            DataFrame or None if failed
        """
        try:
            path = Path(file_path)
            
            if not path.exists():
                self.logger.error(f"File not found: {file_path}")
                return None
            
            # Determine file type
            extension = path.suffix.lower()
            
            if extension == '.csv':
                df = pd.read_csv(file_path, dtype=str, keep_default_na=False)
            elif extension in ['.xlsx', '.xls']:
                result = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, dtype=str, keep_default_na=False)
                # Handle multi-sheet files - returns dict instead of DataFrame
                if isinstance(result, dict):
                    if not result:
                        self.logger.error(f"Excel file has no sheets: {file_path}")
                        return None
                    # Use first sheet
                    first_sheet = list(result.keys())[0]
                    df = result[first_sheet]
                    self.logger.info(f"Using first sheet: '{first_sheet}'")
                else:
                    df = result
            else:
                self.logger.error(f"Unsupported file type: {extension}")
                return None
            
            self.logger.info(f"Successfully read {file_path}: {len(df)} rows, {len(df.columns)} columns")
            return df
            
        except Exception as e:
            self.logger.error(f"Failed to read Excel file {file_path}: {str(e)}")
            return None
    
    def get_sheet_names(self, file_path: str) -> List[str]:
        """Get list of sheet names from Excel file."""
        try:
            workbook = openpyxl.load_workbook(file_path, read_only=True)
            sheets = workbook.sheetnames
            workbook.close()
            return sheets
        except Exception as e:
            self.logger.error(f"Failed to get sheet names: {str(e)}")
            return []
    
    def validate_and_prepare(
        self,
        df: pd.DataFrame,
        table_name: str,
        required_columns: Optional[List[str]] = None
    ) -> Tuple[bool, Optional[pd.DataFrame], str]:
        """
        Validate DataFrame and prepare for database insertion.
        
        Returns:
            Tuple of (is_valid, prepared_df, message)
        """
        # Validate
        is_valid, errors, warnings = self.validator.validate_dataframe(
            df,
            required_columns=required_columns,
            allow_empty=False
        )
        
        for warning in warnings:
            self.logger.warning(warning)
        
        if not is_valid:
            error_msg = '; '.join(errors)
            self.logger.error(f"Validation failed: {error_msg}")
            return False, None, error_msg
        
        # Prepare
        prepared_df, table_name, columns = self.validator.prepare_for_insert(
            df,
            table_name,
            sanitize_columns=True
        )
        
        self.logger.info(f"Data validated and prepared for table: {table_name}")
        return True, prepared_df, f"Ready to insert {len(prepared_df)} rows into {table_name}"
    
    def infer_sql_types(self, df: pd.DataFrame) -> Dict[str, str]:
        """
        Infer SQL data types from DataFrame columns.
        
        Returns:
            Dictionary mapping column names to SQL types
        """
        type_mapping = {
            'int64': 'BIGINT',
            'int32': 'INT',
            'int16': 'SMALLINT',
            'float64': 'FLOAT',
            'float32': 'REAL',
            'object': 'NVARCHAR(500)',
            'bool': 'BIT',
            'datetime64[ns]': 'DATETIME',
        }
        
        sql_types = {}
        for col in df.columns:
            dtype = str(df[col].dtype)
            
            # Check if column contains datetime strings
            if dtype == 'object':
                sample = df[col].dropna().head(10)
                if self._is_datetime_column(sample):
                    sql_types[col] = 'DATETIME'
                    continue
            
            sql_types[col] = type_mapping.get(dtype, 'NVARCHAR(500)')
        
        # DEBUG: Log SQL types
        self.logger.info(f"SQL types inferred: {sql_types}")
        
        return sql_types
    
    def _is_datetime_column(self, series: pd.Series) -> bool:
        """Check if column contains datetime values."""
        try:
            pd.to_datetime(series, errors='raise')
            return True
        except:
            return False
    
    def generate_create_table_sql(
        self,
        table_name: str,
        df: pd.DataFrame,
        email_master_a: int = None,
        email_details_a: int = None,
        include_email_id: bool = True
    ) -> str:
        """
        Generate CREATE TABLE SQL statement based on DataFrame columns.
        
        Args:
            table_name: Name for the new table
            df: DataFrame to analyze
            email_master_a: Email_Master_A ID for prefix
            email_details_a: Email_Details_A ID for prefix
            include_email_id: Whether to add email_id tracking column
            
        Returns:
            CREATE TABLE SQL statement
        """
        # Infer SQL types
        sql_types = self.infer_sql_types(df)
        
        # Build table name with prefix if IDs provided
        if email_master_a and email_details_a:
            # Extract filename from table_name
            filename = table_name
            if '.' in filename:
                filename = filename.split('.')[0]
            # Create prefixed table name with just numeric values
            prefixed_table_name = f"PY_{email_master_a}_{email_details_a}_{filename}"
        else:
            prefixed_table_name = table_name
        
        # Build column definitions
        columns = []
        
        # Add ID column
        columns.append("    id INT IDENTITY(1,1) PRIMARY KEY")
        
        # Add tracking columns only for non-prefixed tables
        # For prefixed tables (PY_1_2_...), add Email_Details_A for join purposes
        if include_email_id and not (email_master_a and email_details_a):
            columns.append("    sender_email NVARCHAR(255)")
            columns.append("    processed_date DATETIME DEFAULT GETDATE()")
        elif include_email_id and (email_master_a and email_details_a):
            # For prefixed tables, add Email_Details_A for join and processed_date
            columns.append(f"    [Email_Details_A] INT")
            columns.append("    processed_date DATETIME DEFAULT GETDATE()")
        
        # Add data columns
        for col in df.columns:
            sql_type = sql_types.get(col, 'NVARCHAR(500)')
            # Sanitize column name for SQL
            safe_col = str(col).strip().replace(' ', '_').replace('-', '_')
            safe_col = safe_col.rstrip('_')  # Remove trailing underscores
            if safe_col[0].isdigit():
                safe_col = f"col_{safe_col}"
            # Wrap in brackets to handle reserved keywords (e.g., 'Add', 'Select', etc.)
            # But don't double-wrap if already has brackets
            if not safe_col.startswith('[') and not safe_col.endswith(']'):
                safe_col = f"[{safe_col}]"
            columns.append(f"    {safe_col} {sql_type}")
        
        # Build CREATE TABLE statement
        create_sql = f"""CREATE TABLE {prefixed_table_name} (
{',\n'.join(columns)}
);"""
        
        return create_sql
    
    def create_data_preview(
        self,
        df: pd.DataFrame,
        max_rows: int = 5
    ) -> Dict[str, Any]:
        """
        Create a preview of the DataFrame for logging.
        
        Returns:
            Dictionary with preview information
        """
        return {
            'total_rows': len(df),
            'total_columns': len(df.columns),
            'columns': list(df.columns),
            'sample_data': df.head(max_rows).to_dict(orient='records'),
            'dtypes': {col: str(dtype) for col, dtype in df.dtypes.items()},
            'null_counts': df.isnull().sum().to_dict()
        }
