import pandas as pd
from typing import Dict, List, Optional, Tuple, Any
from datetime import datetime


class ExcelValidator:
    """Validates Excel data before database insertion."""
    
    def __init__(self):
        self.errors: List[str] = []
        self.warnings: List[str] = []
    
    def validate_dataframe(
        self,
        df: pd.DataFrame,
        required_columns: Optional[List[str]] = None,
        column_types: Optional[Dict[str, type]] = None,
        allow_empty: bool = False
    ) -> Tuple[bool, List[str], List[str]]:
        """
        Validate a pandas DataFrame.
        
        Args:
            df: DataFrame to validate
            required_columns: List of required column names
            column_types: Dictionary of column names to expected types
            allow_empty: Whether to allow empty DataFrames
        
        Returns:
            Tuple of (is_valid, errors, warnings)
        """
        self.errors = []
        self.warnings = []
        
        # Check if DataFrame is empty
        if df.empty:
            if not allow_empty:
                self.errors.append("DataFrame is empty")
            return (len(self.errors) == 0, self.errors, self.warnings)
        
        # Validate required columns
        if required_columns:
            missing_cols = [col for col in required_columns if col not in df.columns]
            if missing_cols:
                self.errors.append(f"Missing required columns: {missing_cols}")
        
        # Validate column types
        if column_types:
            for col, expected_type in column_types.items():
                if col in df.columns:
                    actual_type = df[col].dtype
                    if not self._check_type_compatibility(actual_type, expected_type):
                        self.warnings.append(
                            f"Column '{col}' has type {actual_type}, expected {expected_type}"
                        )
        
        # Check for null values in required columns
        if required_columns:
            for col in required_columns:
                if col in df.columns:
                    null_count = df[col].isnull().sum()
                    if null_count > 0:
                        self.warnings.append(f"Column '{col}' has {null_count} null values")
        
        # Check for duplicate rows
        dup_count = df.duplicated().sum()
        if dup_count > 0:
            self.warnings.append(f"DataFrame contains {dup_count} duplicate rows")
        
        return (len(self.errors) == 0, self.errors, self.warnings)
    
    def _check_type_compatibility(self, actual_dtype, expected_type: type) -> bool:
        """Check if actual dtype is compatible with expected type."""
        type_mapping = {
            int: ['int64', 'int32', 'int16', 'int8'],
            float: ['float64', 'float32'],
            str: ['object', 'string'],
            bool: ['bool'],
            datetime: ['datetime64[ns]', 'datetime64'],
        }
        
        actual_str = str(actual_dtype)
        
        if expected_type in type_mapping:
            return actual_str in type_mapping[expected_type]
        
        return True
    
    def sanitize_column_names(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Sanitize column names for SQL compatibility.

        Args:
            df: DataFrame with original column names

        Returns:
            DataFrame with sanitized column names
        """
        import re
        from datetime import datetime
        df = df.copy()

        new_columns = []
        for col in df.columns:
            # Handle datetime-like column names (e.g., pandas Timestamp) or Excel serial dates (numeric)
            if isinstance(col, (datetime, pd.Timestamp)):
                # Format as DD-MMM-YYYY
                clean_col = col.strftime("%d-%b-%Y")
            elif isinstance(col, (int, float)):
                # Treat as Excel serial date (days since 1899-12-30)
                try:
                    base_date = datetime(1899, 12, 30)
                    clean_col = (base_date + pd.Timedelta(days=int(col))).strftime("%d-%b-%Y")
                except Exception:
                    clean_col = str(col)
            else:
                clean_col = str(col).strip()
                # Expand two‑digit year dates like "01-May-26" to "01-May-2026"
                try:
                    if re.match(r"^\d{2}-[A-Za-z]{3}-\d{2}$", clean_col):
                        dt = datetime.strptime(clean_col, "%d-%b-%y")
                        clean_col = dt.strftime("%d-%b-%Y")
                except Exception:
                    pass
            # Determine if this column is a date header (e.g., 01-May-2026 or 01_May_2026)
            is_date_header = bool(re.match(r"^\d{2}[-_]\w{3}[-_]\d{4}$", clean_col))
            # Replace spaces with underscores; preserve hyphens for date headers
            clean_col = clean_col.replace(' ', '_')
            if not is_date_header:
                # For non‑date columns, replace any non‑alphanumeric/underscore characters (including hyphens) with underscore
                clean_col = ''.join(c if c.isalnum() or c == '_' else '_' for c in clean_col)
            # Remove trailing underscores
            clean_col = clean_col.rstrip('_')
            # Do not prefix with "col_" for date headers or any column containing letters
            new_columns.append(clean_col)

        df.columns = new_columns
        return df
    
    def prepare_for_insert(
        self,
        df: pd.DataFrame,
        table_name: str,
        sanitize_columns: bool = True
    ) -> Tuple[pd.DataFrame, str, List[str]]:
        """
        Prepare DataFrame for database insertion.
        
        Args:
            df: DataFrame to prepare
            table_name: Target table name
            sanitize_columns: Whether to sanitize column names
        
        Returns:
            Tuple of (prepared_df, table_name, column_names)
        """
        if sanitize_columns:
            df = self.sanitize_column_names(df)
        
        # Convert datetime columns to string for SQL compatibility
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M:%S')
        
        # Replace NaN with None for SQL NULL
        df = df.where(pd.notnull(df), None)
        
        # FORCE all columns to string to ensure character data is preserved
        df = df.astype(str)
        df = df.replace('None', '')  # Replace string 'None' with empty string
        
        return df, table_name, list(df.columns)
