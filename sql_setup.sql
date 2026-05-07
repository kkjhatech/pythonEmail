-- SQL Server Database Setup Script for Email Automation
-- Run this script to create the database and initial schema

-- Create Database
IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = 'EmailAutomationDB')
BEGIN
    CREATE DATABASE EmailAutomationDB;
END
GO

USE EmailAutomationDB;
GO

-- Table to track processing history
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'processing_log')
BEGIN
    CREATE TABLE processing_log (
        id INT IDENTITY(1,1) PRIMARY KEY,
        email_id NVARCHAR(255) NOT NULL,
        sender_email NVARCHAR(255),
        subject NVARCHAR(500),
        file_name NVARCHAR(500),
        table_name NVARCHAR(100),
        rows_processed INT DEFAULT 0,
        status NVARCHAR(50) DEFAULT 'PENDING',
        error_message NVARCHAR(MAX),
        processed_date DATETIME DEFAULT GETDATE(),
        CONSTRAINT UQ_processing_log_email UNIQUE (email_id, file_name)
    );
    
    -- Index for faster lookups
    CREATE INDEX IX_processing_log_email ON processing_log(email_id);
    CREATE INDEX IX_processing_log_date ON processing_log(processed_date);
    CREATE INDEX IX_processing_log_status ON processing_log(status);
END
GO

-- Table to track data import details
IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'import_details')
BEGIN
    CREATE TABLE import_details (
        id INT IDENTITY(1,1) PRIMARY KEY,
        email_id NVARCHAR(255) NOT NULL,
        file_name NVARCHAR(500),
        sheet_name NVARCHAR(100),
        table_name NVARCHAR(100),
        column_count INT,
        row_count INT,
        validation_errors NVARCHAR(MAX),
        import_date DATETIME DEFAULT GETDATE()
    );
    
    CREATE INDEX IX_import_details_email ON import_details(email_id);
END
GO

-- View to show processing statistics
IF EXISTS (SELECT * FROM sys.views WHERE name = 'vw_processing_stats')
    DROP VIEW vw_processing_stats;
GO

CREATE VIEW vw_processing_stats AS
SELECT 
    CAST(processed_date AS DATE) as process_date,
    status,
    COUNT(*) as email_count,
    SUM(rows_processed) as total_rows
FROM processing_log
GROUP BY CAST(processed_date AS DATE), status;
GO

-- Stored Procedure to get daily summary
IF EXISTS (SELECT * FROM sys.procedures WHERE name = 'sp_GetDailySummary')
    DROP PROCEDURE sp_GetDailySummary;
GO

CREATE PROCEDURE sp_GetDailySummary
    @Date DATE = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    IF @Date IS NULL
        SET @Date = CAST(GETDATE() AS DATE);
    
    SELECT 
        pl.status,
        COUNT(*) as email_count,
        SUM(pl.rows_processed) as total_rows,
        COUNT(DISTINCT pl.sender_email) as unique_senders
    FROM processing_log pl
    WHERE CAST(pl.processed_date AS DATE) = @Date
    GROUP BY pl.status;
END
GO

-- Stored Procedure to clean old logs
IF EXISTS (SELECT * FROM sys.procedures WHERE name = 'sp_CleanOldLogs')
    DROP PROCEDURE sp_CleanOldLogs;
GO

CREATE PROCEDURE sp_CleanOldLogs
    @DaysToKeep INT = 90
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @CutoffDate DATETIME = DATEADD(DAY, -@DaysToKeep, GETDATE());
    
    DELETE FROM processing_log WHERE processed_date < @CutoffDate;
    DELETE FROM import_details WHERE import_date < @CutoffDate;
    
    SELECT 
        @@ROWCOUNT as deleted_rows,
        @CutoffDate as cutoff_date;
END
GO

-- Function to check if email was already processed
IF EXISTS (SELECT * FROM sys.objects WHERE name = 'fn_IsEmailProcessed')
    DROP FUNCTION fn_IsEmailProcessed;
GO

CREATE FUNCTION fn_IsEmailProcessed(
    @EmailId NVARCHAR(255),
    @FileName NVARCHAR(500) = NULL
)
RETURNS BIT
AS
BEGIN
    DECLARE @Result BIT = 0;
    
    IF @FileName IS NOT NULL
    BEGIN
        IF EXISTS (
            SELECT 1 FROM processing_log 
            WHERE email_id = @EmailId 
            AND file_name = @FileName 
            AND status = 'SUCCESS'
        )
            SET @Result = 1;
    END
    ELSE
    BEGIN
        IF EXISTS (
            SELECT 1 FROM processing_log 
            WHERE email_id = @EmailId 
            AND status = 'SUCCESS'
        )
            SET @Result = 1;
    END
    
    RETURN @Result;
END
GO

-- Sample query to check processing history
-- SELECT TOP 10 * FROM processing_log ORDER BY processed_date DESC;

-- Sample query to get today's summary
-- EXEC sp_GetDailySummary;

-- Sample query to clean logs older than 30 days
-- EXEC sp_CleanOldLogs @DaysToKeep = 30;
