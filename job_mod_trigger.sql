CREATE OR ALTER  TRIGGER [DBO].[TRIG_SYSJOBS_INSERT_UPDATE_DELETE]
ON [DBO].[SYSJOBS]
FOR INSERT, UPDATE, DELETE
AS 
BEGIN
    SET NOCOUNT ON;

IF OBJECT_ID(N'TEMPDB..#DELETED') IS NOT NULL
BEGIN
DROP TABLE #DELETED
END

	CREATE TABLE #DELETED  ([JOB_ID] [UNIQUEIDENTIFIER] NOT NULL,
	[ORIGINATING_SERVER_ID] [INT] NOT NULL,
	[NAME] [SYSNAME] NOT NULL,
	[ENABLED] [TINYINT] NOT NULL,
	[DESCRIPTION] [NVARCHAR](512) NULL,
	[START_STEP_ID] [INT] NOT NULL,
	[CATEGORY_ID] [INT] NOT NULL,
	[OWNER_SID] [VARBINARY](85) NOT NULL,
	[NOTIFY_LEVEL_EVENTLOG] [INT] NOT NULL,
	[NOTIFY_LEVEL_EMAIL] [INT] NOT NULL,
	[NOTIFY_LEVEL_NETSEND] [INT] NOT NULL,
	[NOTIFY_LEVEL_PAGE] [INT] NOT NULL,
	[NOTIFY_EMAIL_OPERATOR_ID] [INT] NOT NULL,
	[NOTIFY_NETSEND_OPERATOR_ID] [INT] NOT NULL,
	[NOTIFY_PAGE_OPERATOR_ID] [INT] NOT NULL,
	[DELETE_LEVEL] [INT] NOT NULL,
	[DATE_CREATED] [DATETIME] NOT NULL,
	[DATE_MODIFIED] [DATETIME] NOT NULL,
	[VERSION_NUMBER] [INT] NOT NULL)

IF OBJECT_ID(N'TEMPDB..#INSERTED') IS NOT NULL
BEGIN
DROP TABLE #INSERTED
END

CREATE TABLE #INSERTED  ([JOB_ID] [UNIQUEIDENTIFIER] NOT NULL,
	[ORIGINATING_SERVER_ID] [INT] NOT NULL,
	[NAME] [SYSNAME] NOT NULL,
	[ENABLED] [TINYINT] NOT NULL,
	[DESCRIPTION] [NVARCHAR](512) NULL,
	[START_STEP_ID] [INT] NOT NULL,
	[CATEGORY_ID] [INT] NOT NULL,
	[OWNER_SID] [VARBINARY](85) NOT NULL,
	[NOTIFY_LEVEL_EVENTLOG] [INT] NOT NULL,
	[NOTIFY_LEVEL_EMAIL] [INT] NOT NULL,
	[NOTIFY_LEVEL_NETSEND] [INT] NOT NULL,
	[NOTIFY_LEVEL_PAGE] [INT] NOT NULL,
	[NOTIFY_EMAIL_OPERATOR_ID] [INT] NOT NULL,
	[NOTIFY_NETSEND_OPERATOR_ID] [INT] NOT NULL,
	[NOTIFY_PAGE_OPERATOR_ID] [INT] NOT NULL,
	[DELETE_LEVEL] [INT] NOT NULL,
	[DATE_CREATED] [DATETIME] NOT NULL,
	[DATE_MODIFIED] [DATETIME] NOT NULL,
	[VERSION_NUMBER] [INT] NOT NULL)

	INSERT INTO #INSERTED SELECT * FROM DELETED
	INSERT INTO #DELETED SELECT * FROM INSERTED

DECLARE @ACTION AS CHAR(1);
DECLARE @JOB_NAME VARCHAR(1024);
DECLARE @AFFECTED_COL NVARCHAR(100);
DECLARE @COLUMN VARCHAR(100)
DECLARE @CMD NVARCHAR (500)
DECLARE @COLAFF TABLE (COLUMNS NVARCHAR(200))

    SET @ACTION = 'I'; -- SET ACTION TO INSERT BY DEFAULT.
	SELECT @JOB_NAME =NAME FROM INSERTED
    IF EXISTS(SELECT * FROM DELETED)
    BEGIN
        SET @ACTION = 
            CASE
                WHEN EXISTS(SELECT * FROM INSERTED) THEN 'U' -- SET ACTION TO UPDATED.
                ELSE 'D' -- SET ACTION TO DELETED.       
            END
		SELECT @JOB_NAME =NAME FROM DELETED
    END
    ELSE
	BEGIN
        IF NOT EXISTS(SELECT * FROM INSERTED) RETURN; -- NOTHING UPDATED OR INSERTED.
		END


DECLARE DB_CURSOR CURSOR FOR 
		SELECT NAME FROM SYS.COLUMNS WHERE OBJECT_ID =  OBJECT_ID('SYSJOBS')
	OPEN DB_CURSOR  
	FETCH NEXT FROM DB_CURSOR INTO @COLUMN  
		WHILE @@FETCH_STATUS = 0  
		BEGIN  
				SET @CMD = 		'DECLARE @AFFECTEDCOL VARCHAR (100)
				IF ((SELECT  '+@COLUMN +' FROM #DELETED  ) <> (SELECT '+@COLUMN+' FROM #INSERTED))
				BEGIN
				SELECT '''+@COLUMN+'''
		
				END'
		
		
				INSERT INTO @COLAFF  EXEC (@CMD)
				--PRINT @CMD

			  FETCH NEXT FROM DB_CURSOR INTO @COLUMN 
		END 

	CLOSE DB_CURSOR  
	DEALLOCATE DB_CURSOR 



SET @AFFECTED_COL = (SELECT   DISTINCT
							SUBSTRING(
										(
											SELECT ','+COLUMNS  
											FROM @COLAFF   
            
											FOR XML PATH ('')
										), 2, 1000) [COLUMNS]
					 FROM @COLAFF)
 
INSERT INTO DBO.DELJOB (LOGINNAME, JOB_NAME, DATE_TIME, ACTION, AFFECTED_COL) 
			VALUES (ORIGINAL_LOGIN(), @JOB_NAME, GETDATE() ,@ACTION, @AFFECTED_COL) 

DROP TABLE #DELETED
DROP TABLE #INSERTED

END