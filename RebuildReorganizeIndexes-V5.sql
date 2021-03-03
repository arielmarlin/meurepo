/*
	This script automates the index maintenance for the ePO database and is based on example D from this Microsoft
	article:
	
	https://msdn.microsoft.com/en-us/library/ms188917.aspx
	
	In the script below the items labelled as 'EDITABLE' can be changed based on your environment.
	
	If the fragmentation is between 20% and 30% the index will be reorganized and if the fragmentation is greater 
	than 30% it will be rebuilt.
	
	Instructions:
		1. Edit the line below that starts with "USE [ePO database name here]" and change the value between the 
		brackets to the name of your ePO database.
		2. Configure the script to run as a SQL Agent Job. Use the following link as a reference: 
		https://msdn.microsoft.com/en-us/library/ms190268.aspx
	
	Notes:
		- For indexes that do not have ALLOW_PAGE_LOCKS set to ON this script will temporarily enable that option so
		that the maintenance can be done, and then it will revert back to the previous setting for ALLOW_PAGE_LOCKS.
		
		- Ensure a USE <databasename> statement has been executed first or the ePO database is selected. If this script
		is used in a SQL Agent job then be sure to select the ePO database on the "Job Step" page.
		
		- Index fragmentation is caused by updates/inserts/deletes to table data, which is a normal activity. Due to this
		it is expected that indexes become fragmented on a regular basis. Fragmented indexes do not always lead to
		slower application performance, but it is recommended to run a SQL Agent job that includes this script logic 
		and schedule it to run at least weekly during a designated maintenance window.

*/
USE [ePO database name here]  -- EDITABLE: change to valid ePO database name

SET NOCOUNT ON;

IF ((OBJECT_ID('dbo.EPOServerInfo') IS NULL) AND (Object_ID('dbo.EPOEvents')) IS NULL) -- make sure we are in the ePO database
BEGIN;
	RAISERROR('Ensure the ePO database is selected or that this script begins with:   USE [ePO database name here]', 16, 1)
	RETURN;
END;

DECLARE 
	@schemaname				NVARCHAR(130)
,	@objectname				NVARCHAR(130)
,	@indexname				NVARCHAR(130) 
,	@frag					FLOAT
,	@command				NVARCHAR(4000) 
,	@pagelock				INTEGER
,	@MinPageCount			INTEGER
,	@UpperBoundFragReorg	FLOAT
,	@LowerBoundFrag			FLOAT

SELECT
	@MinPageCount			= 100 	-- EDITABLE: this can be updated to a higher page count if needed
,	@UpperBoundFragReorg	= 30.0 	-- EDITABLE: this is the industry accepted upper bound for reorganizing an index, but in some environments it may be necessary to change this value
,	@LowerBoundFrag			= 20.0;	-- EDITABLE: this is the industry accepted lower bound for reorganizing an index, but in some environments it may be necessary to change this value

IF @MinPageCount < 0
BEGIN;
	RAISERROR('The variable @MinPageCount must be a non-negative integer value. Please review the variable assignment at the top of this script.', 16, 1)
	RETURN;
END;

IF @LowerBoundFrag >= @UpperBoundFragReorg OR @LowerBoundFrag < 0
BEGIN;
	RAISERROR('The variables @LowerBoundFrag and @UpperBoundFragReorg are non-negative float values. The variable @LowerBoundFrag should be set to a non-negative float value that is less than @UpperBoundFragReorg. Please review the variable assignments at the top of this script.', 16, 1)
	RETURN;
END;

IF OBJECT_ID('tempdb.dbo.#work_to_do') IS NOT NULL
BEGIN;
	DROP TABLE #work_to_do;
END;

SELECT
    QUOTENAME(sch.name)						AS SchemaName
,	QUOTENAME(obj.name)						AS ObjectName
,	QUOTENAME(ind.name)						AS IndexName
,	indexStat.avg_fragmentation_in_percent	AS Frag
,	ind.allow_page_locks					AS AllowPageLocks
INTO 
	#work_to_do
FROM 
	sys.dm_db_index_physical_stats (DB_ID(), NULL, NULL , NULL, 'LIMITED') indexStat
JOIN
	sys.indexes ind 
		ON indexStat.index_id = ind.index_id
		AND indexStat.object_id = ind.object_id
JOIN
	sys.objects obj
		ON obj.object_id = indexStat.object_id
JOIN
	sys.schemas sch
		ON obj.schema_id = sch.schema_id
WHERE 
	avg_fragmentation_in_percent >= @LowerBoundFrag 
AND indexStat.index_id > 0 -- ensure only the b-tree indexes are reorg/rebuilt and not the heap structures  
AND indexStat.page_count > @MinPageCount;         

-- Declare the cursor for the list of indexes to be processed
DECLARE IndexCursor CURSOR LOCAL FOR 
SELECT 
	SchemaName
,	ObjectName
,	IndexName
,	Frag
,	AllowPageLocks
FROM 
	#work_to_do;

OPEN IndexCursor;

WHILE (1=1)
BEGIN;
	FETCH NEXT FROM IndexCursor INTO 
		@schemaname	
	,	@objectname	
	,	@indexname	
	,	@frag
	,	@pagelock;

	IF @@FETCH_STATUS < 0 
	BEGIN;
		BREAK;
	END;
 
	IF @pagelock = 0 AND @frag <= @UpperBoundFragReorg
	BEGIN;
		SET @command = N'ALTER INDEX ' + @indexname + N' ON ' + @schemaname + N'.' + @objectname + N' SET (ALLOW_PAGE_LOCKS = ON);'
		EXEC (@command); -- temporarily allow page locks
	END;

    IF @frag <= @UpperBoundFragReorg -- the query used in the cursor will filter for the @LowerBoundFrag
	BEGIN;
		SET @command = N'ALTER INDEX ' + @indexname + N' ON ' + @schemaname + N'.' + @objectname + N' REORGANIZE; ' +
						N'UPDATE STATISTICS ' +  @schemaname + N'.' + @objectname + N' ' + @indexname +';' -- the reorganize operation does not update any stats
	END;
	ELSE IF @frag > @UpperBoundFragReorg
	BEGIN;
		SET @command = N'ALTER INDEX ' + @indexname + N' ON ' + @schemaname + N'.' + @objectname + N' REBUILD';	-- the rebuild operation will update index stats but not table column stats	
	END;

	EXEC (@command);
	PRINT @command;
        
	IF @pagelock = 0 AND @frag <= @UpperBoundFragReorg
	BEGIN;
		SET @command = N'ALTER INDEX ' + @indexname + N' ON ' + @schemaname + N'.' + @objectname + N' SET (ALLOW_PAGE_LOCKS = OFF);'
		EXEC (@command); -- change the allow_page_locks back
	END;
END;

-- Close and deallocate the cursor. Since it is a local cursor it will be deallocated automatically anyways
CLOSE IndexCursor;
DEALLOCATE IndexCursor;

-- Drop the temporary table.
DROP TABLE #work_to_do;
GO