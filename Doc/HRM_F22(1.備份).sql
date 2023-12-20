--執行：備份
USE [master]
GO

--*******設定檔
--1.bak及mdf路徑
declare @DBFolder varchar(1000) -- 目錄絕對路徑
set @DBFolder= 'D:\TempDB'
--2.資料來源(資料庫名稱)
DECLARE @from_table nvarchar(max);   --資料來源
Set @from_table = 'FTIS-T8-DataSupply'
--3.資料備份(資料庫名稱)
DECLARE @to_table nvarchar(max);   --資料備份
Set @to_table = 'Temp_FTIS-T8-DataSupply'
--*************

--maptable
--Select * From [dbo].[maptable] Where ftable = 'F22cmmHoliday'

--一、資料夾目錄
-----利用T-SQL  判斷資料夾路徑是否存在，不存在新增
DECLARE @qry nvarchar(max);

-- 開啟 xp_cmdshell
EXEC sp_configure 'show advanced options', 1; RECONFIGURE; EXEC sp_configure 'xp_cmdshell', 1; RECONFIGURE;

Set @qry = '

declare @result INT  -- 回傳結果

-- 執行檢查目錄
EXEC @result = master.dbo.xp_cmdshell ''dir ' + @DBFolder + ''', NO_OUTPUT

IF (@result) <> 0
-- 不存在則新增
BEGIN
    -- 執行新增目錄
    EXEC master.dbo.xp_cmdshell ''mkdir ' + @DBFolder + ''', NO_OUTPUT
END   
'

--Print @qry
EXEC sp_executesql @qry

-- 關閉 xp_cmdshell
EXEC sp_configure 'show advanced options', 1; RECONFIGURE; EXEC sp_configure 'xp_cmdshell', 0; RECONFIGURE;

---------------------------------
--二、備份與還原
Set @qry = ''

--(1)備份(Copy覆蓋原有的bak) (WITH =>INIT 覆蓋)
Set @qry = 'BACKUP DATABASE [' + @from_table + '] TO  
	DISK = N''' + @DBFolder + '\' + @to_table + '.bak'' 
	WITH NOFORMAT, INIT,  NAME = N''' + @from_table + '-完整備份'', 
	SKIP, NOREWIND, NOUNLOAD,  STATS = 10

'

--(2)還原
--注意錯誤：備份組包含現有的 'Temp_FTIS-T8' 資料庫以外的資料庫備份

Set @qry = @qry + 'RESTORE DATABASE [' + @to_table + ']
FROM  DISK = N''' + @DBFolder + '\' + @to_table + '.bak'' 
WITH  FILE = 1,  
MOVE N''' + @from_table + ''' TO N''' + @DBFolder + '\' + @to_table + '.mdf'',  
MOVE N''' + @from_table + '_log'' TO N''' + @DBFolder + '\' + @to_table + '_log.ldf'',  
NOUNLOAD,  REPLACE,  STATS = 5
'

--Print @qry
Exec sp_executesql @qry


---------------------------------
--三.產生差異比對(新舊資料表 fmaptable)
Set @qry = ''

Set @qry = 'USE [' + @to_table + ']'

Set @qry = @qry + '
	IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N''[dbo].[maptable]'') AND type in (N''U''))
		DROP TABLE [dbo].[maptable]

	CREATE TABLE [dbo].[maptable](
		[ftable] [nvarchar](50) NOT NULL,
		[fcolumn] [nvarchar](50) NOT NULL,
		[tcolumn] [nvarchar](50) NOT NULL,
	 CONSTRAINT [PK_maptable] PRIMARY KEY CLUSTERED 
	(
		[ftable] ASC,
		[fcolumn] ASC,
		[tcolumn] ASC
	)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
	) ON [PRIMARY]

	DECLARE @p1 AS nvarchar(128)
	DECLARE vend_cursor CURSOR
	FOR SELECT TABLE_NAME 
	FROM INFORMATION_SCHEMA.TABLES
	Where table_name <> ''maptable''
	OPEN vend_cursor

	FETCH NEXT FROM vend_cursor
		INTO @p1
	while @@fetch_status = 0 
	BEGIN	            
	
		--取得所有欄位
		Insert Into maptable(ftable, fcolumn, tcolumn)
		select TABLE_NAME AS ftable, COLUMN_NAME AS fcolumn, COLUMN_NAME AS tcolumn
		from INFORMATION_SCHEMA.COLUMNS 
		where table_name=@p1

		FETCH NEXT FROM vend_cursor INTO @p1
	END

	CLOSE vend_cursor
	DEALLOCATE vend_cursor
'

--Print @qry
Exec sp_executesql @qry

-----------------------------------------------------

