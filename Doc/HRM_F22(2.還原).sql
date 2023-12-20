--執行：還原
USE [FTIS-T8-DataSupply]

--*******設定檔
--1.資料來源(資料庫名稱)
DECLARE @from_table nvarchar(max);   --資料來源
Set @from_table = 'Temp_FTIS-T8-DataSupply'
--*************

DECLARE @qry nvarchar(max);
DECLARE @p1 AS nvarchar(128)



----------------------------------
--設定差異比對資料表(新增)
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[maptable]') AND type in (N'U'))
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


Set @qry = ''
--取得temp內，將更新的資料表和欄位
Set @qry = @qry + 'Insert Into [dbo].[maptable]
					Select * 
					From [' + @from_table + '].[dbo].[maptable]
				  '

--****設定不同步的資料表****
Set @qry = @qry + 'Where ftable Not In (''ZpvRepair'', ''ZpvPass'', ''ZpvAttendClass'', ''ZpvClassHours'')'

--Print @qry
EXECUTE sp_executesql @qry

----------------------------------

-- 全部停用：資料庫內的 FOREIGN KEY 條件約束(CONSTRAINT)、CHECK 條件約束(CONSTRAINT)
EXEC sp_MSforeachtable @command1='ALTER TABLE ? NOCHECK CONSTRAINT ALL'

--一、清空所有資料表
-------------------------------------------------
Set @qry = ''

DECLARE vend_cursor CURSOR
FOR SELECT TABLE_NAME 
FROM INFORMATION_SCHEMA.TABLES
Where TABLE_NAME <> 'maptable'
And TABLE_NAME <> '__MigrationHistory'  --***Model異動表不可更新***
OPEN vend_cursor

FETCH NEXT FROM vend_cursor
    INTO @p1
while @@fetch_status = 0 
BEGIN	            
	
	SET @qry = @qry + 'Delete ' + '[dbo].' + '[' + @p1 + '] '
	
    FETCH NEXT FROM vend_cursor INTO @p1
END

CLOSE vend_cursor
DEALLOCATE vend_cursor


--Print @qry
EXECUTE sp_executesql @qry

---------------------------------
--二、資料還原
Set @qry = ''

DECLARE @to_table nvarchar(100)    
Declare @tcolumns AS Nvarchar(max)
Declare @fcolumns AS Nvarchar(max)

DECLARE vend_cursor CURSOR
FOR SELECT TABLE_NAME 
FROM INFORMATION_SCHEMA.TABLES
Where TABLE_NAME <> 'maptable'
And TABLE_NAME <> '__MigrationHistory'  --***Model異動表不可更新***
OPEN vend_cursor

FETCH NEXT FROM vend_cursor
    INTO @p1
while @@fetch_status = 0 
BEGIN
	
	--取得所有欄位
	Set @fcolumns = null
	Set @tcolumns = null 

	Select @fcolumns = case when @fcolumns is null then '[' + fcolumn + ']' else @fcolumns + ',' + '[' + fcolumn + ']' end,
		   @tcolumns = case when @tcolumns is null then '[' + tcolumn + ']' else @tcolumns + ',' + '[' + tcolumn + ']' end
	From [dbo].[maptable] a
	Where ftable = @p1
	And Exists
	(
		Select * From INFORMATION_SCHEMA.COLUMNS b
		where table_name=@p1
		And a.tcolumn = b.COLUMN_NAME
	)

	IF(@fcolumns Is Not Null And @tcolumns Is Not Null)   --(開始)有對應temp資料
	Begin
		Set @to_table = '[dbo].' + '[' + @p1 + ']'

		--id是否有設定自動+1(IDENTITY_INSERT)
		IF EXISTS(select * from INFORMATION_SCHEMA.COLUMNS 
			where table_name=@p1
			And columnproperty(object_id(@p1), column_name,'IsIdentity')=1)	   
		Begin

			--將自動+1欄位，從0算起
			SET @qry = @qry + 'DBCC CHECKIDENT(''' + @to_table + ''', RESEED, 0)
						'
			SET @qry = @qry + 'set identity_insert ' + @to_table + ' ON;
						' 
		End

		SET @qry = @qry + ' Insert Into ' + @to_table + '(' + @tcolumns + ')' + '
							Select ' + @fcolumns + '
							From ' + '[' + @from_table + '].[dbo].' + '[' + @p1 + ']
							'
		IF EXISTS(select * from INFORMATION_SCHEMA.COLUMNS 
			where table_name=@p1
			And columnproperty(object_id(@p1), column_name,'IsIdentity')=1)	   
		Begin
			SET @qry = @qry + 'set identity_insert ' + @to_table + ' OFF;
			' 
		End

		--qry太長，逐步執行(資料表過多)
		--Print @qry
		EXECUTE sp_executesql @qry

		Set @qry = ''
	End

	FETCH NEXT FROM vend_cursor INTO @p1
END

CLOSE vend_cursor
DEALLOCATE vend_cursor


-- 全部啟用：資料庫內的 FOREIGN KEY 條件約束(CONSTRAINT)、CHECK 條件約束(CONSTRAINT)
EXEC sp_MSforeachtable @command1='ALTER TABLE ? WITH NOCHECK CHECK CONSTRAINT ALL'
-----------------------------------------------------------------------------------------

----------------------------------
--設定差異比對資料表(移除)
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[maptable]') AND type in (N'U'))
		DROP TABLE [dbo].[maptable]
