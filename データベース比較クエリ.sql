DECLARE @db1_name VARCHAR(40);
DECLARE @db2_name VARCHAR(40);

--------------------------------------------------------------------------------
-- 比較するデータベースを指定します
SET @db1_name = 'foo';  -- データベース1
SET @db2_name = 'bar';  -- データベース2
--------------------------------------------------------------------------------

IF (OBJECT_ID('[tempDB]..[#db1_tables]', 'U') IS NOT NULL)
BEGIN
    DROP TABLE #db1_tables;
END
CREATE TABLE #db1_tables (
    table_name VARCHAR(40)
);
IF (OBJECT_ID('[tempDB]..[#db2_tables]', 'U') IS NOT NULL)
BEGIN
    DROP TABLE #db2_tables;
END
CREATE TABLE #db2_tables (
    table_name VARCHAR(40)
);
IF (OBJECT_ID('[tempDB]..[#unmatched_tables]', 'U') IS NOT NULL)
BEGIN
    DROP TABLE #unmatched_tables;
END
CREATE TABLE #unmatched_tables (
    table_name VARCHAR(40)
);

DECLARE @sql VARCHAR(8000);

SET @sql = '';
SET @sql += ' INSERT INTO #db1_tables ';
SET @sql += ' SELECT name FROM ' + @db1_name + '.sys.tables WHERE lob_data_space_id = 0;';
EXEC (@sql);

SET @sql = '';
SET @sql += ' INSERT INTO #db2_tables ';
SET @sql += ' SELECT name FROM ' + @db2_name + '.sys.tables WHERE lob_data_space_id = 0;';
EXEC (@sql);

-- データベース1にあってデータベース2にないものを列挙します
SELECT table_name AS [db1にあってdb2にないテーブル] FROM (
    SELECT table_name FROM #db1_tables
    EXCEPT
    SELECT table_name FROM #db2_tables
) AS m
ORDER BY table_name;

-- データベース2にあってデータベース1にないものを列挙します
SELECT table_name AS [db2にあってdb1にないテーブル] FROM (
    SELECT table_name FROM #db2_tables
    EXCEPT
    SELECT table_name FROM #db1_tables
) AS m
ORDER BY table_name;

DECLARE cur_tables CURSOR FOR
SELECT table_name FROM #db1_tables
UNION
SELECT table_name FROM #db2_tables;

DECLARE @table_name VARCHAR(40);

OPEN cur_tables;

FETCH NEXT FROM cur_tables INTO @table_name;

WHILE (@@FETCH_STATUS = 0)
BEGIN
    IF (
        (EXISTS(SELECT 'X' FROM #db1_tables WHERE table_name = @table_name)) AND
        (EXISTS(SELECT 'X' FROM #db2_tables WHERE table_name = @table_name))
    )
    BEGIN
        PRINT @table_name;

        SET @sql = '';
        SET @sql += ' IF (EXISTS(';
        SET @sql += '     (';
        SET @sql += '         SELECT * FROM ' + @db1_name + '.dbo.' + @table_name;
        SET @sql += '         UNION';
        SET @sql += '         SELECT * FROM ' + @db2_name + '.dbo.' + @table_name;
        SET @sql += '     )';
        SET @sql += '     EXCEPT';
        SET @sql += '     (';
        SET @sql += '         SELECT * FROM ' + @db1_name + '.dbo.' + @table_name;
        SET @sql += '         INTERSECT';
        SET @sql += '         SELECT * FROM ' + @db2_name + '.dbo.' + @table_name;
        SET @sql += '     )';
        SET @sql += ' ))';
        SET @sql += ' BEGIN';
        SET @sql += '     INSERT INTO #unmatched_tables (table_name) VALUES (''' + @table_name + ''');';
        SET @sql += ' END';
        EXEC (@sql);
    END
    FETCH NEXT FROM cur_tables INTO @table_name;
END

CLOSE cur_tables;
DEALLOCATE cur_tables;

-- データベース1とデータベース2の内容が不一致なものを列挙します
SELECT table_name AS [db1とdb2が不一致なテーブル] FROM #unmatched_tables
ORDER BY table_name;
