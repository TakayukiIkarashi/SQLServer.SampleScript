/*
********************************************************************************
概要　：同一システムのデータベースにおいて、あるデータベースから別のデータベース
　　　　へすべてのテーブルの内容をコピーします。
--------------------------------------------------------------------------------
詳細　：コピー元のテーブルの列名から、列名指定のINSERT命令を作成し、コピー先の同
　　　　一テーブルに追加します。
　　　　同一システムの旧バージョンのデータベースから、現行バージョンのデータベー
　　　　スへのデータ移行の際に使用します。現行バージョンでは新たなカラムが追
　　　　加されている場合にも対応できます。
--------------------------------------------------------------------------------
使い方：コピー元データベース名が[db1]、コピー先データベース名が[db2]となっていま
　　　　す。これを、それぞれ環境に合わせた正しいデータベース名に置換し、このSQL
　　　　を実行します。db1データベースのすべてのテーブルの内容が、db2データベース
　　　　のテーブルにコピーされます。旧バージョンにのみ存在するテーブルが存在する
　　　　場合、当該テーブルのみエラーとなり、PRINT文でテーブル名を出力します。
********************************************************************************
*/

--結果件数を表示しない
SET NOCOUNT ON;

/*
==============================
変数定義
==============================
*/
--SQL組み立て文字列
DECLARE @sql VARCHAR(MAX);

/*
==============================
処理部
==============================
*/
--表名カーソル作成
DECLARE [cur表名] CURSOR FOR
SELECT [object_id], [name] FROM [db1].[sys].[tables];

--表名カーソル内で使用する変数の宣言
DECLARE @object_id INT;
SET @object_id = -1;
DECLARE @表名 VARCHAR(100);
SET @表名 = '';

--表名カーソルを開き、1件目のデータを取得
OPEN [cur表名];
FETCH [cur表名] INTO @object_id, @表名;

--表名カーソルのデータを1件ずつ処理
WHILE (@@fetch_status = 0)
BEGIN
    --表のobject_idに該当する列名カーソル作成
    DECLARE [cur列名] CURSOR FOR
    SELECT [name], [is_identity] FROM [db1].[sys].[columns]
    WHERE [object_id] = @object_id;

    --列名カーソル内で使用する変数の宣言
    DECLARE @列名 VARCHAR(100);
    SET @列名 = '';
    DECLARE @is_identity INT;
    SET @is_identity = 0;

    --AUTO_INCREMENT列かどうかを判断するフラグ変数を宣言
    DECLARE @自動付番 INT;
    SET @自動付番 = 0;

    --INSERT命令を作成する際の列名を列挙する変数を宣言
    DECLARE @列名列挙 VARCHAR(MAX);
    SET @列名列挙 = '';

    --列名カーソルを開き、1件目のデータを取得
    OPEN [cur列名];
    FETCH [cur列名] INTO @列名, @is_identity;

    --列名カーソルのデータを1件ずつ処理
    WHILE (@@fetch_status = 0)
    BEGIN
        --AUTO_INCREMENT列の場合
        IF (@is_identity = 1) AND (@自動付番 = 0)
        BEGIN
            --AUTO_INCREMENT列フラグをON
            SET @自動付番 = 1;
        END

        --2件め以降に列名を追加するならカンマを追記
        IF (@列名列挙 <> '')
        BEGIN
            SET @列名列挙 = @列名列挙 + ',';
        END

        --列名を追記
        SET @列名列挙 = @列名列挙 + @列名;

        --次のレコードへ
        FETCH [cur列名] INTO @列名, @is_identity;
    END

    --列名カーソルを閉じ、解放
    CLOSE [cur列名];
    DEALLOCATE [cur列名];

    --エラーフラグ変数を宣言し、初期値を代入
    DECLARE @err有 INT;
    SET @err有 = 0;

    --コピー先のテーブルからデータを全削除
    SET @sql = '';
    SET @sql = @sql + ' DELETE FROM [db2].[dbo].[' + @表名 + '];';
    BEGIN TRY
        EXECUTE (@sql);
    END TRY
    BEGIN CATCH
        --削除に失敗したら失敗した表名を表示し、エラーフラグをON
        PRINT @sql;
        SET @err有 = 1;
        PRINT '削除失敗：' + @表名;
    END CATCH

    --列名の列挙を終えたらテーブルをコピー
    IF (@err有 = 0)
    BEGIN
        SET @sql = '';

        --AUTO_INCREMENT列の場合
        IF (@自動付番 = 1)
        BEGIN
            --IDENTITY_INSERTを解除
            SET @sql = @sql + 'SET IDENTITY_INSERT [db2].[dbo].[' + @表名 + '] ON;';
        END

        --INSERTコマンド
        SET @sql = @sql + ' INSERT INTO [db2].[dbo].[' + @表名 + '] (' + @列名列挙 + ')';
        SET @sql = @sql + ' SELECT ' + @列名列挙
        SET @sql = @sql + ' FROM [db1].[dbo].[' + @表名 + '];';

        --AUTO_INCREMENT列の場合
        IF (@自動付番 = 1)
        BEGIN
            --IDENTITY_INSERTを再開
            SET @sql = @sql + 'SET IDENTITY_INSERT [db2].[dbo].[' + @表名 + '] OFF;';
        END

        --レコードを1件追加
        BEGIN TRY
            EXECUTE (@sql);
        END TRY
        BEGIN CATCH
            --追加に失敗したら失敗した表名を表示
            PRINT @sql;
            PRINT '追加失敗：' + @表名;
        END CATCH
    END

    --次のレコードへ
    FETCH [cur表名] INTO @object_id, @表名;
END

--表名カーソルを閉じ、解放
CLOSE [cur表名];
DEALLOCATE [cur表名];
