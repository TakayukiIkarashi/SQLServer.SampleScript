--概要　　　：任意の年月の第ｎ回目の〇曜日の日付を求めます
--引数　　　：[@year] 　　…対象年
--　　　　　　[@month]　　…対象月
--　　　　　　[@num]　　　…何番目の曜日か、第1曜日なら1、第3曜日なら3
--　　　　　　[@dayofweek]…1（日曜日）から7（土曜日）までの数字を返します
--戻り値　　：任意の年月の第n日曜日の日付
IF (EXISTS(SELECT * FROM sys.objects WHERE (type = 'FN') AND (name = 'fn_getdate_dayofweek')))
BEGIN
    DROP FUNCTION fn_getdate_dayofweek;
END
GO

CREATE FUNCTION fn_getdate_dayofweek
(
    @yyyy INT
  , @mm INT
  , @num INT
  , @dayofweek INT
)
RETURNS DATETIME
AS
BEGIN
    --年および月の文字列型を変数に取得します
    DECLARE @str_yyyy VARCHAR(4);
    SET @str_yyyy = CONVERT(VARCHAR, @yyyy);

    DECLARE @str_mm VARCHAR(2);
    SET @str_mm = RIGHT('00' + CONVERT(VARCHAR, @mm), 2);

    --日付データの作業用変数です
    DECLARE @date_work DATETIME;
    SET @date_work = NULL;

    --指定した年月の1日の曜日を取得します
    SET @date_work = CONVERT(DATETIME, @str_yyyy + '-' + @str_mm + '-01');
    DECLARE @dayofweek_first INT;
    SET @dayofweek_first = DATEPART(weekday, @date_work);

    --指定した曜日の第1曜日の日を求めます
    DECLARE @int_dd INT;
    SET @int_dd = @dayofweek - @dayofweek_first + 1;
    IF (@int_dd <= 0)
    BEGIN
        SET @int_dd = @int_dd + 7;
    END

    --求めた日を日付型に変換します。
    SET @date_work = CONVERT(DATETIME, @str_yyyy + '-' + @str_mm + '-' + CONVERT(VARCHAR, @int_dd));

    --指定した回数分、週を移動します。
    RETURN (DATEADD(day, (@num - 1) * 7, @date_work));
END
GO
