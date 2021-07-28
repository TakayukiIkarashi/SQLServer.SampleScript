--概要　　　：引数に指定された年の春分の日を返します
--引数　　　：[@yyyy]…対象年
--戻り値　　：引数に指定された年の春分の日
IF (EXISTS(SELECT * FROM sys.objects WHERE (type = 'FN') AND (name = 'fn_getdate_monthend')))
BEGIN
    DROP FUNCTION fn_getdate_monthend;
END
GO

CREATE FUNCTION fn_getdate_monthend
(
    @year INT
  , @month INT
)
RETURNS DATETIME
AS
BEGIN
    --対象月の月初の文字列型を求めます
    DECLARE @strdate_start VARCHAR(10);
    SET @strdate_start = CONVERT(VARCHAR, @year) + '-' + CONVERT(VARCHAR, @month) + '-1';

    --対象月の月初の日付型を求めます
    DECLARE @date_month_start DATETIME;
    SET @date_month_start = CONVERT(DATETIME, @strdate_start);

    --対象月の月初の1カ月後を求めます
    DECLARE @date_month_plus1 DATETIME;
    SET @date_month_plus1 = DATEADD(month, 1, @date_month_start);

    --対象月の月初の1カ月後の前日を返します
    RETURN DATEADD(day, -1, @date_month_plus1);
END
GO
