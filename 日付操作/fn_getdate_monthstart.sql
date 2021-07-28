--概要　　　：月初を返します
--引数　　　：[@year] …対象年
--　　　　　　[@month]…対象月
--戻り値　　：対象年月の月初
IF (EXISTS(SELECT * FROM sys.objects WHERE (type = 'FN') AND (name = 'fn_getdate_monthstart')))
BEGIN
    DROP FUNCTION fn_getdate_monthstart;
END
GO

CREATE FUNCTION fn_getdate_monthstart
(
    @year INT,
    @month INT
)
RETURNS DATETIME
AS
BEGIN
    DECLARE @strdate VARCHAR(10);
    SET @strdate = CONVERT(VARCHAR, @year) + '-' + CONVERT(VARCHAR, @month) + '-1';

    RETURN CONVERT(DATETIME, @strdate);
END
GO
