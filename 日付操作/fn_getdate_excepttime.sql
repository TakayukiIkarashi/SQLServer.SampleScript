--概要　　　：引数に指定された日付型から時刻要素を取り除いて返します。
--引数　　　：[@date]…DATETIME型
--戻り値　　：時刻要素を取り除いたDATETIME型
IF (EXISTS(SELECT * FROM sys.objects WHERE (type = 'FN') AND (name = 'fn_getdate_excepttime')))
BEGIN
    DROP FUNCTION fn_getdate_excepttime;
END
GO

CREATE FUNCTION fn_getdate_excepttime
(
    @date DATETIME
)
RETURNS DATETIME
AS
BEGIN
    RETURN CONVERT(DATETIME, CONVERT(nvarchar, @date, 111), 120);
END
GO
