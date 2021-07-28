--概要　　　：引数に指定された年の秋分の日を返します
--引数　　　：[@yyyy]…対象年
--戻り値　　：引数に指定された年の秋分の日
IF (EXISTS(SELECT * FROM sys.objects WHERE (type = 'FN') AND (name = 'fn_getdate_syuubun')))
BEGIN
    DROP FUNCTION fn_getdate_syuubun;
END
GO

CREATE FUNCTION fn_getdate_syuubun
(
    @year INT
)
RETURNS INT
AS
BEGIN
    RETURN (CONVERT(INT, ((23.2488 + 0.242194 * (@year - 1980)) - (CONVERT(INT, ((@year - 1980) / 4.000000))))));
END
GO
