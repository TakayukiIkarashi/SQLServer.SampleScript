--概要　　　：引数に指定された年の春分の日を返します
--引数　　　：[@yyyy]…対象年
--戻り値　　：引数に指定された年の春分の日
IF (EXISTS(SELECT * FROM sys.objects WHERE (type = 'FN') AND (name = 'fn_getdate_syunbun')))
BEGIN
    DROP FUNCTION fn_getdate_syunbun;
END
GO

CREATE FUNCTION fn_getdate_syunbun
(
    @year INT
)
RETURNS INT
AS
BEGIN
    RETURN (CONVERT(INT, ((20.8431 + 0.242194 * (@year - 1980)) - ((@year - 1980) / 4.000000))));
END
GO
