--概要　　　：任意の年月の最終営業日を求めます。
--引数　　　：[@year] …対象年
--　　　　　　[@month]…対象月
--戻り値　　：任意の年月の第n日曜日の日付
IF (EXISTS(SELECT * FROM sys.objects WHERE (type = 'FN') AND (name = 'fn最終営業日')))
BEGIN
    DROP FUNCTION fn最終営業日;
END
GO

CREATE FUNCTION fn最終営業日
(
    @yyyy INT,
    @mm INT
)
RETURNS DATETIME
AS
BEGIN
    --月末日付を取得します。
    DECLARE @月末日付 DATETIME;
    SET @月末日付 = dbo.fn_getdate_monthend(@yyyy, @mm);

    --戻り値用日付を定義します。
    DECLARE @r日付 DATETIME;
    SET @r日付 = @月末日付;

    --休日テーブルに休日として登録されていない日付を求めます。
    WHILE (0 = 0)
    BEGIN
        IF (NOT EXISTS(SELECT * FROM [休日] WHERE [日付] = @r日付))
        BEGIN
            BREAK;
        END
        SET @r日付 = DATEADD(d, -1, @r日付);
    END

    RETURN @r日付;
END
GO
