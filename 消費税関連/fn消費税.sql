--概要　　　：指定した日付の期間に該当する税抜金額から消費税額を返します。
--引数　　　：[@税抜金額]
--　　　　　　[@対象日付]
--戻り値　　：税抜金額
IF (EXISTS(SELECT * FROM sysobjects WHERE (type = 'FN') AND (name = 'fn消費税')))
BEGIN
    DROP FUNCTION fn消費税;
END
GO

CREATE FUNCTION fn消費税
(
    @税抜金額 MONEY,
    @対象日付 DATETIME
)
RETURNS MONEY
AS
BEGIN
    --対象日付に該当する消費税率を消費税率マスタから取得します。
    DECLARE @消費税率 INT;
    SET @消費税率 = 0;
    SELECT @消費税率 = [税率] FROM [消費税率]
    WHERE @対象日付 BETWEEN [開始日付] AND [終了日付];

    --取得した消費税率と税抜金額を乗算して消費税額を求めます。
    RETURN ROUND(@税抜金額 * @消費税率 / 100, 0, 1);
END
GO
