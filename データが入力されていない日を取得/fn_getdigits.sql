/*
********************************************************************************
 概要：数値の開始と終了を指定し、該当する範囲の数列の集合をテーブルとして返す
********************************************************************************
*/
CREATE FUNCTION fn_getdigits
(
    @start  INT     --開始数値
  , @end    INT     --終了数値
)
RETURNS @table TABLE
(
    num_value   INT
)
AS
BEGIN
    --数値を格納する変数を定義し、初期値として開始数値を代入します
    DECLARE @i INT;
    SET @i = @start;

    --終了数値になるまで処理を繰り返します
    WHILE (@i <= @end)
    BEGIN
        --テンポラリテーブルに現在の数値データを追加します
        INSERT INTO @table (num_value) VALUES (@i);

        --次の数値を準備します
        SET @i = @i + 1;
    END

    --終了します
    RETURN;
END
