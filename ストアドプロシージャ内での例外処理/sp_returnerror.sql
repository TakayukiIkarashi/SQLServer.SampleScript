--概要　　　：エラーの詳細を返します
--引数　　　：[@message]…エラー詳細に含める任意の文字列
--戻り値　　：なし
--結果セット：エラー情報
CREATE PROCEDURE [dbo].[sp_returnerror]
    @message VARCHAR(255)   --エラー詳細に含める任意の文字列
AS
BEGIN
   --結果件数を表示しないようにします。
    SET NOCOUNT ON;

    --各種エラー関数から結果セットに含めるレコードを生成します。
    SELECT
        ERROR_NUMBER() AS ErrorNumber,
        ERROR_SEVERITY() AS ErrorSeverity,
        ERROR_STATE() AS ErrorState,
        ERROR_PROCEDURE() AS ErrorProcedure,
        ERROR_LINE() AS ErrorLine,
        ERROR_MESSAGE() AS ErrorMessage,
        @message AS ApplicationMessage;
END
