--概要　　　：指定したspidのプロセスを削除します。
--引数　　　：[@target_spid]...プロセスを削除するspid
--戻り値　　：正常終了なら0、そうでなければ-1
--結果セット：例外が発生した場合、エラー情報
CREATE PROCEDURE [sp_kill_process]
    @target_spid SMALLINT
AS
BEGIN
    SET NOCOUNT ON;

    --KILLコマンドを実行する動的SQLを作成します。
    --※「KILL @spid;」はエラーとなるため、動的SQLで対応
    DECLARE @sql VARCHAR(MAX);
    SET @sql = 'KILL ' + CONVERT(VARCHAR, @target_spid);

    --spidを削除します。
    BEGIN TRY
        EXECUTE (@sql);
    END TRY
    BEGIN CATCH
        EXECUTE sp_returnerror 'sp_kill_process:プロセスの削除に失敗しました。';
        RETURN (-1);
    END CATCH

    --正常終了を返します。
    RETURN (0);
END
GO
