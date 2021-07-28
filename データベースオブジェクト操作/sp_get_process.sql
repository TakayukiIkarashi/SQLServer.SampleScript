--概要　　　：指定したホスト名が生成したSQL Serverのプロセスを取得し、spidを返します。
--引数　　　：[@hostname]...ホスト名
--戻り値　　：正常終了なら0、そうでなければ-1
--結果セット：正常終了した場合、指定したホストのspid結果リスト
--　　　　　　例外が発生した場合、エラー情報
CREATE PROCEDURE [sp_get_process]
    @hostname VARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;

    --spidを取得します。
    BEGIN TRY
        SELECT [spid]
        FROM [sys].[sysprocesses] WITH (nolock)
        WHERE [hostname] = @hostname
        AND [spid] <> @@spid
        ORDER BY [spid];
    END TRY
    BEGIN CATCH
        EXECUTE sp_returnerror 'sp_get_process:プロセスの取得に失敗しました。';
        RETURN (-1);
    END CATCH

    --正常終了を返します。
    RETURN (0);
END
GO
