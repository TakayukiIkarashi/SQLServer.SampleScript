/*
********************************************************************************
 関数名：fn_create_hash
 概要　：引数に指定された文字列を暗号化して返します
 引数　：[@str]      ...暗号化する文字列
 　　　　[@salt_base]...saltのベースとなる文字列
 戻り値：暗号化したバイナリ
********************************************************************************
*/
CREATE FUNCTION fn_create_hash
(
    @str VARCHAR(20)        --暗号化する文字列
  , @salt_base VARCHAR(20)  --saltのベースとなる文字列
)
RETURNS VARCHAR(100)
AS
BEGIN

    /*
    暗号化アルゴリズムは、次のとおりです。

        MD5([暗号化する文字列] + salt)

    MD5()は、引数の文字列をMD5形式で暗号化する関数とします。
    salt値は、次のロジックで生成します。

        HASHBYTES('MD5', 'gihyo')
    */

    RETURN UPPER(
        master.dbo.fn_varbintohexstr(
            HASHBYTES(
                'MD5', @str + UPPER(
                    master.dbo.fn_varbintohexstr(
                        HASHBYTES('MD5', @salt_base)
                    )
                )
            )
        )
    );
END
GO
