Option Explicit

'==============================
' 処理部
'==============================
'RewriteAuth_MSSQLServiceクラスをインスタンス化します
Dim ra
Set ra = New RewriteAuth_MSSQLService

'RewriteAuth_MSSQLServiceクラスのExecute()メソッドを実行します
Dim bRet
bRet = ra.Execute()

'完了メッセージを表示します
If (bRet) Then
    MsgBox "SQLServerサービスの実行権限を書き換えました。"
Else
    MsgBox "失敗しました。", vbCritical + vbOkOnly
End If

'クラス名：RewriteAuth_MSSQLService
'概要　　：SQLServerサービスの権限を書き変えるためのクラスです
Class RewriteAuth_MSSQLService

    '================================================================================
    ' プロパティ：SQLServerサービスのアクセス権限をバックアップするパスを取得します
    '================================================================================
    Private Property Get SD_PATH()
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")

        SD_PATH = fso.GetSpecialFolder(2).Path & "\mssqlserver_sd.txt"
    End Property

    '********************************************************************************
    ' メソッド名：Execute()
    ' 概要　　　：権限書換実行
    ' パラメータ：なし
    ' 戻り値　　：正常終了ならTrue、そうでなければFalse
    '********************************************************************************
    Public Function Execute()

        '戻り値を初期化します
        Execute = False

        'SQLServerサービスのアクセス権限リストを取得します
        Dim sAcc
        If (GetSD(sAcc) = False) Then
            Exit Function
        End If

        'すでにPowerUsersグループのアクセス権限が存在する場合は処理を抜けます
        If (IsDone(sAcc)) Then
            Execute = True
            Exit Function
        End If

        'アクセス権限リストを書き変えます
        sAcc = RemakeDACL(sAcc)

        'アクセス権限リストの書き換えが行えなかった場合は処理を抜けます
        If (sAcc = "") Then
            Call MsgBox("アクセス権限リストの形式が想定外のため、処理できませんでした。")
            Exit Function
        End If

        '書き換えたアクセス権限リストをSQLServerサービスに再設定します
        If (SetSD(sAcc) = False) Then
            Exit Function
        End If

        '正常終了を返します
        Execute = True

    End Function

    '********************************************************************************
    ' メソッド名：GetSD()
    ' 概要　　　：SQLServerサービスのアクセス権限を取得します
    ' パラメータ：[sAcc]...アクセス権限リスト
    ' 戻り値　　：正常に取得できればTrue、そうでなければFalse
    '********************************************************************************
    Private Function GetSD(ByRef sAcc)

        '戻り値を初期化します
        GetSD = False

        '変数を初期化します
        sAcc = ""

        '実行するコマンドを作成します
        Dim cmd
        cmd = "sc sdshow MSSQLSERVER > " & SD_PATH

        '作成したコマンドを実行します（処理が完了するまで待機します）
        If (CommandBatShell(cmd, True) = False) Then
            Exit Function
        End If

        'アクセス権限リストの出力に失敗した場合は処理を抜けます
        If (IsFileExists(SD_PATH) = False) Then
            Call MsgBox("SQLServerサービスのアクセス権限リストの出力に失敗しました。")
            Exit Function
        End If

        'コマンドの実行結果を読み込みます
        sAcc = ReadText(SD_PATH)

        '改行コードを取り除きます
        sAcc = Replace(sAcc, vbCrLf, "")

        '正常終了を返します
        GetSD = True

    End Function

    '********************************************************************************
    ' メソッド名：IsDone()
    ' 概要　　　：すでにPowerUsersのアクセス権限が設定されているかどうかを返します
    ' パラメータ：[sAcc]...アクセス権限リスト
    ' 戻り値　　：PowerUsersのアクセス権限が設定されていればTrue、設定されていなければFalse
    '********************************************************************************
    Private Function IsDone(ByVal sAcc)

        'PowerUsersのアクセス権限を含む文字列の存在チェックを行います
        If (0 < InStr(1, sAcc, ";;;PU)")) Then
            'あれば戻り値としてTrueを返します
            IsDone = True
        Else
            'なければ戻り値としてFalseを返します
            IsDone = False
        End If

    End Function

    '********************************************************************************
    ' メソッド名：RemakeDACL()
    ' 概要　　　：引数に指定されたアクセス権限リストを変更します
    ' パラメータ：[sAcc]...アクセス権限リスト
    ' 戻り値　　：変更したアクセス権限リスト
    '********************************************************************************
    Private Function RemakeDACL(ByVal sAcc)

        'PowerUsersに付与する権限を格納する変数を定義します
        Dim sAccPU
        sAccPU = ""

        '位置を示す変数を取得します
        Dim pos

        '"D:"の位置を取得します
        pos = InStr(1, sAcc, "D:")

        '"D:"が見つからなければ処理を抜けます
        If (pos < 1) Then
            Call MsgBox("アクセス権限リストの記述が想定外のため、処理できませんでした。")
            Exit Function
        End If

        'POSの位置をアクセス権限の文字列の開始位置まで移動します
        pos = InStr(1, sAcc, "(")

        'ユーザー名を検索します
        Do
            'アクセス権限リストの文字列を終端まで検索した場合はループを抜けます
            If (Len(sAcc) < pos) Then
                Exit Do
            End If

            '次に続く文字がユーザー定義のブロックではない場合はループを抜けます
            If (Mid(sAcc, pos, 1) <> "(") Then
                Exit Do
            End If

            'ユーザー単位の権限の末尾を取得します
            Dim posE
            posE = InStr(pos, sAcc, ")")

            'ユーザー名を取得します
            Dim usr
            usr = Mid(sAcc, posE - 2, 2)

            '見つかったユーザーが"SY"（Local System）ユーザーの場合はその権限をコピーします
            If (usr = "SY") Then
                sAccPU = Mid(sAcc, pos, posE - pos + 1)
            End If

            '次のユーザー権限に移行します
            pos = posE + 1
        Loop

        'PowerUsersにコピーする予定のLocal Systemユーザーのアクセス権限を取得できなければ処理を抜けます
        If (sAccPU = "") Then
            Call MsgBox("Local Systemユーザーのアクセス権限が取得できませんでした。")
            Exit Function
        End If

        'Local Systemユーザーのアクセス権限が想定外の文字列の場合は処理を抜けます
        If (InStr(1, sAccPU, ";;;SY") < 1) Then
            Call MsgBox("Local Systemユーザーのアクセス権限が想定外のため、処理できませんでした。")
            Exit Function
        End If

        'Local Systemユーザーのアクセス権限の文字列にて、ユーザー名をPowerUsersに置き換えた文字列を取得します
        sAccPU = Replace(sAccPU, ";;;SY", ";;;PU")

        '戻り値となる変数を定義します
        Dim sRet
        sRet = Left(sAcc, posE) & sAccPU & Right(sAcc, Len(sAcc) - posE)

        '書き換えの結果が想定の場合、戻り値として空文字列を返します（編集失敗）
        If (Replace(sRet, sAccPU, "") <> sAcc) Then
            sRet = ""
        End If

        '編集後の文字列を返します
        RemakeDACL = sRet

    End Function

    '********************************************************************************
    ' メソッド名：SetSD()
    ' 概要　　　：SQLServerサービスのアクセス権限を設定します
    ' パラメータ：[sAcc]...アクセス権限リスト
    ' 戻り値　　：正常に設定できればTrue、そうでなければFalse
    '********************************************************************************
    Private Function SetSD(ByVal sAcc)

        '戻り値を初期化します
        SetSD = False

        '実行するコマンドを作成します
        Dim cmd
        cmd = "sc sdset MSSQLSERVER " & sAcc

        '作成したコマンドを実行します（処理が完了するまで待機しません）
        If (CommandBatShell(cmd, False) = False) Then
            Exit Function
        End If

        '正常終了を返します
        SetSD = True

    End Function

    '********************************************************************************
    ' メソッド名：CommandBatShell()
    ' 概要　　　：引数に指定されたコマンドを実行します
    ' パラメータ：[cmd]...実行するコマンド
    ' 　　　　　　[bWait]...処理が完了するまで待機する場合はTrue、待機しない場合はFalseを指定
    ' 戻り値　　：正常に実行できればTrue、そうでなければFalse
    '********************************************************************************
    Private Function CommandBatShell(ByVal cmd, ByVal bWait)

        '戻り値を初期化します
        CommandBatShell = False

        'バッチファイルをテンポラリに作成します
        Dim batpath
        batpath = GetTempPath() & "\sccmd.bat"
        If (IsFileExists(batpath)) Then
            'すでに前回実行したバッチファイルが存在する場合は削除します
            If (RemoveFile(batpath) = False) Then
                Exit Function
            End If
        End If

        '引数に指定されたコマンドをバッチファイルに出力します
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim bat
        Set bat = fso.CreateTextFile(batpath)
        bat.WriteLine(cmd)
        bat.Close()

        '待機が必要な場合
        If (bWait) Then
            'バッチファイルを実行し、終了するまで待機します
            Dim intRet
            intRet = WScript.CreateObject("WScript.Shell").Run(batpath, 1, true)

            'エラーコードが返ってきた場合は処理を抜けます
            If (intRet <> 0) Then
                Dim msg
                msg = ""
                msg = msg & "コマンド実行に失敗しました。" & vbCrLf
                msg = msg & vbCrLf
                msg = msg & "エラーコード：" & CStr(intRet)

                Call MsgBox(msg, vbCritical + vbOkOnly)
                Exit Function
            End If

        '待機が不要な場合は管理者モードで実行
        Else
            'バッチファイルを管理者として実行します
            Dim sh
            Set sh = Wscript.CreateObject("Shell.Application")

            Call sh.ShellExecute(batpath, , , "runas")
        End If

        '正常終了を返します
        CommandBatShell = True

    End Function

    '********************************************************************************
    ' メソッド名：IsFileExists()
    ' 概要　　　：引数に指定されたファイルの存在を確認します
    ' パラメータ：[fpath]...存在確認を行うファイルのフルパス
    ' 戻り値　　：ファイルが存在する場合はTrue、存在しない場合はFalse
    '********************************************************************************
    Function IsFileExists(ByVal fpath)

        'Scripting.FileSystemObjectのCOMを参照します
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")

        'FileExists()メソッドの戻り値をこの関数の戻り値として返します
        IsFileExists = fso.FileExists(fpath)

    End Function

    '********************************************************************************
    ' メソッド名：RemoveFile()
    ' 概要　　　：引数に指定されたファイルを削除します
    ' パラメータ：[fpath]...削除するファイルのフルパス
    ' 戻り値　　：正常に削除できればTrue、そうでなければFalse
    '********************************************************************************
    Function RemoveFile(ByVal fpath)

        '戻り値を初期化します
        RemoveFile = False

        'Scripting.FileSystemObjectのCOMを参照します
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")

        'エラーが発生した場合でも処理を続行します
        On Error Resume Next

        '引数に指定されたファイルを削除します
        Call fso.DeleteFile(fpath)

        'エラーが発生していた場合のその内容をメッセージ表示して処理を抜けます
        If (Err.Number <> 0) Then
            Call MsgBox(CStr(Err.Number) & ": " & Err.Description, vbCritical + vbOkOnly)
            Exit Function
        End If

        '正常終了を返します
        RemoveFile = True

    End Function

    '********************************************************************************
    ' メソッド名：ReadText()
    ' 概要　　　：引数に指定されたテキストファイルを読み込み、その内容を返します
    ' パラメータ：[tfpath]...読み込み対象となるテキストファイルのフルパス
    ' 戻り値　　：テキストファイルの内容
    '********************************************************************************
    Function ReadText(ByVal tfpath)

        '戻り値を初期化します
        ReadText = ""

        'Scripting.FileSystemObjectのCOMを参照します
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")

        '引数に指定されたファイルを読み込みます
        Dim s
        s = fso.OpenTextFile(tfpath, 1).ReadAll()

        '読み込んだ内容を戻り値として返します
        ReadText = s

    End Function

    '********************************************************************************
    ' メソッド名：GetTempPath()
    ' 概要　　　：テンポラリパスを取得します
    ' パラメータ：なし
    ' 戻り値　　：テンポラリパス
    '********************************************************************************
    Function GetTempPath()

        'Scripting.FileSystemObjectのCOMを参照します
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")

        'テンポラリのフルパスを取得します
        Dim tp
        tp = fso.GetSpecialFolder(2).Path

        '取得したテンポラリのフルパスを返します
        GetTempPath = tp

    End Function

End Class
