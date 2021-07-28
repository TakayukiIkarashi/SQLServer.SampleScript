Attribute VB_Name = "basMain"
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' SQLServerのサービスが停止していた場合、起動します
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

'----------------------------------------
' Win32 API
'----------------------------------------
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

'----------------------------------------
' 定数定義
'----------------------------------------
'サービスの状態
Public Const STATUS_RUNNING As String = "RUNNING"
Public Const STATUS_STOPPED As String = "STOPPED"
Public Const STATUS_START_PENDING As String = "START PENDING"
Public Const STATUS_STOP_PENDING As String = "STOP PENDING"

'********************************************************************************
' 関数名：Main
' 概要　：起動時処理
' 引数　：なし
' 戻り値：なし
'********************************************************************************
Sub Main()

    'SQL Server サービスのインスタンスを定義します
    Dim Service As Object

    'SQL Server サービスの状態を取得します
    Dim s As String
    s = GetStatus(Service)

    'SQL Server サービスの状態を確認します
    Select Case UCase(s)
    '停止中
    Case STATUS_STOPPED
        'SQL Server サービスを起動します
        If (StartService(Service) = False) Then
            '起動に失敗した場合、ダイアログを表示してサービスが起動するまで待機します
            Call ShowDialog
        End If

    '実行開始中・停止移行中
    Case STATUS_START_PENDING, STATUS_STOP_PENDING
        'ダイアログを表示してサービスが起動するまで待機します
        Call ShowDialog

    '実行中（もしくはそれ以外）
    Case Else
        '特に何もしない

    End Select

End Sub

'********************************************************************************
' 関数名：GetStatus()
' 概要　：SQL Server サービスの状態を文字列で取得します
' 引数　：[Service]...SQL Server サービスのオブジェクト
' 戻り値：SQL Server サービスの状態（文字列）
'********************************************************************************
Public Function GetStatus(ByRef Service As Object) As String

    '戻り値を初期化します
    GetStatus = ""

    Set Service = Nothing

    'エラーが発生した場合は例外処理に移行します
    On Error GoTo Exception

    'SQL Serverサービスの状態を取得するためのWQLを定義します
    Dim WQL As String
    WQL = "SELECT * FROM Win32_Service WHERE Name = 'MSSQLSERVER'"

    '定義したWQLを実行します
    Dim ServiceList As Object
    Set ServiceList = CreateObject("WbemScripting.SWbemLocator").ConnectServer.ExecQuery(WQL)

    'SQL Serverサービスの状態を取得します
    For Each Service In ServiceList
        '戻り値をセットします
        GetStatus = Service.state
        Exit For
    Next

    Exit Function

Exception:
    'エラーが発生しても何もしません

End Function

'********************************************************************************
' 関数名：StartService
' 概要　：サービスを開始します
' 引数　：[Service]...開始するサービス
' 戻り値：サービスの開始に成功したらTrue、失敗したらFalse
'********************************************************************************
Public Function StartService(ByRef Service As Object) As Boolean

    'サービスを起動します
    Dim lngRet As Long
    lngRet = Service.StartService()

    '戻り値が0なら成功、0以外なら失敗
    If (lngRet = 0) Then
        StartService = True
    Else
        StartService = False
    End If

End Function

'********************************************************************************
' 関数名：StopService
' 概要　：サービスを停止します
' 引数　：[Service]...停止するサービス
' 戻り値：サービスの停止に成功したらTrue、失敗したらFalse
'********************************************************************************
Public Function StopService(ByRef Service As Object) As Boolean

    'サービスを停止します
    Dim lngRet As Long
    lngRet = Service.StopService()

    '戻り値が0なら成功、0以外なら失敗
    If (lngRet = 0) Then
        StopService = True
    Else
        StopService = False
    End If

End Function

'********************************************************************************
' 関数名：ShowDialog
' 概要　：サービス開始中ダイアログを表示します
' 引数　：なし
' 戻り値：なし
'********************************************************************************
Private Sub ShowDialog()

    'サービス起動中ダイアログを表示します
    Dim f As New frmDialog
    f.Show vbModal

End Sub
