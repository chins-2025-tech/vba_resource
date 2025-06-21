# Excel VBA 高機能ロギングモジュール

## 概要
Excel VBAの編集作業を自動化する際に利用する、高機能で再利用可能なロギングモジュール。以下の特徴を持つ。

- シングルトンパターン: アプリケーション全体で唯一のロガーインスタンスを共有し、設定管理を簡素化。
- 出力先の選択: イミディエイトウィンドウとテキストファイルのどちらかに出力を選択可能。
- デフォルト設定: デフォルトの出力先はテキストファイル。ログファイルは、ユーザーの一時フォルダ (%TEMP%) 配下に EXCEL_VBA_{タイムスタンプ}.log という名前で自動生成される。
- 高精度タイムスタンプ: ログにはミリ秒単位のタイムスタンプが付与され、詳細な時間追跡が可能。
- タグ機能: ログメッセージに任意のタグ（例: [MAIN], [ERROR]）を付与でき、ログの可読性を向上。

## 構成ファイル

### 1. clsLogger (クラスモジュール)
ロギング機能の本体をカプセル化したクラス。

```vba
'================================================================================
' Class: clsLogger
' Description: ログ出力を管理するクラスモジュール。
'              シングルトンパターンでの利用を想定。
'================================================================================
Option Explicit

'--- 列挙体: ログの出力先を定義
Public Enum LogOutputType
    logOutputImmediate  ' 0: イミディエイトウィンドウ
    logOutputFile       ' 1: テキストファイル
End Enum

'--- プライベート変数
Private pOutputType As LogOutputType
Private pLogFilePath As String
Private pIsEnabled As Boolean
Private pFileNum As Integer

'================================================================================
' プロパティ
'================================================================================

Public Property Get IsEnabled() As Boolean
    IsEnabled = pIsEnabled
End Property
Public Property Let IsEnabled(ByVal value As Boolean)
    pIsEnabled = value
End Property

Public Property Get OutputType() As LogOutputType
    OutputType = pOutputType
End Property
Public Property Let OutputType(ByVal value As LogOutputType)
    pOutputType = value
End Property

Public Property Get LogFilePath() As String
    LogFilePath = pLogFilePath
End Property
Public Property Let LogFilePath(ByVal value As String)
    pLogFilePath = value
End Property

'================================================================================
' イベントプロシージャ (コンストラクタとデストラクタ)
'================================================================================

' インスタンス生成時に実行される (コンストラクタ)
Private Sub Class_Initialize()
    pIsEnabled = True
    pOutputType = logOutputFile ' デフォルトはファイル出力

    ' デフォルトのログファイルパスをTEMPフォルダに設定
    On Error Resume Next
    Dim tempPath As String
    tempPath = Environ("TEMP")
    If tempPath = "" Then tempPath = ThisWorkbook.Path ' TEMPが取得できない場合のフォールバック
    On Error GoTo 0
    
    pLogFilePath = tempPath & "\EXCEL_VBA_" & Format(Now, "yyyymmdd_hhmmss") & ".log"
    
    pFileNum = 0
End Sub

' インスタンス破棄時に実行される (デストラクタ)
Private Sub Class_Terminate()
    ' ファイルが開かれていれば、安全に閉じる
    If pFileNum <> 0 Then
        Close #pFileNum
        pFileNum = 0
    End If
End Sub

'================================================================================
' メインメソッド
'================================================================================

' ログを記録する (タグ機能付き)
Public Sub Log(ByVal message As String, Optional ByVal tag As String = "")
    If Not pIsEnabled Then Exit Sub

    ' タイムスタンプ生成 (ミリ秒対応)
    Dim ms As String
    ms = Right(Format(Timer, "#0.000"), 3)
    Dim timestamp As String
    timestamp = "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "." & ms & "]"

    ' タグがあればフォーマットに追加
    Dim tagStr As String
    If tag <> "" Then
        tagStr = " [" & tag & "]"
    End If

    ' 最終的なログメッセージを生成
    Dim logMessage As String
    logMessage = timestamp & tagStr & " " & message

    ' 設定された出力先に応じて処理を分岐
    Select Case pOutputType
        Case logOutputImmediate
            Debug.Print logMessage
        Case logOutputFile
            WriteToFile logMessage
    End Select
End Sub

'================================================================================
' プライベートメソッド (内部処理)
'================================================================================

Private Sub WriteToFile(ByVal text As String)
    If pLogFilePath = "" Then
        Debug.Print "【Logger Error】LogFilePath is not set."
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    If pFileNum = 0 Then
        pFileNum = FreeFile
        Open pLogFilePath For Append As #pFileNum
    End If
    Print #pFileNum, text
    Exit Sub
ErrorHandler:
    Debug.Print "【Logger Error】Failed to write to file: " & pLogFilePath & ". Error: " & Err.Description
    If pFileNum <> 0 Then
        Close #pFileNum
        pFileNum = 0
    End If
End Sub
```

### 2. LoggerManager (標準モジュール)
clsLogger のシングルトンインスタンスを管理・提供するモジュール。

```vba
'================================================================================
' Module: LoggerManager
' Description: clsLoggerのシングルトンインスタンスを管理する。
'              アプリケーション全体で唯一のロガーへのアクセスを提供する。
'================================================================================
Option Explicit

Private gLogger As clsLogger

'--- シングルトンインスタンスへのアクセス関数 ---
' この関数を通じて、常に同じロガーインスタンスを取得する
Public Function Logger() As clsLogger
    If gLogger Is Nothing Then
        Set gLogger = New clsLogger
    End If
    Set Logger = gLogger
End Function

'--- ロガーを明示的に初期化（またはリセット）するプロシージャ ---
' マクロの最初に呼び出すことで、ログファイル名をその時点の時刻で確定させる
Public Sub InitializeLogger()
    Set gLogger = Nothing
    Set gLogger = New clsLogger
End Sub

'--- インスタンスを明示的に解放するプロシージャ ---
' マクロの最後に呼び出し、ログファイルを即座に閉じる
Public Sub ReleaseLogger()
    Set gLogger = Nothing
End Sub
```

## 使い方 (サンプルコード)
以下は、このロギングモジュールを利用したマクロのサンプル。
### 3. SampleUsage (標準モジュール)

```vba
'================================================================================
' Module: SampleUsage
' Description: シングルトンロガーの使い方を示すサンプルプロシージャ
'================================================================================
Option Explicit

Sub MainProcess()
    ' 1. ロガーを初期化。ログファイル名がこの時点で確定する。
    InitializeLogger
    
    ' 2. (任意) デフォルト設定を上書きする場合
    ' With Logger
    '     .OutputType = logOutputImmediate ' 出力先をイミディエイトに変更
    '     .LogFilePath = ThisWorkbook.Path & "\MyCustomLog.txt" ' ログファイルパスをカスタム
    ' End With

    ' 3. ログを記録
    Logger.Log "メイン処理を開始します。", "MAIN"

    On Error GoTo ErrorHandler
    
    Call SubProcess1
    Call SubProcess2
    
    Logger.Log "全ての処理が正常に完了しました。", "MAIN"
    GoTo Finally

ErrorHandler:
    Logger.Log "致命的なエラーが発生しました。処理を中断します。", "ERROR"
    Logger.Log "エラー番号: " & Err.Number, "ERROR"
    Logger.Log "エラー内容: " & Err.Description, "ERROR"

Finally:
    ' 4. ロガーを解放し、ログファイルを閉じる
    Logger.Log "ログセッションを終了します。", "SYSTEM"
    ReleaseLogger
End Sub


Private Sub SubProcess1()
    Logger.Log "サブ処理1を開始...", "SUB1"
    Application.Wait (Now + TimeValue("00:00:01")) ' 1秒待機
    Logger.Log "サブ処理1が完了。", "SUB1"
End Sub

Private Sub SubProcess2()
    Logger.Log "サブ処理2を開始...", "SUB2"
    Application.Wait (Now + TimeValue("00:00:01")) ' 1秒待機
    Logger.Log "サブ処理2が完了。", "SUB2"
End Sub
```