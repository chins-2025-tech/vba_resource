# エクスポート・インポートの自動化

## 重要な事前設定

VBAコードでVBAプロジェクトを操作するには、Excelのセキュリティ設定を変更する必要があります。これを有効にしないと、マクロはエラーとなり実行できません。

1. Excelを開き、「ファイル」タブ → 「オプション」をクリックします。
2. 「Excelのオプション」ダイアログで、「トラストセンター」（または「セキュリティセンター」） → 「トラストセンターの設定」をクリックします。
3. 「マクロの設定」を選び、「VBAプロジェクト オブジェクト モデルへのアクセスを信頼する」にチェックを入れます。
4. 「OK」を数回クリックしてダイアログを閉じます。

**【警告】**<br>
この設定は、悪意のあるマクロが他のマクロを勝手に書き換えることを許可するものでもあります。信頼できるファイルでのみマクロを実行するように、引き続き注意してください。


## コード
```vba
Option Explicit

'================================================================================
' 機能：ブック内の全VBAコンポーネントを指定フォルダにバックアップ/復元する
' 特徴：
' ・インポート処理にFSO.Filesコレクションを使用し、可読性を向上
' ・バックアップファイルを軸にした上書き更新（追加モジュールを削除しない）
' ・ドキュメントモジュールの復元にAddFromFileを使用し、効率化
' 作成者：chins
'================================================================================

' --- 設定項目 ---
Private Const BACKUP_FOLDER_NAME As String = "VBA_Backup"
Private Const THIS_MODULE_NAME As String = "projVerManger"

'================================================================================
' ■ エクスポート（バックアップ）実行
'================================================================================
Public Sub ExportAllVbaComponents()
    Dim vbProj As Object 'VBProject
    Dim vbComp As Object 'VBComponent
    Dim exportFolder As String

    On Error GoTo ErrorHandler

    Set vbProj = ThisWorkbook.VBProject
    

    ' フォルダ選択ダイアログを表示してエクスポート先フォルダを指定
    exportFolder = GetExportFolder()
    If exportFolder = "" Then　Exit Sub
    If Dir(exportFolder, vbDirectory) = "" Then MkDir exportFolder
    
    For Each vbComp In vbProj.VBComponents
        Select Case vbComp.Type
            Case 1: ExportComponentFile vbComp, exportFolder, ".bas"
            Case 2: ExportComponentFile vbComp, exportFolder, ".cls"
            Case 3: ExportComponentFile vbComp, exportFolder, ".frm"
            Case 100
                If vbComp.CodeModule.CountOfLines > 0 Then
                    If vbComp.name = "ThisWorkbook" Then
                        ExportCodeOnlyToFile vbComp, exportFolder, ".book"
                    Else
                        ExportCodeOnlyToFile vbComp, exportFolder, ".sheet"
                    End If
                End If
        End Select
    Next vbComp
    
    MsgBox "VBAコードのエクスポートが完了しました。" & vbCrLf & "保存先: " & exportFolder, vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "エクスポート中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical
End Sub


'================================================================================
' ■ インポート（復元）実行
'================================================================================
Public Sub ImportAllVbaComponents_Final()
    Dim vbProj As Object 'VBProject
    Dim fso As Object
    Dim importFolder As Object 'Folder
    Dim targetFile As Object 'File
    
    On Error GoTo ErrorHandler
    
    If MsgBox("バックアップフォルダのファイルに基づいてVBAコードを上書き更新します。" & vbCrLf & _
              "この操作は元に戻せません。よろしいですか？", _
              vbQuestion + vbYesNo, "最終確認") = vbNo Then
        Exit Sub
    End If
    
    Set vbProj = ThisWorkbook.VBProject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' バックアップフォルダのパスをチェック
    Dim folderPath As String
    folderPath = GetExportFolder()
    If folderPath = "" Then
        GoTo CleanUp
    End If

    If Not fso.FolderExists(folderPath) Then
        MsgBox "バックアップフォルダが見つかりません: " & folderPath, vbCritical
        GoTo CleanUp
    End If
    
    Set importFolder = fso.GetFolder(folderPath)
    
    For Each targetFile In importFolder.Files
        Debug.Print targetFile.name
    Next targetFile
    
   ' --- バックアップフォルダ内の全ファイルをループ ---
    For Each targetFile In importFolder.Files
        Dim compName As String, fileExt As String
        compName = fso.GetBaseName(targetFile.name)
        fileExt = fso.GetExtensionName(targetFile.name)
        
        If compName = THIS_MODULE_NAME Then GoTo SkipLoop
        
        If ComponentExists(vbProj, compName) Then
            ' --- 存在する場合の処理 ---
            Select Case LCase(fileExt)
                Case "bas", "cls", "frm"
                    ' 既存のものを削除してからインポート
                    vbProj.VBComponents.Remove vbProj.VBComponents(compName)
                    vbProj.VBComponents.Import targetFile.Path
                Case "book", "sheet"
                    ' コードを上書き
                    ImportCodeFromFile vbProj.VBComponents(compName), targetFile.Path
            End Select
        Else
            ' --- 存在しない場合の処理 (削除されたモジュールの復元) ---
            Select Case LCase(fileExt)
                Case "bas", "cls", "frm"
                    ' 新規にインポート
                    vbProj.VBComponents.Import targetFile.Path
                Case "book", "sheet"
                    ' ドキュメントモジュールは新規作成できないので、何もしない（or 警告）
                    Debug.Print "警告: コンポーネント '" & compName & "' が存在しないため、" & _
                                "ファイル '" & targetFile.name & "' のインポートをスキップしました。"
            End Select
        End If
        
SkipLoop:
    Next targetFile
    
    MsgBox "VBAコードのインポート（上書き更新）が完了しました。", vbInformation

CleanUp:
    ' オブジェクトを解放
    Set targetFile = Nothing
    Set importFolder = Nothing
    Set fso = Nothing
    Set vbProj = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "インポート中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical
    GoTo CleanUp
End Sub

'================================================================================
' 機能：エクスポート先フォルダをユーザーに選択させるダイアログを表示し、パスを返す
' 戻り値：選択されたフォルダのパス（キャンセル時は空文字列）
'================================================================================
Private Function GetExportFolder() As String
    Dim fd As Object
    Set fd = Application.FileDialog(4) ' msoFileDialogFolderPicker = 4
    With fd
        .Title = "エクスポート先フォルダを選択してください"
        .InitialFileName = ThisWorkbook.Path & "\" & BACKUP_FOLDER_NAME
        If .Show = -1 Then
            GetExportFolder = .SelectedItems(1)
        Else
            MsgBox "フォルダが選択されませんでした。処理を中止します。", vbExclamation
            GetExportFolder = ""
        End If
    End With
    Set fd = Nothing
End Function

'================================================================================
' 機能：指定された名前のVBAコンポーネントが存在するかどうかを安全に判定する
' 引数：
'   proj: VBProject オブジェクト
'   name: 判定したいコンポーネント名
' 戻り値：
'   存在するなら True, 存在しないなら False
'================================================================================
Private Function ComponentExists(ByVal proj As Object, ByVal name As String) As Boolean
    Dim comp As Object
    On Error Resume Next
    Set comp = proj.VBComponents(name)
    On Error GoTo 0 ' エラーハンドリングを通常に戻す
    ComponentExists = Not (comp Is Nothing)
End Function

Private Sub ExportComponentFile(ByVal comp As Object, ByVal folderPath As String, ByVal ext As String)
    Dim exportPath As String
    exportPath = folderPath & "\" & comp.name & ext
    If Dir(exportPath) <> "" Then Kill exportPath
    comp.Export exportPath
End Sub

Private Sub ExportCodeOnlyToFile(ByVal comp As Object, ByVal folderPath As String, ByVal ext As String)
    Dim exportPath As String
    Dim fso As Object, ts As Object
    
    exportPath = folderPath & "\" & comp.name & ext
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(exportPath, True) ' 上書き
    ts.Write comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
    ts.Close
    Set fso = Nothing
End Sub

Private Sub ImportCodeFromFile(ByVal comp As Object, ByVal filePath As String)
    If comp.CodeModule.CountOfLines > 0 Then
        comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
    End If
    comp.CodeModule.AddFromFile filePath
End Sub

```