VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "acCSV取込補助"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Private clsテーブル名 As String
Private clsCmd As DoCmd


Private Sub Class_Initialize()

    clsテーブル名 = ""
    Set clsCmd = DoCmd

End Sub

Private Sub Class_Terminate()

    clsテーブル名 = ""
    Set clsCmd = Nothing

End Sub

Private Function getFilePicker(Optional dTitle As String = "ファイル選択", Optional dURL As String = "This")

    '2003以降
    Const msoFileDialogFilePicker As Integer = 3
    Dim fDlg As Object
    Set fDlg = Application.FileDialog(msoFileDialogFilePicker)

    With fDlg

        .Title = dTitle
        If dURL = "This" Then
            .InitialFileName = CurrentProject.Path
        ElseIf dURL <> "" Then
            If Dir(dURL, vbDirectory) <> "" Then
                .InitialFileName = dURL
            End If
        End If
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "すべてのファイル", "*.*"
        .Filters.Add "CSV ﾌｧｲﾙ (*.csv)", "*.csv"
        .FilterIndex = 1

        If .Show Then getFilePicker = .SelectedItems(1) Else getFilePicker = ""
    
    End With

End Function

Private Function Chk() As Boolean

    If clsテーブル名 = "" Then
    
        Chk = True
        
        MsgBox "操作テーブル設定が必要です。" & vbCrLf & "操作するテーブル名を選択して下さい。"
        
    Else
    
        Chk = False
    
    End If

End Function

Public Function テーブル(ByVal テーブル名 As String) As acCSV取込補助
Attribute テーブル.VB_UserMemId = 0

    clsテーブル名 = テーブル名
    
    Set テーブル = Me

End Function

Public Sub 取込()

    Dim msg As String

    msg = getFilePicker
    If msg = "" Or Chk Then Exit Sub

    On Error GoTo err_sample
    
    With clsCmd
        .SetWarnings False
        .RunSQL "DELETE * FROM [" & clsテーブル名 & "]"
        .TransferText acImportDelim, , clsテーブル名, msg, True
        .SetWarnings True
    End With

    Me.更新

    MsgBox "インポートが終了しました。"
    Exit Sub
err_sample:
    Select Case Err.Number
    Case 3011
        MsgBox "ファイルが見つかりません。処理を終了します。"
    Case Else
        MsgBox Err.Number & ":" & Err.Description
    End Select

End Sub

Public Function 開く() As Long

    If Chk Then Exit Function
    
    On Error Resume Next

    With clsCmd
        .Requery
        .OpenTable clsテーブル名, acViewNormal, acReadOnly
    End With
    
    開く = Err.Number
    
    On Error GoTo 0

End Function

Public Function 閉じる() As Long

    If Chk Then Exit Function
    
    On Error Resume Next

    clsCmd.Close acTable, clsテーブル名, acSavePrompt
    
    閉じる = Err.Number
    
    On Error GoTo 0

End Function

Public Function 更新() As Long

    If Chk Then Exit Function

    On Error Resume Next

    Me.閉じる
    clsCmd.Requery
    
    更新 = Err.Number
    
    On Error GoTo 0

End Function

