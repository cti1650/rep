VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OfficeDoc文字列一括置換"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function MakeSureDirectoryPathExists _
  Lib "imagehlp.dll" (ByVal lpPath As String) As Long

Private myExcel As Object
Private myBk As Object
Private mySel As Object
Private myWord As Object
Private myDoc As Object
Private myURL As String

Private Const ErrCheck As Boolean = False

Private Sub Class_Initialize()

#If ErrCheck = False Then

    Application.DisplayAlerts = False

#Else

    Application.DisplayAlerts = True

#End If

    Set myExcel = createExcel
    Set myWord = createWord
    
    myURL = ""

End Sub

Private Sub Class_Terminate()

    myExcel.Quit
    myWord.Quit
    
    myURL = ""

    Set mySel = Nothing
    Set myBk = Nothing
    Set myExcel = Nothing
    Set myDoc = Nothing
    Set myWord = Nothing
    
    Application.DisplayAlerts = True

End Sub

Private Function createExcel()

#If ErrCheck = False Then

    On Error Resume Next

#Else

    On Error GoTo 0

#End If
    
    Dim objExcel As Object

    Err.Clear

    Set objExcel = CreateObject("Excel.Application")
    
#If ErrCheck = False Then

    objExcel.Visible = False

#Else

    objExcel.Visible = True

#End If
    
    If Err.Number = 0 Then
    
        Set createExcel = objExcel
        
    Else
    
        Set createExcel = Nothing
    
    End If
    
    On Error GoTo 0

End Function

Private Function createWord()

#If ErrCheck = False Then

    On Error Resume Next

#Else

    On Error GoTo 0

#End If
    
    Dim objWord As Object

    Err.Clear

    Set objWord = CreateObject("Word.Application")
    
#If ErrCheck = False Then

    objWord.Visible = False

#Else

    objWord.Visible = True

#End If
    
    If Err.Number = 0 Then
    
        Set createWord = objWord
        
    Else
    
        Set createWord = Nothing
    
    End If
    
    On Error GoTo 0

End Function

Friend Function 開く(ByVal ファイル名 As String)

#If ErrCheck = False Then

    On Error Resume Next

#Else

    On Error GoTo 0

#End If

    Err.Clear
    
    Select Case True
    Case ファイル名 Like "*.xls*"

        If myExcel Is Nothing Then Exit Function
    
        If Not ファイル名 Like "*\*" Then
    
            ファイル名 = ThisWorkbook.Path & "\" & ファイル名

        End If
    
        Set myBk = myExcel.Workbooks.Open(ファイル名)
    
        If Err.Number = 0 Then
    
            myURL = ファイル名
    
            Set 開く = myBk
        
        Else
    
            myURL = ""
    
            Set 開く = Nothing
    
        End If
    
    Case ファイル名 Like "*.doc*"
    
        If myWord Is Nothing Then Exit Function
    
        If Not ファイル名 Like "*\*" Then
    
            ファイル名 = ThisWorkbook.Path & "\" & ファイル名

        End If
    
        Set myDoc = myWord.Documents.Open(ファイル名)
    
        Set mySel = myWord.Selection
    
        If Err.Number = 0 Then
    
            myURL = ファイル名
    
            Set 開く = myDoc
        
        Else
    
            myURL = ""
    
            Set 開く = Nothing
    
        End If
        
    Case Else
    
        myURL = ""
    
        Set 開く = Nothing
    
    End Select
    
    On Error GoTo 0
    
End Function

Friend Sub 閉じる(Optional ByVal 保存 As Boolean = True)

#If ErrCheck = False Then

    On Error Resume Next

#Else

    On Error GoTo 0

#End If

    Err.Clear
    
    Select Case True
    Case myURL Like "*.xls*"

        If myBk Is Nothing Then Exit Sub
        
        Set mySel = Nothing
    
        If 保存 = True Then
    
            myBk.Save
        
        End If
    
        myBk.Close
    
        myURL = ""
        
        Set myBk = Nothing
    
    Case myURL Like "*.doc*"
    
        If myDoc Is Nothing Then Exit Sub
    
        Set mySel = Nothing
    
        If 保存 = True Then
    
            myDoc.Save
        
        End If
    
        myDoc.Close
    
        myURL = ""
        
        Set myDoc = Nothing
    
    End Select
    
    On Error GoTo 0
    
End Sub

Friend Property Get 種類() As Variant

    Select Case True
    Case myURL Like "*.xls*"

        種類 = "Excel"
        
    Case myURL Like "*.doc*"
    
        種類 = "Word"
        
    Case Else
    
        種類 = "対象外"
    
    End Select

End Property

Friend Property Get ファイルパス() As Variant

    ファイルパス = myURL

End Property

Friend Property Let ファイルパス(ByVal URL As Variant)

    Me.開く URL

End Property

Friend Sub 文字置換(ByVal 検索文字 As String, ByVal 置換文字 As String)

#If ErrCheck = False Then

    On Error Resume Next

#Else

    On Error GoTo 0

#End If

    Err.Clear
    
    Select Case True
    Case myURL Like "*.xls*"

        If myBk Is Nothing Then Exit Sub
    
        Dim i As Integer
        Dim sht As Excel.Worksheet
    
        For i = 1 To myBk.Sheets.Count
            Set sht = myBk.Sheets(i)
            If sht.Name Like "*" & 検索文字 & "*" Then
                sht.Name = Replace(sht.Name, 検索文字, 置換文字)
            End If
            If Not sht.Cells.Find(検索文字) Is Nothing Then
                sht.Cells.Replace What:=検索文字, Replacement:=置換文字, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, MatchByte:=False
            End If
        Next
    
    Case myURL Like "*.doc*"
    
        Const wdReplaceAll = 2

        If mySel Is Nothing Then Exit Sub
    
        mySel.Find.ClearFormatting
        mySel.Find.Replacement.ClearFormatting

        mySel.Find.Text = 検索文字
        mySel.Find.Forward = True
        mySel.Find.MatchWholeWord = False

        mySel.Find.Replacement.Text = 置換文字
        mySel.Find.Execute Replace:=wdReplaceAll
    
    End Select
    

    On Error GoTo 0

End Sub

Friend Function ファイル名文字置換(ByVal 検索文字 As String, ByVal 置換文字 As String, ParamArray 除外文字() As Variant) As String

#If ErrCheck = False Then

    On Error Resume Next

#Else

    On Error GoTo 0

#End If
    
    Dim 変更前ファイル名 As String
    Dim 変更後ファイル名 As String
    Dim 検証用検索文字 As String
    Dim i As Long
    Dim 初期除外値 As Variant
    
    初期除外値 = Array("[", "]", "*")

    変更前ファイル名 = myURL
    
    検証用検索文字 = 検索文字
    
    For i = LBound(初期除外値) To UBound(初期除外値)
        検証用検索文字 = Replace(検証用検索文字, 初期除外値(i), "?")
    Next i
    
    For i = LBound(除外文字) To UBound(除外文字)
        検証用検索文字 = Replace(検証用検索文字, 除外文字(i), "?")
    Next i

    If 変更前ファイル名 Like "*" & 検証用検索文字 & "*" Or 変更前ファイル名 = 検索文字 Then
    
        変更後ファイル名 = Replace(変更前ファイル名, 検索文字, 置換文字)
    
        Me.閉じる
            
        Name 変更前ファイル名 As 変更後ファイル名
        
        Me.開く 変更後ファイル名
            
    Else
    
#If ErrCheck = True Then
MsgBox ("失敗 " & 検索文字)
#End If
         
        変更後ファイル名 = 変更前ファイル名
                
    End If
    
    If Err.Number = 0 Then
    
        ファイル名文字置換 = 変更後ファイル名
        
    Else
    
        ファイル名文字置換 = ""
        
    End If
    
    On Error GoTo 0


End Function

Friend Function フォルダ作成(ByVal フォルダ名 As String) As String

    Select Case True
    Case Not フォルダ名 Like "?[\:]*"
        フォルダ名 = ThisWorkbook.Path & "\" & フォルダ名 & "\"
    Case Not フォルダ名 Like "*\"
        フォルダ名 = フォルダ名 & "\"
    Case フォルダ名 Like "*\*\"
    End Select
    
    If フォルダ名 Like "*" & vbCrLf & "*" Then フォルダ名 = Replace(フォルダ名, vbCrLf, " ")
    
    If MakeSureDirectoryPathExists(フォルダ名) Then
    
        フォルダ作成 = フォルダ名
        
    Else
    
        フォルダ作成 = ""
        
    End If

End Function

Friend Function フォルダコピー(ByVal コピー元パス As String, ByVal コピー先パス As String) As String

    Dim objFileSys    As Object
        
#If ErrCheck = False Then

    On Error Resume Next
    
#End If
        
    Err.Clear
    
    Set objFileSys = CreateObject("Scripting.FileSystemObject")
    
    Select Case True
    Case Not コピー元パス Like "?[\:]*"
        コピー元パス = ThisWorkbook.Path & "\" & コピー元パス
    Case コピー元パス Like "*\*\"
    End Select
    
    Select Case True
    Case Not コピー先パス Like "?[\:]*"
        コピー先パス = ThisWorkbook.Path & "\" & コピー先パス
    Case コピー先パス Like "*\*\"
    End Select
    
    If Dir(コピー元パス, vbDirectory) = "" Then
    
        フォルダコピー = ""
        
        Exit Function
        
    End If
    
    If Dir(コピー先パス & "\*.*") <> "" Then
    
        フォルダコピー = コピー先パス
        
        Exit Function
        
    End If
    
    コピー先パス = Me.フォルダ作成(コピー先パス)
    
    If コピー先パス Like "*\" Then コピー先パス = Left(コピー先パス, Len(コピー先パス) - 1)
    
    If コピー先パス <> "" Then
    
        objFileSys.CopyFolder コピー元パス, コピー先パス
        
    Else
    
        フォルダコピー = ""
        
        Exit Function
    
    End If

    If Err.Number = 0 Then

        フォルダコピー = コピー先パス
        
    Else
    
        Debug.Print "コピーに失敗しました。 " & Err.Description
        
        フォルダコピー = ""
    
    End If
    
    On Error GoTo 0

    Set folderObj = Nothing
    Set objFileSys = Nothing

End Function
