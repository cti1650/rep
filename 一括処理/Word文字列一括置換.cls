VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Word文字列一括置換"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private myWord As Object
Private myDoc As Object
Private mySel As Object
Private myURL As String

Private Sub Class_Initialize()

    Set myWord = createWord
    
    myURL = ""

End Sub

Private Sub Class_Terminate()

    myWord.Quit
    
    myURL = ""

    Set mySel = Nothing
    Set myDoc = Nothing
    Set myWord = Nothing

End Sub

Private Function createWord()

    On Error Resume Next
    
    Dim objWord As Object

    Err.Clear

    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False
    
    If Err.Number = 0 Then
    
        Set createWord = objWord
        
    Else
    
        Set createWord = Nothing
    
    End If
    
    On Error GoTo 0

End Function

Friend Function DocOpen(ByVal ファイル名 As String)

    On Error Resume Next

    Err.Clear

    If myWord Is Nothing Then Exit Function
    
    If Not ファイル名 Like "*\*" Then
    
        ファイル名 = ThisWorkbook.Path & "\" & ファイル名

    End If
    
    Set myDoc = myWord.Documents.Open(ファイル名)
    
    Set mySel = myWord.Selection
    
    If Err.Number = 0 Then
    
        myURL = ファイル名
    
        Set DocOpen = myDoc
        
    Else
    
        myURL = ""
    
        Set DocOpen = Nothing
    
    End If
    
    On Error GoTo 0
    
End Function

Friend Sub DocClose(Optional ByVal 保存 As Boolean = True)

    On Error Resume Next

    Err.Clear

    If myDoc Is Nothing Then Exit Sub
    
    Set mySel = Nothing
    
    If 保存 = True Then
    
        myDoc.Save
        
    End If
    
    myDoc.Close
    
    myURL = ""
        
    Set myDoc = Nothing
    
    On Error GoTo 0
    
End Sub

Friend Property Get ファイルパス() As Variant

    ファイルフルパス = myURL

End Property

Friend Property Let ファイルパス(ByVal URL As Variant)

    Me.DocOpen URL

End Property

Friend Sub 文字置換(ByVal 検索文字 As String, ByVal 置換文字 As String, Optional ByVal 先頭検索 As Boolean = True, Optional ByVal 全文一致 As Boolean = False)

    Const wdReplaceAll = 2

    If mySel Is Nothing Then Exit Sub
    
    mySel.Find.ClearFormatting
    mySel.Find.Replacement.ClearFormatting

    mySel.Find.Text = 検索文字
    mySel.Find.Forward = 先頭検索
    mySel.Find.MatchWholeWord = 全文一致

    mySel.Find.Replacement.Text = 置換文字
    mySel.Find.Execute Replace:=wdReplaceAll

End Sub

Friend Function ファイル名文字置換(ByVal 検索文字 As String, ByVal 置換文字 As String) As String

    On Error Resume Next
    
    Dim 変更前ファイル名 As String
    Dim 変更後ファイル名 As String

    変更前ファイル名 = myURL

    If 変更前ファイル名 Like "*" & 検索文字 & "*" Then
    
        変更後ファイル名 = Replace(変更前ファイル名, 検索文字, 置換文字)
    
        Me.DocClose
            
        Name 変更前ファイル名 As 変更後ファイル名
        
        Me.ファイルパス = 変更後ファイル名
            
    Else
                
        変更後ファイル名 = 変更前ファイル名
                
    End If
    
    If Err.Number = 0 Then
    
        ファイル名文字置換 = 変更後ファイル名
        
    Else
    
        ファイル名文字置換 = ""
        
    End If
    
    On Error GoTo 0


End Function
