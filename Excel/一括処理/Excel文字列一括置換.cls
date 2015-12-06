VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Excel文字列一括置換"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private myExcel As Object
Private myBk As Object
Private mySel As Object
Private myURL As String

Private Sub Class_Initialize()

    Application.DisplayAlerts = False

    Set myExcel = createExcel
    
    myURL = ""

End Sub

Private Sub Class_Terminate()

    myExcel.Quit
    
    myURL = ""

    Set mySel = Nothing
    Set myBk = Nothing
    Set myExcel = Nothing
    
    Application.DisplayAlerts = True

End Sub

Private Function createExcel()

    On Error Resume Next
    
    Dim objExcel As Object

    Err.Clear

    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = False
    
    If Err.Number = 0 Then
    
        Set createExcel = objExcel
        
    Else
    
        Set createExcel = Nothing
    
    End If
    
    On Error GoTo 0

End Function

Friend Function BkOpen(ByVal ファイル名 As String)

    On Error Resume Next

    Err.Clear

    If myExcel Is Nothing Then Exit Function
    
    If Not ファイル名 Like "*\*" Then
    
        ファイル名 = ThisWorkbook.Path & "\" & ファイル名

    End If
    
    Set myBk = myExcel.Workbooks.Open(ファイル名)
    
    If Err.Number = 0 Then
    
        myURL = ファイル名
    
        Set BkOpen = myBk
        
    Else
    
        myURL = ""
    
        Set BkOpen = Nothing
    
    End If
    
    On Error GoTo 0
    
End Function

Friend Sub BkClose(Optional ByVal 保存 As Boolean = True)

    On Error Resume Next

    Err.Clear

    If myBk Is Nothing Then Exit Sub
    
    Set mySel = Nothing
    
    If 保存 = True Then
    
        myBk.Save
        
    End If
    
    myBk.Close
    
    myURL = ""
        
    Set myBk = Nothing
    
    On Error GoTo 0
    
End Sub

Friend Property Get ファイルパス() As Variant

    ファイルフルパス = myURL

End Property

Friend Property Let ファイルパス(ByVal URL As Variant)

    Me.BkOpen URL

End Property

Friend Sub 文字置換(ByVal 検索文字 As String, ByVal 置換文字 As String)

    On Error Resume Next

    Err.Clear

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

    On Error GoTo 0

End Sub

Friend Function ファイル名文字置換(ByVal 検索文字 As String, ByVal 置換文字 As String) As String

    On Error Resume Next
    
    Dim 変更前ファイル名 As String
    Dim 変更後ファイル名 As String

    変更前ファイル名 = myURL

    If 変更前ファイル名 Like "*" & 検索文字 & "*" Then
    
        変更後ファイル名 = Replace(変更前ファイル名, 検索文字, 置換文字)
    
        Me.BkClose
            
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
