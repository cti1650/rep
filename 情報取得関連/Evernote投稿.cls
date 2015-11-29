VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Evernote投稿"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private clsReminder As String
Private clsNotebook As String
Private clsTag As String
Private clsTitle As String
Private clsText As String

Private Sub Class_Initialize()
    
    Me.クリア
    
End Sub

Private Sub Class_Terminate()

    Me.クリア

End Sub


Public Sub クリア()
    
    clsReminder = ""
    clsNotebook = ""
    clsTag = ""
    clsTitle = ""
    clsText = ""

End Sub

Public Function リマインダー(ByVal 日付 As String) As String ' today or 2015/5/31

    Select Case 日付
    Case "ToDo"
        clsReminder = "today"
        clsTag = "ToDo"
    Case "今日"
        clsReminder = Format(Now, "yyyy/mm/dd")
    Case "明日"
        clsReminder = Format(Now + 1, "yyyy/mm/dd")
    Case "明後日"
        clsReminder = Format(Now + 2, "yyyy/mm/dd")
    Case "来週"
        clsReminder = Format(Now + 7, "yyyy/mm/dd")
    Case "再来週"
        clsReminder = Format(Now + 14, "yyyy/mm/dd")
    Case Else
        clsReminder = 日付
    End Select
    
    リマインダー = clsReminder

End Function

Public Function ノートブック(ByVal ノートブック名 As String) As String

    clsNotebook = ノートブック名
    
    ノートブック = clsNotebook

End Function


Public Function タグ(ParamArray タグ名() As Variant) As String

    Dim buf As String
    
    If IsArray(タグ名) = False Then
    
        buf = ""
    
    Else
    
        buf = Join$(タグ名, ",")
        
    End If

    clsTag = buf
    
    タグ = clsTag

End Function

Public Function ノート作成(ByVal タイトル As String, ByVal 本文 As String, Optional ByVal 追記 As Boolean = False) As Boolean

On Error Resume Next

Dim http As Object
Dim URL As String
Dim URLText As String

URL = "https://script.google.com/macros/s/AKfycbwzXqMLaDKWDPrpnMTdqIkoWB7KaNz7DQ3NlE3tafDCH_icU_c/exec?"

If タイトル <> "" Then

    Err.Clear
    Set http = CreateObject("MSXML2.XMLHTTP")
    If Err.Number <> 0 Then
        Set http = CreateObject("MSXML.XMLHTTPRequest")
    End If
    If http Is Nothing Then
        ノート作成 = False
        Exit Function
    End If

    URLText = URL & makeURL(タイトル, 追記) & "&value=" & UrlEncode(本文)

    Err.Clear

    http.Open "GET", URLText, False
    http.Send
    
    If Err.Number <= 0 Then
        ノート作成 = True
    Else
        ノート作成 = False
    End If
    
    Set http = Nothing
    
Else

    ノート作成 = False

End If

On Error GoTo 0


End Function

Private Function makeURL(ByVal Title As String, ByVal 追記 As Boolean) As String

    Dim buf As String

    buf = ""
    
    If Title <> "" Then
    
        buf = "title=" & UrlEncode(Title)
        
        If clsReminder <> "" Then
        
            buf = buf & "&reminder=" & UrlEncode(clsReminder)
        
        End If
        
        If clsNotebook <> "" Then
        
            buf = buf & "&notebook=" & UrlEncode(clsNotebook)
        
        End If
        
        If clsTag <> "" Then
        
            buf = buf & "&tag=" & UrlEncode(clsTag)
        
        End If
        
        If 追記 = True Then
        
            buf = buf & "&add=true"
            
        End If
    
    End If
    
    makeURL = buf

End Function

Private Function URLChk(ByVal URLText As String) As Boolean

    If URLText Like "*[^<>#""%{}|\^[]`;/?:@&=+$,'()]*" Then
    
        URLChk = False
    
    Else
    
        URLChk = True
    
    End If

End Function


' 32bitOSの場合のみ実行可能なURLエンコード関数
Private Function UrlEncode32(ByRef strSource As String) As String

    With CreateObject("ScriptControl")
        .Language = "Jscript"
        UrlEncode = .CodeObject.encodeURI(strSource)
    End With
    
End Function

' 32,64bitOS共に実行可能なURLエンコード関数
Private Function UrlEncode(ByVal strSource As String) As String

  Dim d As Object
  Dim elm As Object
  
  sWord = Replace(strSource, "\", "\\")
  sWord = Replace(sWord, "'", "\'")
  sWord = Replace(sWord, vbCrLf, "\n")
  sWord = Replace(sWord, vbCr, "\n")
  sWord = Replace(sWord, vbLf, "\n")
  Set d = CreateObject("htmlfile")
  Set elm = d.createElement("span")
  elm.setAttribute "id", "result"
  d.appendChild elm
  d.parentWindow.execScript "document.getElementById('result').innerText = encodeURIComponent('" & sWord & "');", "JScript"
  
  UrlEncode = elm.innerText
  
End Function
