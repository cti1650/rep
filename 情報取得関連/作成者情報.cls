VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "作成者情報"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private clsJSON As Object

Private Sub Class_Initialize()
    Dim URL_ As String
    URL_ = "https://script.google.com/macros/s/AKfycbzoRJCMurS5MCnxyS2YzJK4cs4hzbg5ADdhCq5R4v8eMDiMcnqC/exec?user=y.sakagami&pass=0280200"
    Set clsJSON = jsonParse(getHTTP(URL_))
End Sub

Private Sub Class_Terminate()
    Set clsJSON = Nothing
End Sub

Private Function jsonParse(ByVal json As String) As Object

    Dim objSC      As Object 'Script Control
    Dim strFunc As String '関数文字列
 
    Dim strJSON As String 'JSONデータ(文字列)
    Dim objJSON As Object 'JSONファイルをパースしたもの
    Set objSC = CreateObject("ScriptControl")
 
    objSC.Language = "JScript"
    strFunc = "function jsonParse(s) { return eval('(' + s + ')'); }"
    objSC.AddCode strFunc
 
    strJSON = json
    Set jsonParse = objSC.CodeObject.jsonParse(strJSON)

End Function

Private Function getHTTP(ByVal URL As String) As String

    On Error Resume Next

    Dim http As Object
    Err.Clear
    Set http = CreateHttpRequest

    Err.Clear

    http.SetOption 2, http.GetOption(2) - SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
    http.Open "GET", URL, False
    http.setRequestHeader "Pragma", "no-cache"
    http.setRequestHeader "Cache-Control", "no-cache"
    http.setRequestHeader "If-Modified-Since", "Thu, 01 Jun 1970 00:00:00 GMT"

    http.Send

    ' ダウンロード待ち
    Do While http.readyState <> 4
        DoEvents
    Loop
    
    ' 結果出力
    If Err.Number <= 0 Then
        getHTTP = http.responseText
    Else
        getHTTP = ""
    End If

    Set http = Nothing
    
    On Error GoTo 0

End Function

Private Function CreateHttpRequest() As Object
'WinHttpRequest/XMLHTTPRequestオブジェクト作成
'http://www.f3.dion.ne.jp/~element/msaccess/AcTipsWinHTTP1.html 参考
  Dim progIDs As Variant
  Dim ret As Object
  Dim i As Long
  
  Set ret = Nothing '初期化
  progIDs = Array("MSXML2.ServerXMLHTTP", _
                  "Msxml2.XMLHTTP.6.0", _
                  "Msxml2.XMLHTTP.5.0", _
                  "Msxml2.XMLHTTP.4.0", _
                  "Msxml2.XMLHTTP.3.0", _
                  "Msxml2.XMLHTTP", _
                  "Microsoft.XMLHTTP")
  On Error Resume Next
  For i = LBound(progIDs) To UBound(progIDs)
    Set ret = CreateObject(progIDs(i))
    If Not ret Is Nothing Then Exit For
  Next
  On Error GoTo 0
  Set CreateHttpRequest = ret
End Function

Private Function 項目有効確認(ByRef cls As Variant) As variant

    On Error Resume Next

    If Not clsJSON Is Nothing Then

        項目有効確認 = cls
        
    Else
    
        項目有効確認 = "取得失敗"
    
    End If

    On Error Goto 0

    If Err.Number <> 0 Then 項目有効確認 = "取得失敗"

End Function

' ↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓


Public Function 会社名() As String

    会社名 = 項目有効確認(clsJSON.corp)

End Function

Public Function 部署() As String

　　部署 = 項目有効確認(clsJSON.unit)

End Function

Public Function 氏名() As String

　　氏名 = 項目有効確認(clsJSON.builder)

End Function

Public Function 連絡先() As String

　　連絡先 = 項目有効確認(clsJSON.tel)

End Function

Public Function 携帯番号() As String

　　携帯番号 = 項目有効確認(clsJSON.mobile)

End Function

Public Function アドレス() As String

　　アドレス = 項目有効確認(clsJSON.mail)

End Function

Public Function 使用PC() As String

　　使用PC = 項目有効確認(clsJSON.pc)

End Function

Public Function 使用IP() As String

　　使用IP = 項目有効確認(clsJSON.pcip)

End Function
