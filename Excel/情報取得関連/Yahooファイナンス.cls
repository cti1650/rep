VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Yahooファイナンス"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 正規表現 条件定義
Private Const 銘柄コード条件 As String = "code""\ value\=""[^""]*"
Private Const 銘柄コード条件置換 As String = "code"" value="""
Private Const 銘柄番号条件 As String = "<dt>[^<]*"
Private Const 銘柄番号条件置換 As String = "<dt>"
Private Const 株価詳細条件 As String = "[^>]*<\/strong><span\ class=""date\ yjSt"">"
Private Const 株価詳細条件置換 As String = "</strong><span class=""date yjSt"">"
Private Const 配当利回り条件 As String = "[^>]*<\/strong>%<span\ class=""date\ yjSt"">"
Private Const 配当利回り条件置換 As String = "</strong>%<span class=""date yjSt"">"
Private Const PS条件 As String = "[^>]*<\/a><\/strong><span\ class=""date\ yjSt"">"
Private Const PS条件置換 As String = "</a></strong><span class=""date yjSt"">"
Private Const 銘柄名条件 As String = "<h1>[^<]*"
Private Const 銘柄名条件置換 As String = "<h1>"
Private Const 業種条件 As String = "[^>]*<\/a><\/dd>"
Private Const 業種条件置換 As String = "</a></dd>"
Private Const 取引証券条件 As String = "stockMainTabName\"">[^<]*"
Private Const 取引証券条件置換 As String = "stockMainTabName"">"
Private Const 株価条件 As String = "stoksPrice"">[\d,.]*"
Private Const 株価条件置換 As String = "stoksPrice"">"
Private Const 時価総額条件 As String = "[\d,.]*<\/strong>百万円"
Private Const 時価総額条件置換 As String = "</strong>百万円"
Private Const 売買代金条件 As String = "[\d,.]*<\/strong>万円"
Private Const 売買代金条件置換 As String = "</strong>万円"
Private Const 倍率条件 As String = "[\d,.]*<\/strong>倍"
Private Const 倍率条件置換 As String = "</strong>倍"
Private Const 株数条件 As String = "[\d,.-]*<\/strong>株"
Private Const 株数条件置換 As String = "</strong>株"
Private Const 最低購入代金条件 As String = "0""><strong>[\d,.]*"
Private Const 最低購入代金条件置換 As String = "0""><strong>"
Private Const 会社PF情報条件 As String = "[^>]*<\/td>"
Private Const 会社PF情報条件置換 As String = "</td>"

Private 銘柄HP As String
Private 株主優待HP As String
Private 銘柄ニュースHP As String
Private 会社PFHP As String
Private 株価予報PBRHP As String
Private 株価予報PERHP As String

Private RE As Object
Private http As Object

Private Sub Class_Initialize()

    Set RE = CreateObject("VBScript.RegExp")
    DoEvents
    Set http = CreateObject("MSXML2.XMLHTTP.3.0")

End Sub

Private Sub Class_Terminate()

    Set RE = Nothing
    Set http = Nothing

End Sub

Public Function 銘柄設定(ByVal 銘柄コード As String) As Yahooファイナンス

    銘柄HP = getSite("http://stocks.finance.yahoo.co.jp/stocks/detail/?code=" & 銘柄コード)
    DoEvents
    株主優待HP = getSite("http://info.finance.yahoo.co.jp/stockholder/detail/?code=" & 銘柄コード)
    DoEvents
    銘柄ニュースHP = getSite("http://stocks.finance.yahoo.co.jp/stocks/news/?code=" & 銘柄コード)
    DoEvents
    会社PFHP = getSite("http://stocks.finance.yahoo.co.jp/stocks/profile/?code=" & 銘柄コード)
    DoEvents
    株価予報PBRHP = getSite("http://kabuyoho.ifis.co.jp/index.php?action=tp1&sa=report_pbr&bcode=" & 銘柄コード)
    DoEvents
    株価予報PERHP = getSite("http://kabuyoho.ifis.co.jp/index.php?action=tp1&sa=report_per&bcode=" & 銘柄コード)
    
    Set 銘柄設定 = Me

End Function

Private Function getSite(ByVal URL As String) As String

    'Dim http As Object
    'Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "GET", URL, False
    DoEvents
    http.Send

    ' ダウンロード待ち
    Do While http.ReadyState <> 4
        DoEvents
    Loop

    ' 結果出力
    getSite = http.responseText

    'Set http = Nothing

End Function

Private Function RegE(ByVal str As String, ByVal 参照文字 As String, ByVal 置換文字 As String, Optional ByVal No As Long = 0) As String

    'Dim RE As Variant
    Dim strPattern As String
    Dim Tes As Variant
    Dim buf As String
    'Set RE = CreateObject("VBScript.RegExp")
    strPattern = 参照文字
    With RE
        DoEvents
        .Pattern = strPattern       ''検索パターンを設定
        .IgnoreCase = True          ''大文字と小文字を区別しない
        .Global = True              ''文字列全体を検索
        Set Tes = .Execute(str)
        If Not Tes Is Nothing Then
        
            If Tes.Count > No Then
            
                buf = Tes.Item(No).Value
                
                buf = Replace$(buf, 置換文字, "")
                
                RegE = buf
                
                'Set RE = Nothing
                
                Exit Function

            End If
            
        End If
        
        RegE = ""
    End With
    'Set RE = Nothing
    
End Function

Public Function URL_月間銘柄チャート1ヶ月()

    URL_月間銘柄チャート1ヶ月 = "http://chart.yahoo.co.jp/?code=" & Me.銘柄コード & "&tm=1m&type=c&log=off&size=m&over=m65,m130,s,v&add=v&comp="

End Function

Public Function URL_年間銘柄チャート1年()

    URL_年間銘柄チャート1年 = "http://chart.yahoo.co.jp/?code=" & Me.銘柄コード & "&tm=1y&type=c&log=off&size=m&over=m65,m130,s,v&add=v&comp="

End Function

Public Function URL_年間銘柄チャート5年()

    URL_年間銘柄チャート5年 = "http://chart.yahoo.co.jp/?code=" & Me.銘柄コード & "&tm=5y&type=c&log=off&size=m&over=m65,m130,s,v&add=v&comp="

End Function

Public Function URL_年間日経平均株価チャート5年()

    URL_年間日経平均株価チャート5年 = "http://chart.yahoo.co.jp/?code=998407.O&tm=5y&type=c&log=off&size=m&over=m65,m130,s,v&add=&comp="

End Function

Public Function URL_年間NYダウチャート5年()

    URL_年間NYダウチャート5年 = "http://chart.yahoo.co.jp/?code=^DJI&tm=5y&type=c&log=off&size=m&over=m65,m130&add=&comp="

End Function

Public Function URL_年間ドル_円為替チャート5年()

    URL_年間ドル_円為替チャート5年 = "http://chart.yahoo.co.jp/?code=USDJPY=X&tm=5y&type=c&log=off&size=m&over=m130,m260&add=&comp="

End Function

Public Function 銘柄コード() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        銘柄コード = RegE(buf, 銘柄コード条件, 銘柄コード条件置換, 0)
        
        Exit Function
    
    End If
    
    銘柄コード = ""

End Function

Public Function 銘柄番号() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        銘柄番号 = RegE(buf, 銘柄番号条件, 銘柄番号条件置換, 0)
        
        Exit Function
    
    End If
    
    銘柄番号 = ""

End Function

Public Function 配当利回り() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        配当利回り = RegE(buf, 配当利回り条件, 配当利回り条件置換, 0)
        
        Exit Function
    
    End If
    
    配当利回り = ""

End Function

Public Function 前日終値() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        前日終値 = RegE(buf, 株価詳細条件, 株価詳細条件置換, 0)
        
        Exit Function
    
    End If
    
    前日終値 = ""

End Function

Public Function 始値() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        始値 = RegE(buf, 株価詳細条件, 株価詳細条件置換, 1)
        
        Exit Function
    
    End If
    
    始値 = ""

End Function

Public Function 高値() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        高値 = RegE(buf, 株価詳細条件, 株価詳細条件置換, 2)
        
        Exit Function
    
    End If
    
    高値 = ""

End Function

Public Function 安値() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        安値 = RegE(buf, 株価詳細条件, 株価詳細条件置換, 3)
        
        Exit Function
    
    End If
    
    安値 = ""

End Function


Public Function 値幅制限() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        値幅制限 = RegE(buf, 株価詳細条件, 株価詳細条件置換, 4)
        
        Exit Function
    
    End If
    
    値幅制限 = ""

End Function

Public Function 株価1株配当() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        株価1株配当 = RegE(buf, 株価詳細条件, 株価詳細条件置換, 5)
        
        Exit Function
    
    End If
    
    株価1株配当 = ""

End Function

Public Function 年初来高値() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        年初来高値 = RegE(buf, 株価詳細条件, 株価詳細条件置換, 9)
        
        Exit Function
    
    End If
    
    年初来高値 = ""

End Function

Public Function 年初来安値() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        年初来安値 = RegE(buf, 株価詳細条件, 株価詳細条件置換, 10)
        
        Exit Function
    
    End If
    
    年初来安値 = ""

End Function

Public Function EPS() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        EPS = RegE(buf, PS条件, PS条件置換, 0)
        
        Exit Function
    
    End If
    
    EPS = ""

End Function

Public Function BPS() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        BPS = RegE(buf, PS条件, PS条件置換, 1)
        
        Exit Function
    
    End If
    
    BPS = ""

End Function

Public Function 銘柄名() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        銘柄名 = RegE(buf, 銘柄名条件, 銘柄名条件置換, 0)
        
        Exit Function
    
    End If
    
    銘柄名 = ""

End Function

Public Function 業種() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        業種 = RegE(buf, 業種条件, 業種条件置換, 0)
        
        Exit Function
    
    End If
    
    業種 = ""

End Function

Public Function 取引証券() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        取引証券 = RegE(buf, 取引証券条件, 取引証券条件置換, 0)
        
        Exit Function
    
    End If
    
    取引証券 = ""

End Function

Public Function 株価() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        株価 = RegE(buf, 株価条件, 株価条件置換, 0)
        
        Exit Function
    
    End If
    
    株価 = ""

End Function

Public Function 時価総額() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        時価総額 = RegE(buf, 時価総額条件, 時価総額条件置換, 0)
        
        Exit Function
    
    End If
    
    時価総額 = ""

End Function

Public Function 売買代金() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        売買代金 = RegE(buf, 売買代金条件, 売買代金条件置換, 0)
        
        Exit Function
    
    End If
    
    売買代金 = ""

End Function

Public Function PER() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        PER = RegE(buf, 倍率条件, 倍率条件置換, 0)
        
        Exit Function
    
    End If
    
    PER = ""

End Function

Public Function PBR() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        PBR = RegE(buf, 倍率条件, 倍率条件置換, 1)
        
        Exit Function
    
    End If
    
    PBR = ""

End Function

Public Function 貸借倍率() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        貸借倍率 = RegE(buf, 倍率条件, 倍率条件置換, 2)
        
        Exit Function
    
    End If
    
    貸借倍率 = ""

End Function

Public Function 出来高() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        出来高 = RegE(buf, 株数条件, 株数条件置換, 0)
        
        Exit Function
    
    End If
    
    出来高 = ""

End Function

Public Function 発行済株式数() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        発行済株式数 = RegE(buf, 株数条件, 株数条件置換, 1)
        
        Exit Function
    
    End If
    
    発行済株式数 = ""

End Function

Public Function 単元株数() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        単元株数 = RegE(buf, 株数条件, 株数条件置換, 2)
        
        Exit Function
    
    End If
    
    単元株数 = ""

End Function

Public Function 信用買残() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        信用買残 = RegE(buf, 株数条件, 株数条件置換, 3)
        
        Exit Function
    
    End If
    
    信用買残 = ""

End Function

Public Function 信用買残前週比() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        信用買残前週比 = RegE(buf, 株数条件, 株数条件置換, 4)
        
        Exit Function
    
    End If
    
    信用買残前週比 = ""

End Function

Public Function 信用売残() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        信用売残 = RegE(buf, 株数条件, 株数条件置換, 5)
        
        Exit Function
    
    End If
    
    信用売残 = ""

End Function

Public Function 信用売残前週比() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        信用売残前週比 = RegE(buf, 株数条件, 株数条件置換, 6)
        
        Exit Function
    
    End If
    
    信用売残前週比 = ""

End Function

Public Function 最低購入代金() As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" Then
    
        最低購入代金 = RegE(buf, 最低購入代金条件, 最低購入代金条件置換, 13)
        
        Exit Function
    
    End If
    
    最低購入代金 = ""

End Function

Public Function 企業特色() As String

    Dim buf As String
    
    buf = 会社PFHP
    
    If buf <> "" Then
    
        企業特色 = RegE(buf, 会社PF情報条件, 会社PF情報条件置換, 3)
        
        Exit Function
    
    End If
    
    企業特色 = ""

End Function

Public Function 企業連結事業() As String

    Dim buf As String
    
    buf = 会社PFHP
    
    If buf <> "" Then
    
        企業連結事業 = RegE(buf, 会社PF情報条件, 会社PF情報条件置換, 4)
        
        Exit Function
    
    End If
    
    企業連結事業 = ""

End Function

Public Function 企業電話番号() As String

    Dim buf As String
    
    buf = 会社PFHP
    
    If buf <> "" Then
    
        企業電話番号 = RegE(buf, 会社PF情報条件, 会社PF情報条件置換, 7)
        
        Exit Function
    
    End If
    
    企業電話番号 = ""

End Function

Public Function 企業英文社名() As String

    Dim buf As String
    
    buf = 会社PFHP
    
    If buf <> "" Then
    
        企業英文社名 = RegE(buf, 会社PF情報条件, 会社PF情報条件置換, 9)
        
        Exit Function
    
    End If
    
    企業英文社名 = ""

End Function

Public Function 企業代表者名() As String

    Dim buf As String
    
    buf = 会社PFHP
    
    If buf <> "" Then
    
        企業代表者名 = RegE(buf, 会社PF情報条件, 会社PF情報条件置換, 10)
        
        Exit Function
    
    End If
    
    企業代表者名 = ""

End Function

Public Function 企業設立年月日() As String

    Dim buf As String
    
    buf = 会社PFHP
    
    If buf <> "" Then
    
        企業設立年月日 = RegE(buf, 会社PF情報条件, 会社PF情報条件置換, 11)
        
        Exit Function
    
    End If
    
    企業設立年月日 = ""

End Function

Public Function 企業市場名() As String

    Dim buf As String
    
    buf = 会社PFHP
    
    If buf <> "" Then
    
        企業市場名 = RegE(buf, 会社PF情報条件, 会社PF情報条件置換, 12)
        
        Exit Function
    
    End If
    
    企業市場名 = ""

End Function

Public Function 企業上場年月日() As String

    Dim buf As String
    
    buf = 会社PFHP
    
    If buf <> "" Then
    
        企業上場年月日 = RegE(buf, 会社PF情報条件, 会社PF情報条件置換, 13)
        
        Exit Function
    
    End If
    
    企業上場年月日 = ""

End Function

Public Function 企業決算() As String

    Dim buf As String
    
    buf = 会社PFHP
    
    If buf <> "" Then
    
        企業決算 = RegE(buf, 会社PF情報条件, 会社PF情報条件置換, 14)
        
        Exit Function
    
    End If
    
    企業決算 = ""

End Function

Public Function 企業従業員数() As String

    Dim buf As String
    
    buf = 会社PFHP
    
    If buf <> "" Then
    
        企業従業員数 = RegE(buf, 会社PF情報条件, 会社PF情報条件置換, 16)
        
        Exit Function
    
    End If
    
    企業従業員数 = ""

End Function

Public Function 企業平均年齢() As String

    Dim buf As String
    
    buf = 会社PFHP
    
    If buf <> "" Then
    
        企業平均年齢 = RegE(buf, 会社PF情報条件, 会社PF情報条件置換, 18)
        
        Exit Function
    
    End If
    
    企業平均年齢 = ""

End Function

Public Function 企業平均年収() As String

    Dim buf As String
    
    buf = 会社PFHP
    
    If buf <> "" Then
    
        企業平均年収 = RegE(buf, 会社PF情報条件, 会社PF情報条件置換, 19)
        
        Exit Function
    
    End If
    
    企業平均年収 = ""

End Function

Public Function 企業本社所在地() As String

    Dim buf As String
    
    buf = 会社PFHP
    
    If buf <> "" Then
    
        企業本社所在地 = RegE(buf, "〒[^\[]*", "", 0)
        
        Exit Function
    
    End If
    
    企業本社所在地 = ""

End Function

Public Function 上値目途_PBRアプローチ() As String

    Dim buf As String
    
    buf = 株価予報PBRHP
    
    If buf <> "" Then
    
        上値目途_PBRアプローチ = RegE(buf, "medo_value"">[^<]*", "medo_value"">", 0)
        
        Exit Function
    
    End If
    
    上値目途_PBRアプローチ = ""

End Function

Public Function 妥当株価_PBRアプローチ() As String

    Dim buf As String
    
    buf = 株価予報PBRHP
    
    If buf <> "" Then
    
        妥当株価_PBRアプローチ = RegE(buf, "medo_value"">[^<]*", "medo_value"">", 1)
        
        Exit Function
    
    End If
    
    妥当株価_PBRアプローチ = ""

End Function

Public Function 下値目途_PBRアプローチ() As String

    Dim buf As String
    
    buf = 株価予報PBRHP
    
    If buf <> "" Then
    
        下値目途_PBRアプローチ = RegE(buf, "medo_value"">[^<]*", "medo_value"">", 2)
        
        Exit Function
    
    End If
    
    下値目途_PBRアプローチ = ""

End Function

Public Function 上値目途_PERアプローチ() As String

    Dim buf As String
    
    buf = 株価予報PERHP
    
    If buf <> "" Then
    
        上値目途_PERアプローチ = RegE(buf, "medo_value"">[^<]*", "medo_value"">", 0)
        
        Exit Function
    
    End If
    
    上値目途_PERアプローチ = ""

End Function

Public Function 妥当株価_PERアプローチ() As String

    Dim buf As String
    
    buf = 株価予報PERHP
    
    If buf <> "" Then
    
        妥当株価_PERアプローチ = RegE(buf, "medo_value"">[^<]*", "medo_value"">", 1)
        
        Exit Function
    
    End If
    
    妥当株価_PERアプローチ = ""

End Function

Public Function 下値目途_PERアプローチ() As String

    Dim buf As String
    
    buf = 株価予報PERHP
    
    If buf <> "" Then
    
        下値目途_PERアプローチ = RegE(buf, "medo_value"">[^<]*", "medo_value"">", 2)
        
        Exit Function
    
    End If
    
    下値目途_PERアプローチ = ""

End Function

Public Function 銘柄関連株(Optional ByVal No As Long = 1, Optional ByVal ハイパーリンク As Boolean = False) As String

    Dim buf As String
    
    buf = 銘柄HP
    
    If buf <> "" And No <= 15 And No > 0 Then
    
        If ハイパーリンク = True Then
        
            If No > 0 And No < 6 Then
            
                銘柄関連株 = "=HYPERLINK(""" & RegE(buf, "<\/em><\/span><a\ href=""[^""]*", "</em></span><a href=""", No - 1) & """,""" & _
                                RegE(buf, "[^>]*<\/a><\/li><li", "</a></li><li", No) & """)"
                                
            Else
            
                銘柄関連株 = "=HYPERLINK(""" & RegE(buf, "<\/em><\/span><a\ href=""[^""]*", "</em></span><a href=""", No - 1) & """,""" & _
                                RegE(buf, "[^>]*<\/a><\/li>\n", "</a></li>" & vbLf, 27 + (No - 1)) & """)"
            
            End If
            
        Else
        
            If No > 0 And No < 6 Then
            
                銘柄関連株 = RegE(buf, "[^>]*<\/a><\/li><li", "</a></li><li", No)
                                
            Else
            
                銘柄関連株 = RegE(buf, "[^>]*<\/a><\/li>\n", "</a></li>" & vbLf, 27 + (No - 1))
            
            End If
        
        End If
        
        Exit Function
    
    End If
    
    銘柄関連株 = ""

End Function

Public Function 銘柄関連ニュース(Optional ByVal No As Long = 1, Optional ByVal ハイパーリンク As Boolean = False) As String

    Dim buf As String
    
    buf = 銘柄ニュースHP
    
    If buf <> "" And No <= 10 And No > 0 Then
    
        If ハイパーリンク = True Then
        
            銘柄関連ニュース = "=HYPERLINK(""" & RegE(buf, "ymuiArrow1""><a\ href=""[^""]*", "ymuiArrow1""><a href=""", No) & """,""" & _
                                RegE(buf, "[^>]*<\/a><span\ c", "</a><span c", No - 1) & " " & RegE(buf, "[^>]*<\/span><\/li>", "</span></li>", No - 1) & """)"
            
        Else
        
            銘柄関連ニュース = RegE(buf, "[^>]*<\/a><span\ c", "</a><span c", No - 1) & " " & RegE(buf, "[^>]*<\/span><\/li>", "</span></li>", No - 1)
        
        End If
        
        Exit Function
    
    End If
    
    銘柄関連ニュース = ""

End Function

Public Function 株主優待(Optional ByVal No As Long = 1, Optional ByVal ハイパーリンク As Boolean = False) As String

    Dim buf As String
    
    buf = 株主優待HP
    
    If buf <> "" And No < 10 Then
    
        If ハイパーリンク = True Then
        
            株主優待 = "=HYPERLINK(""" & "http://info.finance.yahoo.co.jp/stockholder/detail/?code=" & Me.銘柄コード & """,""" & _
                                Replace(RegE(buf, "\n<div\ class\=""content"">\n<h3>[^<]*<\/h3>", vbLf & "<div class=""content"">" & vbLf & "<h3>", No - 1), "</h3>", "") & """)"
            
        Else
        
            株主優待 = Replace(RegE(buf, "\n<div\ class\=""content"">\n<h3>[^<]*<\/h3>", vbLf & "<div class=""content"">" & vbLf & "<h3>", No - 1), "</h3>", "")
        
        End If
        
        Exit Function
    
    End If
    
    株主優待 = ""

End Function
