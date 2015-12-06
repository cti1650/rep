VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Officeテンプレート"
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
Private myCol As Object

Private 原紙フォルダ As String
Private 保存先フォルダ As String
Private 対象ファイル条件 As String

Private Const 基準格納フォルダ As String = "生成データ"
Private Const ErrCheck As Boolean = False
Private Const 条件定義 As String = "*.????"




Private Sub Class_Initialize()

#If ErrCheck = False Then

    Application.DisplayAlerts = False

#Else

    Application.DisplayAlerts = True

#End If

    Set myCol = CreateObject("Scripting.Dictionary")
    Set myExcel = createExcel
    Set myWord = createWord
    With myExcel
        myExcel.ScreenUpdating = False
    End With
    With myWord
        myWord.ScreenUpdating = False
    End With
    
    myURL = ""
    原紙フォルダ = ThisWorkbook.Path
    保存先フォルダ = ThisWorkbook.Path & "\" & 基準格納フォルダ
    対象ファイル条件 = 条件定義
    
End Sub

Private Sub Class_Terminate()

    With myExcel
        myExcel.ScreenUpdating = True
    End With
    With myWord
        myWord.ScreenUpdating = True
    End With

    myExcel.Quit
    myWord.Quit
    
    myURL = ""
    原紙フォルダ = ""
    保存先フォルダ = ""

    Set mySel = Nothing
    Set myBk = Nothing
    Set myExcel = Nothing
    Set myDoc = Nothing
    Set myWord = Nothing
    Set myCol = Nothing
    
    Application.DisplayAlerts = True

End Sub




Friend Function 原紙(Optional ByVal パス As String = "") As Officeテンプレート

    On Error Resume Next
    
    Dim ErrFlag As Boolean
    
    ErrFlag = False
    
    

    If パス <> "" Then
    
        Select Case True
        Case パス Like "[A-Z\][\:]*"
        Case パス Like "[A-Z\][\:]*[!.]*"
            パス = Replace(パス, Dir(パス, vbNormal), "")
        Case パス Like "\*"
            パス = ThisWorkbook.Path & パス
        Case Else
            パス = ThisWorkbook.Path & "\" & パス
        End Select
        If パス Like "*\" Then
            パス = Left(パス, Len(パス) - 1)
        End If
        原紙フォルダ = パス
      
    Else
    
        原紙フォルダ = ThisWorkbook.Path
    
    End If
    
    
    
    If Err.Number = 0 Then

        Set 原紙 = Me

    Else
    
        Set 原紙 = Me
    
    End If
    
    On Error GoTo 0

End Function

Friend Function ファイル検索条件(Optional ByVal 条件 As String = "*.????") As Officeテンプレート

    On Error Resume Next
    
    Dim ErrFlag As Boolean
    
    ErrFlag = False
    
    

    If 条件 <> "" Then
    
        If 条件 Like "*\" Then
            条件 = Left(条件, Len(条件) - 1)
        End If
        If 条件 Like "\*" Then
            条件 = Mid(条件, 2, Len(条件) - 1)
        End If
        対象ファイル条件 = 条件
      
    Else
    
        対象ファイル条件 = 条件定義
    
    End If
    
    
    
    If Err.Number = 0 Then

        Set ファイル検索条件 = Me

    Else
    
        Set ファイル検索条件 = Me
    
    End If
    
    On Error GoTo 0

End Function

Friend Function 複製(Optional ByVal パス As String = "") As String


    On Error Resume Next
    
    Dim ErrFlag As Boolean
    Dim 作成フォルダ名 As String
    
    ErrFlag = False
    
    作成フォルダ名 = ""

    If パス <> "" Then
    
        Select Case True
        Case パス Like "[A-Z\][\:]*"
        Case パス Like "[A-Z\][\:]*[!.]*"
            パス = Replace(パス, Dir(パス, vbNormal), "")
        Case パス Like "\*"
            パス = ThisWorkbook.Path & "\" & 基準格納フォルダ & パス
        Case Else
            パス = ThisWorkbook.Path & "\" & 基準格納フォルダ & "\" & パス
        End Select
        If パス Like "*\" Then
            パス = Left(パス, Len(パス) - 1)
        End If
        保存先フォルダ = パス
        
        作成フォルダ名 = ファイル検索(原紙フォルダ, 原紙フォルダ)
      
    Else
    
        保存先フォルダ = ThisWorkbook.Path & "\" & 基準格納フォルダ
        
        作成フォルダ名 = ファイル検索(原紙フォルダ, 原紙フォルダ)
    
    End If
    
    Me.項目全削除
    
    If Err.Number = 0 Then

        複製 = 作成フォルダ名

    Else
    
        複製 = ""
    
    End If
    
    On Error GoTo 0

End Function

Friend Function 項目追加(ByVal 項目名 As String, ByVal 値 As Variant, Optional ByVal 書式 As String = "", Optional ByVal 全半角 As String = "", Optional ByVal 大文字小文字 As String = "") As Boolean

    On Error Resume Next
    
    Dim ErrFlag As Boolean
    
    ErrFlag = False
    
    

    If 項目名 <> "" Then
    
        If Not myCol Is Nothing Then
        
            If myCol.Exists(項目名) = False Then
            
                If 書式 <> "" Then
                        
                    値 = Format(値, 書式)
                            
                End If
                        
                Select Case 全半角
                Case "全角"
                    値 = StrConv(値, vbWide)
                Case "半角"
                    値 = StrConv(値, vbNarrow)
                End Select
                
                Select Case 大文字小文字
                Case "大文字"
                    値 = UCase(値)
                Case "小文字"
                    値 = LCase(値)
                Case "先頭大文字"
                    値 = StrConv(値, vbProperCase)
                End Select
                
                myCol.Add 項目名, 値
                ErrFlag = True
                
            End If
            
        End If
        
    End If
    
    
    
    If Err.Number = 0 Then

        項目追加 = ErrFlag

    Else
    
        項目追加 = False
    
    End If
    
    On Error GoTo 0

End Function

Friend Function 項目変更(ByVal 項目名 As String, ByVal 値 As Variant) As Boolean

    On Error Resume Next
    
    Dim ErrFlag As Boolean
    
    ErrFlag = False
    
    

    If 項目名 <> "" Then
    
        If Not myCol Is Nothing Then
        
            If myCol.Exists(項目名) = True Then
                myCol.Remove 項目名
                myCol.Add 項目名, 値
                ErrFlag = True
            Else
                myCol.Add 項目名, 値
                ErrFlag = True
            End If
            
        End If
        
    End If
    
    
    
    If Err.Number = 0 Then

        項目変更 = ErrFlag

    Else
    
        項目変更 = False
    
    End If
    
    On Error GoTo 0

End Function

Friend Function 項目削除(ByVal 項目名 As String) As Boolean

    On Error Resume Next
    
    Dim ErrFlag As Boolean
    
    ErrFlag = False
    
    

    If 項目名 <> "" Then
    
        If Not myCol Is Nothing Then
        
            If myCol.Exists(項目名) = True Then
                myCol.Remove 項目名
                ErrFlag = True
            End If
            
        End If
        
    End If
    
    
    
    If Err.Number = 0 Then

        項目削除 = ErrFlag

    Else
    
        項目削除 = False
    
    End If
    
    On Error GoTo 0

End Function

Friend Function 項目全削除() As Boolean

    On Error Resume Next
    
    Dim ErrFlag As Boolean
    
    ErrFlag = False



    If Not myCol Is Nothing Then
        
        myCol.RemoveAll
        ErrFlag = True
        
    Else
        
        Set myCol = CreateObject("Scripting.Dictionary")
        ErrFlag = True
    
    End If
    
    
    
    If Err.Number = 0 Then

        項目全削除 = ErrFlag

    Else
    
        項目全削除 = False
    
    End If
    
    On Error GoTo 0

End Function


'=====================================================================================================================================
'=====================================================================================================================================
'=====================================================================================================================================
'=====================================================================================================================================


Private Function ファイル検索(ByVal 参照パス As String, Optional ByVal 基準パス As String = "") As String

    On Error Resume Next
    
    Err.Clear

    Dim buf As String, f As Object
    Dim FullPath As String
    Dim 出力フォルダ As String
    Dim 作成フォルダ As String
    Dim 出力ファイル As String

    出力フォルダ = 保存先フォルダ
    作成フォルダ = 文字列置換(出力フォルダ & Replace(参照パス, 基準パス, ""))
    作成フォルダ = フォルダ作成(作成フォルダ)
    buf = Dir(参照パス & "\" & 対象ファイル条件)
    
    Do While buf <> ""
        FullPath = 参照パス & "\" & buf
        出力ファイル = ファイルコピー(作成フォルダ, FullPath, buf)
        If 出力ファイル <> "" Then
            Call ファイル内容変更(出力ファイル)
        End If
        buf = Dir()
    Loop
    
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(参照パス).SubFolders
            Call ファイル検索(f.Path, 基準パス)
        Next f
    End With
    
    If Err.Number = 0 Then
    
        ファイル検索 = 作成フォルダ
        
    Else
    
        ファイル検索 = ""
    
    End If
    
    On Error GoTo 0
    
End Function

Private Function ファイル内容変更(ByVal ファイル名 As String) As Boolean

    
    On Error Resume Next
    
    Dim buf As String
    Dim i As Long
    Dim ArrKey As Variant
    Dim ArrItem As Variant
    
    buf = ファイル名

    If buf <> "" Then
    
        Call 開く(ファイル名)

        If Not myCol Is Nothing Then
        
            If myCol.count > 0 Then
        
                ArrKey = myCol.keys
                ArrItem = myCol.items
        
                For i = LBound(ArrKey) To UBound(ArrKey)
                    Call 文字置換(ArrKey(i), ArrItem(i))
                Next
                    
            End If
    
    
        End If
        
        Call 閉じる
        
    End If
    
    If Err.Number = 0 Then

        ファイル内容変更 = True

    Else
    
        ファイル内容変更 = False
    
    End If
    
    On Error GoTo 0

End Function

' OK
Private Function フォルダ作成(ByVal フォルダ名 As String) As String

    Select Case True
    Case Not フォルダ名 Like "?[\:]*"
        フォルダ名 = ThisWorkbook.Path & "\" & フォルダ名 & "\"
    Case Not フォルダ名 Like "*\"
        フォルダ名 = フォルダ名 & "\"
    Case フォルダ名 Like "*\*\"
    End Select
    
    フォルダ名 = 使用禁止文字除外(フォルダ名)
    
    If MakeSureDirectoryPathExists(フォルダ名) Then
    
        フォルダ作成 = フォルダ名
        
    Else
    
        フォルダ作成 = ""
        
    End If

End Function

Private Function ファイルコピー(ByVal コピー先フォルダ As String, ByVal 絶対パス As String, ByVal ファイル名 As String) As String

    On Error Resume Next

    Dim FSO    As Object
    Dim FilePath As String
    
    Err.Clear
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    FilePath = コピー先フォルダ & "\" & ファイル名
    
    FilePath = 使用禁止文字除外(文字列置換(FilePath))
    
    FSO.CopyFile 絶対パス, FilePath
    
    Set FSO = Nothing
    
    If Err.Number = 0 Then
    
        ファイルコピー = FilePath
        
    Else
        
        ファイルコピー = ""
    
    End If
    
    On Error GoTo 0

End Function

Private Function 使用禁止文字除外(ByVal 値 As String) As String

    On Error Resume Next
    
    Dim buf As String
    Dim i As Long
    Dim ArrKey As Variant
    Dim ArrItem As Variant
    
    ArrKey = Array("|", ",", "*", "?", "<", ">", vbCr, vbLf)
    ArrItem = Array("｜", "，", "＊", "？", "＜", "＞", "", "")
    
    buf = 値

    If buf Like "*[|,*?<>]*" Or buf Like "*" & vbCr & "*" Or buf Like "*" & vbLf & "*" Then

        For i = LBound(ArrKey) To UBound(ArrKey)
        
            buf = Replace(buf, ArrKey(i), ArrItem(i))
                    
        Next
        
    End If
    
    If Err.Number = 0 Then

        使用禁止文字除外 = buf

    Else
    
        使用禁止文字除外 = 値
    
    End If
    
    On Error GoTo 0

End Function


Private Function 文字列置換(ByVal 値 As String) As String

    On Error Resume Next
    
    Dim buf As String
    Dim i As Long
    Dim ArrKey As Variant
    Dim ArrItem As Variant
    
    buf = 値

    If buf <> "" Then

        If Not myCol Is Nothing Then
        
            If myCol.count > 0 Then
        
                ArrKey = myCol.keys
                ArrItem = myCol.items
        
                For i = LBound(ArrKey) To UBound(ArrKey)
                    buf = Replace(buf, ArrKey(i), ArrItem(i))
                Next
                    
            End If
    
    
        End If
        
    End If
    
    If Err.Number = 0 Then

        文字列置換 = buf

    Else
    
        文字列置換 = 値
    
    End If
    
    On Error GoTo 0

End Function


' OK
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

' OK
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

' OK ?
Private Function 開く(ByVal ファイル名 As String)

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

' OK ?
Private Sub 閉じる(Optional ByVal 保存 As Boolean = True)

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

' OK
Private Sub 文字置換(ByVal 検索文字 As String, ByVal 置換文字 As String)

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
    
        For i = 1 To myBk.Sheets.count
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
