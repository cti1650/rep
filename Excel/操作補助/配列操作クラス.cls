VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "配列操作クラス"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'配列操作クラス Ver 1.6
Option Explicit
Option Base 1                       ' 配列の要素を1スタートに設定 List(1,1)

' プライベート変数定義

'オプション機能用　GetTickCountの定義
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private BufferTime As Long          ' オプション機能用 マクロの動作時間検証用

Private cArray() As Variant         ' ClassArray = クラスで使用する変更前の配列を格納する変数

Private cChArray() As Variant       ' ClassChangeArray = クラスで使用する変更後の配列を格納する変数

Private cMArray() As Variant        ' ClassMemoryArray = クラスで使用する配列を一時的に格納する変数

Private c検索一致行Array() As Variant   ' Class検索一致行Array = クラスで検索した検索に一致する行の一覧を配列として格納する変数

Private cBRange As Range            ' ClassBeforeRange = クラスで使用する呼出元セルのオブジェクトを格納する変数

Private cARange As Range            ' ClassAfterRange = クラスで使用する書込み先セルのオブジェクトを格納する変数

Private cAType As String            ' ClassArrayType = クラスで使用する配列の形式を格納する変数

Private cMFlag As Boolean           ' ClassMemoryFlag = クラスで使用するメモリ内のデータの有無を格納する変数

Private CA1L As Long                ' ChangeArray1dimensionalLeast = 変更後の配列の1次元の最小要素数を格納する変数

Private CA1U As Long                ' ChangeArray1dimensionalUtmost = 変更後の配列の1次元の最大要素数を格納する変数

Private CA2L As Long                ' ChangeArray2dimensionalLeast = 変更後の配列の2次元の最小要素数を格納する変数

Private CA2U As Long                ' ChangeArray2dimensionalUtmost = 変更後の配列の2次元の最大要素数を格納する変数

Private c比較方法 As skey           ' Class比較方法 = cChArray配列の検索用比較方法を格納する変数

Private c検索列 As Long             ' Class検索列 = cChArray配列の検索用検索列を格納する変数

' 定数定義

Const デバックモード As Variant = False
Const メッセージモード As Variant = True
Const 配列要素初期値 As Integer = 1
Const 検索列初期値 As Integer = 0
Const メモリ格納情報初期値 As Variant = False
Const データ未格納時メッセージ = "配列またはセルオブジェクトがインプットされていません。"
Const 無効データ格納時メッセージ = "正しい値が入力されていません。"
Const Variant型配列 As String = "Variant()"
Const String型配列 As String = "String()"
Const Range型変数 As String = "Range"
Const 未格納オブジェクト変数 As String = "Nothing"
Const 配列情報表示用1 As String = "【編集中の配列】 一次元 "
Const 配列情報表示用2 As String = " To "
Const 配列情報表示用3 As String = " 、二次元 "


' 列挙型変数定義

Enum ArrayRC
    元リスト = 0
    編集後リスト = 1
    検索一致行リスト = 2
End Enum

Enum skey
    一致 = 0
    一致しない = 1
    含む = 2
    含まない = 3
    より小さい = 4
    より小さいか等しい = 5
    より大きい = 6
    より大きいか等しい = 7
End Enum

' イベントプロシージャ

Private Sub Class_Initialize()

    Erase cArray
    Erase cChArray
    Erase cMArray
    Erase c検索一致行Array
    
    cAType = 未格納オブジェクト変数
    cMFlag = メモリ格納情報初期値
    c比較方法 = 一致
    c検索列 = 検索列初期値
    
    Set cBRange = Nothing
    Set cARange = Nothing
    
End Sub

Private Sub Class_Terminate()

    Erase cArray
    Erase cChArray
    Erase cMArray
    Erase c検索一致行Array
    
    cAType = ""
    cMFlag = メモリ格納情報初期値
    c比較方法 = 一致
    c検索列 = 検索列初期値
    
    Set cBRange = Nothing
    Set cARange = Nothing
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
End Sub

' プライベートプロシージャ

Private Sub 配列情報表示()

    If cAType <> Variant型配列 And cAType <> Range型変数 Then
    
        MsgBox データ未格納時メッセージ
        Exit Sub
        
    End If
    
    Dim vStr(8) As Variant
    
    配列情報更新
    
    vStr(1) = 配列情報表示用1
    vStr(2) = CA1L
    vStr(3) = 配列情報表示用2
    vStr(4) = CA1U
    vStr(5) = 配列情報表示用3
    vStr(6) = CA2L
    vStr(7) = 配列情報表示用2
    vStr(8) = CA2U

    Application.StatusBar = Join(vStr, vbNullString)

End Sub

Private Sub 配列情報更新()

    If cAType <> Variant型配列 And cAType <> Range型変数 Then
    
        MsgBox データ未格納時メッセージ
        Exit Sub
        
    End If
    
    CA1L = LBound(cChArray, 1)
    CA1U = UBound(cChArray, 1)
    CA2L = LBound(cChArray, 2)
    CA2U = UBound(cChArray, 2)

End Sub

Private Function 配列要素数(ByVal vArray As Variant) As Long

    Dim ArrayData() As Variant
    Dim TempData As Variant
    Dim i As Long
    
    ArrayData = vArray

    On Error Resume Next
    
    Do While Err.Number = 0
    
        i = i + 1
        TempData = UBound(ArrayData, i)
        
    Loop
    
    On Error GoTo 0
    
    配列要素数 = i - 1
    
End Function

Private Function 条件比較(ByVal 検索値1 As Variant, ByVal 検索値2 As Variant, Optional ByVal 比較方法 As skey = 0) As Boolean
    
    On Error Resume Next
    
    Err.Clear
    
    Select Case 比較方法
    Case 0
    
        If IsDate(検索値1) And IsDate(検索値2) Then
        
            条件比較 = CDate(検索値1) = CDate(検索値2)
            
        Else
        
            条件比較 = (StrConv(検索値1, vbLowerCase + vbNarrow) = StrConv(検索値2, vbLowerCase + vbNarrow))
            
        End If
        
    Case 1
    
        If IsDate(検索値1) And IsDate(検索値2) Then
        
            条件比較 = Not CDate(検索値1) = CDate(検索値2)
            
        Else
        
            条件比較 = Not (StrConv(検索値1, vbLowerCase + vbNarrow) = StrConv(検索値2, vbLowerCase + vbNarrow))
            
        End If
        
    Case 2
    
        If IsDate(検索値1) And IsDate(検索値2) Then
        
            条件比較 = Year(CDate(検索値1)) = Year(CDate(検索値2)) And Month(CDate(検索値1)) = Month(CDate(検索値2))
            
        Else
        
            条件比較 = StrConv(検索値1, vbLowerCase + vbNarrow) Like StrConv(検索値2, vbLowerCase + vbNarrow)
            
        End If
        
    Case 3
    
        If IsDate(検索値1) And IsDate(検索値2) Then
        
            条件比較 = Not (Year(CDate(検索値1)) = Year(CDate(検索値2))) And Not (Month(CDate(検索値1)) <> Month(CDate(検索値2)))
            
        Else
        
            条件比較 = Not StrConv(検索値1, vbLowerCase + vbNarrow) Like StrConv(検索値2, vbLowerCase + vbNarrow)
            
        End If
        
    Case 4
    
        If IsDate(検索値1) And IsDate(検索値2) Then
        
            条件比較 = CDate(検索値1) < CDate(検索値2)
            
        ElseIf IsNumeric(検索値1) And IsNumeric(検索値2) Then
        
            条件比較 = 検索値1 < 検索値2
            
        Else
        
            条件比較 = False
            
        End If
        
    Case 5
    
        If IsDate(検索値1) And IsDate(検索値2) Then
        
            条件比較 = CDate(検索値1) <= CDate(検索値2)
            
        ElseIf IsNumeric(検索値1) And IsNumeric(検索値2) Then
        
            条件比較 = 検索値1 <= 検索値2
            
        Else
        
            条件比較 = False
            
        End If
        
    Case 6
    
        If IsDate(検索値1) And IsDate(検索値2) Then
        
            条件比較 = CDate(検索値1) > CDate(検索値2)
            
        ElseIf IsNumeric(検索値1) And IsNumeric(検索値2) Then
        
            条件比較 = 検索値1 > 検索値2
            
        Else
        
            条件比較 = False
            
        End If
        
    Case 7
    
        If IsDate(検索値1) And IsDate(検索値2) Then
        
            条件比較 = CDate(検索値1) >= CDate(検索値2)
            
        ElseIf IsNumeric(検索値1) And IsNumeric(検索値2) Then
        
            条件比較 = 検索値1 >= 検索値2
            
        Else
        
            条件比較 = False
            
        End If
        
    End Select
    
    If Err.Number <> 0 Then
    
        条件比較 = False
        
    End If
    
End Function

' パブリックプロシージャ

' Subプロシージャ

Public Sub リセット()

    If Not (cAType = Variant型配列 Or cAType = Range型変数) Then
    
        MsgBox データ未格納時メッセージ
        Exit Sub
        
    End If
    
    cChArray = cArray
    
    配列情報更新
    
End Sub

Public Sub メモリ書込()

    If Not (cAType = Variant型配列 Or cAType = Range型変数) Then
    
        MsgBox データ未格納時メッセージ
        Exit Sub
        
    End If

    cMArray = cChArray
    
    cMFlag = True
    
End Sub

Public Sub メモリ読込()

    If Not (cAType = Variant型配列 Or cAType = Range型変数) Then
    
        MsgBox データ未格納時メッセージ
        Exit Sub
        
    ElseIf cMFlag = False Then
    
        MsgBox "メモリに配列が書き込まれていません。"
        Exit Sub
        
    End If

    cChArray = cMArray
    
    配列情報更新
    
End Sub

Public Sub クリア()

    Erase cArray
    Erase cChArray
    Erase cMArray
    Erase c検索一致行Array
    
    cAType = 未格納オブジェクト変数
    cMFlag = メモリ格納情報初期値
    c比較方法 = 一致
    c検索列 = 検索列初期値
    
    Set cBRange = Nothing
    Set cARange = Nothing
    
End Sub

Public Sub インプット(Optional ByRef 配列 As Variant)

    Dim vArray As Variant
    Dim myArray() As Variant
    Dim 配列補正行, 配列補正列 As Long
    Dim i, j As Long
    Dim a As Variant
    Dim StrArray As Variant

    'Debug.Print "引数形式：" & TypeName(配列)

    Select Case True
    Case TypeName(配列) = Variant型配列
    
        Set cBRange = Nothing
        
        'Debug.Print "加工前配列要素数： " & 配列要素数(配列)
        
        If 配列要素数(配列) = 2 Then
        
            If UBound(配列, 2) = 配列要素初期値 Then
            
                ReDim myArray(UBound(配列, 1) - LBound(配列, 1) + 1, 1)
            
                i = 0
            
                For Each a In 配列
            
                    i = i + 1
                    myArray(i, 1) = a
                
                Next a
            
            Else
            
                vArray = WorksheetFunction.Transpose(WorksheetFunction.Transpose(配列))
            
            End If
            
        Else
        
            If IsArray(配列(LBound(配列, 1))) Then
            
                MsgBox "異常な値が入力されました"
                
                'Debug.Print "【Error】 有効な配列が生成できませんでした"
                
                Exit Sub
                
            End If
            
            ReDim myArray(UBound(配列, 1) - LBound(配列, 1) + 1, 1)
            
            i = 0
            
            For Each a In 配列
            
                i = i + 1
                myArray(i, 1) = a
                
            Next a
            
            vArray = myArray
        
        End If
        
        'Debug.Print "【一次元】 最少要素： " & LBound(vArray, 1) & "  最大要素： " & UBound(vArray, 1) & _
                    "【二次元】 最少要素： " & LBound(vArray, 2) & "  最大要素： " & UBound(vArray, 2)
        
        cArray = vArray
        cChArray = cArray
        cAType = Variant型配列
        
        配列情報更新
        
    Case TypeName(配列) = String型配列
    
        StrArray = 配列
    
        Set cBRange = Nothing
        
        'Debug.Print "加工前配列要素数： " & 配列要素数(配列)
            
        ReDim myArray(UBound(StrArray, 1) - LBound(StrArray, 1) + 配列要素初期値, 1)
            
        i = 0
            
        For Each a In StrArray
            
            i = i + 1
            myArray(i, 1) = a
                
        Next
            
        vArray = myArray
        
        
        'Debug.Print "【一次元】 最少要素： " & LBound(vArray, 1) & "  最大要素： " & UBound(vArray, 1) & _
                    "【二次元】 最少要素： " & LBound(vArray, 2) & "  最大要素： " & UBound(vArray, 2)
        
        cArray = vArray
        cChArray = cArray
        cAType = Variant型配列
        
        配列情報更新
        
    Case TypeName(配列) = Range型変数
    
        Set cBRange = 配列
        
        'Debug.Print "【一次元】 最少要素： " & LBound(cBRange.Value, 1) & "  最大要素： " & UBound(cBRange.Value, 1) & _
                    "【二次元】 最少要素： " & LBound(cBRange.Value, 2) & "  最大要素： " & UBound(cBRange.Value, 2)
        
        cArray = cBRange.Value
        cChArray = cArray
        cAType = Range型変数
        
        配列情報更新
        
    Case Else
    
        MsgBox "異常な値が入力されました。"
        
        'Debug.Print "【Error】 有効な配列が生成できませんでした"
        
    End Select

End Sub

Public Sub アウトプット(Optional ByRef vRange As Object)

    If Not (cAType = Variant型配列 Or cAType = Range型変数 Or cAType = String型配列) Then
    
        MsgBox データ未格納時メッセージ
        Exit Sub

    End If

    Dim セル行数, セル列数 As Long

    If vRange Is Nothing Then

        If MsgBox("参照元のセルを上書きしますが宜しいですか？") = vbOK Then

            Set cARange = cBRange
            
        Else
        
            Exit Sub
            
        End If

    Else

        Set cARange = vRange
        
    End If
    
    'Debug.Print "【一次元】 最少要素： " & LBound(cChArray, 1) & "  最大要素： " & UBound(cChArray, 1) & _
                    "【二次元】 最少要素： " & LBound(cChArray, 2) & "  最大要素： " & UBound(cChArray, 2)
    
    With cARange
        
        '.NumberFormat = "General" ' コメントアウトを解除で書込み先セルの書式を標準に設定
        
        'Debug.Print "加工後配列要素数： " & 配列要素数(vArray)
        
        セル行数 = .Cells.Rows.Count
        セル列数 = .Cells.Columns.Count
            
        If CA1L = 配列要素初期値 And CA2L = 配列要素初期値 Then
            
            Select Case True
                Case CA1U >= セル行数 And CA2U >= セル列数
                    
                    .Cells.Value = cChArray
               
                Case CA1U >= セル行数 And CA2U < セル列数
                    
                    .Cells(1, CA2U + 1).Resize(セル行数, セル列数 - CA2U).Value = ""
                    
                    .Cells.Resize(セル行数, CA2U).Value = cChArray
                
                Case CA2U >= セル列数 And CA1U < セル行数
                
                    .Cells(CA1U + 1, 1).Resize(セル行数 - CA1U, セル列数).Value = ""
                    
                    .Cells.Resize(CA1U, セル列数).Value = cChArray
            
                Case Else
                
                    .Cells(1, CA2U + 1).Resize(セル行数, セル列数 - CA2U).Value = ""
                
                    .Cells(CA1U + 1, 1).Resize(セル行数 - CA1U, セル列数).Value = ""
                    
                    .Cells.Resize(CA1U, CA2U).Value = cChArray
                
            End Select

        Else
                
            MsgBox "正常に値が設定されていません"
                    
        End If
            
    End With

End Sub

Public Function CSVインプット(Optional ByVal URL As String = "", Optional ByVal 区切り文字 As String = ",", Optional ByVal 改行文字 As String = "") As Boolean

    On Error Resume Next

    Dim vArray As Variant
    Dim FileNamePath As Variant

    If Dir(URL) = "" Then
    
        FileNamePath = Application.GetOpenFilename("テキストファイル (*.csv),*.csv", , "取り込むCSVファイルを選択して下さい")
    
        If FileNamePath = False Or 区切り文字 = "" Then
        
            CSVインプット = False
            
            Exit Function
            
        End If
    
        URL = FileNamePath
    
    End If
    
        
        
    Dim fn As String, delim As String, temp As String, myColumns
    Dim x, y, a() As Variant, i As Long, ii As Long, pt As String, op_file As String
    Dim buf As Variant
        
        
    pt = URL
    op_file = Dir("*.csv")
    
    If op_file = "" Then
    
        CSVインプット = False
            
        Exit Function
    
    End If
 
    fn = URL
    delim = 区切り文字
    temp = CreateObject("Scripting.FileSystemObject").OpenTextFile(fn).ReadAll
    
    If 改行文字 = "" Then
    
        Select Case True
        Case temp Like "*" & vbCrLf & "*"
            改行文字 = vbCrLf
        Case temp Like "*" & vbCr & "*"
            改行文字 = vbCr
        Case temp Like "*" & vbLf & "*"
            改行文字 = vbLf
        Case Else
            改行文字 = vbCrLf
        End Select
        
    End If
    
    x = Split(temp, 改行文字)
    y = Split(x(0), delim)
        
    ReDim a(1 To UBound(x) + 1, 1 To UBound(y) + 1)
    For i = LBound(x) To UBound(x)
        y = Split(x(i), delim)
        For ii = LBound(a, 2) To UBound(a, 2)
            If UBound(y) >= ii Then
                buf = y(ii - 1)
                If buf Like """*" Then buf = Mid$(buf, 2, Len(buf) - 1)
                If buf Like "*""" Then buf = Left$(buf, Len(buf) - 1)
                a(i + 1, ii) = buf
            End If
        Next
    Next
        
    Set cBRange = Nothing
        
    vArray = a
        
    cArray = vArray
    cChArray = cArray
    cAType = Variant型配列
        
    配列情報更新
        
    If Err.Number <> 0 Then
        
        CSVインプット = True
            
    Else
        
        CSVインプット = False
        
    End If
        
    On Error GoTo 0

End Function

' Functionプロシージャ

Public Function リスト(Optional ByVal リスト名 As ArrayRC = 0) As Variant

    If Not (cAType = Variant型配列 Or cAType = Range型変数) Then
    
        Dim List(1, 1) As Variant
    
        MsgBox データ未格納時メッセージ
        
        List(1, 1) = ""
        
        リスト = List
        
        Exit Function
        
    End If

    Select Case True
    Case リスト名 = 元リスト
    
        リスト = cArray
        
        'Debug.Print "初期配列要素数： " & 配列要素数(cArray)
        
    Case リスト名 = 編集後リスト
    
        リスト = cChArray
        
        'Debug.Print "加工後配列要素数： " & 配列要素数(cChArray)
        
    Case リスト名 = 検索一致行リスト
    
        リスト = c検索一致行Array
        
        'Debug.Print "加工後配列要素数： " & 配列要素数(c検索一致行Array)
        
    End Select
    
End Function

Public Function 行数(Optional ByVal リスト名 As ArrayRC = 0) As Long

    Select Case True
    Case リスト名 = 元リスト
    
        行数 = UBound(cArray, 1)
        
    Case リスト名 = 編集後リスト
    
        行数 = CA1U
        
    Case リスト名 = 検索一致行リスト
    
        行数 = UBound(c検索一致行Array, 1)
        
    End Select
    
End Function

Public Function 列数(Optional ByVal リスト名 As ArrayRC = 0) As Long

    If 配列要素数(cArray) = 2 Then
    
        Select Case True
        Case リスト名 = 元リスト
        
            列数 = UBound(cArray, 2)
            
        Case リスト名 = 編集後リスト
        
            列数 = CA2U
            
        Case リスト名 = 検索一致行リスト
    
            列数 = 1
            
        End Select
        
    Else
    
        列数 = 1
        MsgBox "1次元配列です"
        
    End If
    
End Function

Public Sub 検索設定(Optional ByVal 比較方法 As skey, Optional ByVal 検索列 As Long)

    If IsEmpty(比較方法) Then Else c比較方法 = 比較方法
    
    If IsEmpty(検索列) Then Else c検索列 = 検索列

End Sub

Public Function 検索(ParamArray 検索値()) As Long

    If Not (cAType = Variant型配列 Or cAType = Range型変数) Then
    
        MsgBox データ未格納時メッセージ
        
        検索 = 0
        
        Exit Function
        
    End If

    Dim i, j, k As Long
    Dim a As Variant
    Dim arg As Variant
    Dim Flag As Boolean
    Dim myFlag As Boolean
    Dim vArray() As Variant
    Dim myArray() As Variant
    
    i = 配列要素初期値
    j = 配列要素初期値
    k = 1
    Flag = False
    myFlag = False
    
    
    If IsArray(検索値) Then
    
        ReDim myArray(1)
    
        If c検索列 = 0 Then
        
            i = CA1L
            
            Do
                
                j = CA2L
                
                Flag = False
                myFlag = False
            
                Do
                    
                    For Each arg In 検索値
                        
                        If 条件比較(cChArray(i, j), arg, c比較方法) And Not c比較方法 = 含まない Then
                        
                            ReDim Preserve myArray(k)
                        
                            myArray(k) = i
                            k = k + 1
                            
                            Flag = True
                            
                            Exit Do
                            
                        ElseIf Not 条件比較(cChArray(i, j), arg, c比較方法) And c比較方法 = 含まない Then
                            
                            myFlag = True
                            
                        ElseIf c比較方法 = 含まない And j = CA2U And myFlag = False Then
                        
                            ReDim Preserve myArray(k)
                        
                            myArray(k) = i
                            k = k + 1
                            
                            Flag = True
                            
                            Exit Do
                            
                        End If
                        
                    Next arg
                
                    j = j + 1
                
                Loop While j <= CA2U And Flag = False
                
                i = i + 1
                
            Loop While i <= CA1U
            
        Else
        
            i = CA1L
            
            Do
                
                For Each arg In 検索値
                        
                    If 条件比較(cChArray(i, c検索列), arg, c比較方法) Then
                        
                        ReDim Preserve myArray(k)
                        
                        myArray(k) = i
                        k = k + 1
                        
                        Exit For
                            
                    End If
                        
                Next arg
            
                i = i + 1
                
            Loop While i <= CA1U
            
        End If
        
        If myArray(1) = "" And UBound(myArray, 1) = 配列要素初期値 Then
            
            c検索一致行Array = myArray
            
            ReDim vArray(1, 1)
            
            検索 = 0
            
        Else
            c検索一致行Array = myArray
        
            ReDim vArray(UBound(myArray), CA2U)
        
            k = 1
        
            For Each a In myArray
                
                j = CA2L
            
                Do
            
                    vArray(k, j) = cChArray(a, j)
                
                    j = j + 1
                
                Loop While j <= CA2U
            
                k = k + 1
                
            Next a
            
            検索 = UBound(myArray)
        
        End If
        
        cChArray = vArray
            
        配列情報更新
        
    Else
    
        MsgBox "異常な値が検索値に設定されました。"
        
    End If
    
End Function

Public Function 重複(ByVal 検索列 As Long, ByVal 検索値 As Variant)

    Dim i As Long
    Dim myC As New 配列操作クラス

    With myC
        .インプット cChArray
        .検索設定 一致, 検索列
        i = .検索(検索値)
        If i Then
            MsgBox "重複があります。"
        Else
            MsgBox "重複はありません。"
        End If
        重複 = i
    End With
    
    Set myC = Nothing

End Function

Public Function 合計() As Double

    If Not (cAType = Variant型配列 Or cAType = Range型変数) Then
    
        MsgBox データ未格納時メッセージ
        Exit Function
        
    End If
    
    合計 = WorksheetFunction.Sum(cChArray)

End Function

Public Function 平均() As Double

    If Not (cAType = Variant型配列 Or cAType = Range型変数) Then
    
        MsgBox データ未格納時メッセージ
        Exit Function
        
    End If
    
    平均 = WorksheetFunction.Average(cChArray)

End Function

Public Function カウント() As Long

    If Not (cAType = Variant型配列 Or cAType = Range型変数) Then
    
        MsgBox データ未格納時メッセージ
        Exit Function
        
    End If
    
    カウント = WorksheetFunction.CountA(cChArray)

End Function

Public Function 行抽出(ParamArray 抽出行()) As Variant

    If Not (cAType = Variant型配列 Or cAType = Range型変数) Then
    
        MsgBox データ未格納時メッセージ
        Exit Function
        
    End If
    
    Dim vArray As Variant
    Dim myArray As Variant
    Dim StrArray As Variant
    Dim a As Variant
    Dim 行 As Variant
    Dim i, j, k, l, 開始行, 終了行 As Long
    
    ReDim vArray(CA2U, 1)
    
    k = l = 0
    
    For Each 行 In 抽出行
    
        If IsNumeric(行) And 行 <= CA1U And Not 行 Like "*,*" And Not 行 Like "*:*" Then
        
            l = l + 1
            ReDim Preserve vArray(CA2U, l)
            
            j = CA2L
            
            Do
            
                vArray(j, l) = cChArray(行, j)
                
                j = j + 1
                
            Loop While j <= CA2U
            
        ElseIf 行 Like "*,*" Then
        
            StrArray = Split(行, ",")
            
            For Each a In StrArray
            
                l = l + 1
                ReDim Preserve vArray(CA2U, l)
            
                j = CA2L
            
                Do
            
                    vArray(j, l) = cChArray(Val(a), j)
                
                    j = j + 1
                
                Loop While j <= CA2U
                
            Next a
            
        ElseIf 行 Like "*:*" And Not 行 Like "*:*:*" Then
        
            StrArray = Split(行, ":")
        
            開始行 = Val(StrArray(0))
            終了行 = Val(StrArray(1))
            
            If 開始行 >= CA1L And 終了行 <= CA1U Then
                
                i = 開始行
                j = CA2L
        
                Do
        
                    l = l + 1
                    ReDim Preserve vArray(CA2U, l)
        
                    Do
            
                        vArray(j, l) = cChArray(i, j)
                
                        j = j + 1
                
                    Loop While j <= CA2U
            
                    j = LBound(vArray, 1)
                    i = i + 1
            
                Loop While i <= 終了行
                
            End If
            
        End If
    
    Next
    
    If UBound(vArray, 2) = 配列要素初期値 Then
    
        ReDim myArray(1, UBound(vArray, 1))
        
        j = LBound(vArray, 1)
        
        Do
            
            myArray(1, j) = vArray(j, 配列要素初期値)
                
            j = j + 1
                
        Loop While j <= UBound(vArray, 1)
        
        cChArray = myArray
        
    Else
        
        ReDim myArray(UBound(vArray, 2), UBound(vArray, 1))
        
        i = LBound(vArray, 2)
        j = LBound(vArray, 1)
        
        Do
        
            Do
            
                myArray(i, j) = vArray(j, i)
                
                j = j + 1
                
            Loop While j <= UBound(vArray, 1)
            
            j = LBound(vArray, 1)
            i = i + 1
            
        Loop While i <= UBound(vArray, 2)
        
        cChArray = myArray
        
    End If
            
    行抽出 = cChArray
            
    配列情報更新

End Function

Public Function 列抽出(ParamArray 抽出列()) As Variant

    If Not (cAType = Variant型配列 Or cAType = Range型変数) Then
    
        MsgBox データ未格納時メッセージ
        Exit Function
        
    End If
    
    Dim vArray As Variant
    Dim myArray As Variant
    Dim StrArray As Variant
    Dim a As Variant
    Dim 列 As Variant
    Dim i, j, k, l, 開始列, 終了列 As Long
    
    ReDim vArray(CA1U, 1)
    
    k = l = 0
    
    For Each 列 In 抽出列
    
        If IsNumeric(列) And 列 <= CA2U And Not 列 Like "*,*" And Not 列 Like "*:*" Then
        
            l = l + 1
            ReDim Preserve vArray(CA1U, l)
            
            j = CA1L
            
            Do
            
                vArray(j, l) = cChArray(j, 列)
                
                j = j + 1
                
            Loop While j <= CA1U
            
        ElseIf 列 Like "*,*" Then
        
            StrArray = Split(列, ",")
            
            For Each a In StrArray
            
                l = l + 1
                ReDim Preserve vArray(CA1U, l)
            
                j = CA1L
                
                Do
            
                    vArray(j, l) = cChArray(j, Val(a))
                
                    j = j + 1
                
                Loop While j <= CA1U
                
            Next a
            
        ElseIf 列 Like "*:*" And Not 列 Like "*:*:*" Then
        
            StrArray = Split(列, ":")
        
            開始列 = Val(StrArray(0))
            終了列 = Val(StrArray(1))
            
            If 開始列 >= CA2L And 終了列 <= CA2U Then
                
                i = 開始列
        
                Do
        
                    l = l + 1
                    ReDim Preserve vArray(CA1U, l)
                    
                    j = CA1L
                    
                    Do
            
                        vArray(j, l) = cChArray(j, i)
                
                        j = j + 1
                
                    Loop While j <= CA1U
            
                    i = i + 1
            
                Loop While i <= 終了列
                
            End If
            
        End If
    
    Next
    
    cChArray = vArray
            
    列抽出 = cChArray
            
    配列情報更新

End Function

Public Sub 列昇順ソート(Optional ByVal 対象行 As Long = 1)

    If Not (TypeName(cARange) = Range型変数 And 対象行 <= CA1U) Then
    
        MsgBox 無効データ格納時メッセージ & vbCrLf & "または、アウトプット後実行して下さい。"
        Exit Sub
        
    End If
    
    Dim vRange As Range
    
    If cARange Is Nothing Then

        MsgBox "アウトプット後実行して下さい。"

        Exit Sub

    Else

        Set vRange = cARange
        
    End If
    
    vRange.Sort vRange.Cells(対象行, 1), xlAscending, Header:=xlNo, MatchCase:=False, Orientation:=xlSortRows
    
End Sub

Public Sub 列降順ソート(Optional ByVal 対象行 As Long = 1)

    If Not (TypeName(cARange) = Range型変数 And 対象行 <= CA1U) Then
    
        MsgBox 無効データ格納時メッセージ & vbCrLf & "または、アウトプット後実行して下さい。"
        Exit Sub
        
    End If
    
    Dim vRange As Range
    
    If cARange Is Nothing Then

        MsgBox "アウトプット後実行して下さい。"

        Exit Sub
        
    Else

        Set vRange = cARange
        
    End If
    
    vRange.Sort vRange.Cells(対象行, 1), xlDescending, Header:=xlNo, MatchCase:=False, Orientation:=xlSortRows
    
End Sub

Public Sub 行昇順ソート(Optional ByVal 対象列 As Long = 1)

    If Not (TypeName(cARange) = Range型変数 And 対象列 <= CA2U) Then
    
        MsgBox 無効データ格納時メッセージ & vbCrLf & "または、アウトプット後実行して下さい。"
        Exit Sub
        
    End If
    
    Dim vRange As Range
    
    If cARange Is Nothing Then

        MsgBox "アウトプット後実行して下さい。"

        Exit Sub
        
    Else

        Set vRange = cARange
        
    End If
    
    vRange.Sort vRange.Cells(1, 対象列), xlAscending, Header:=xlNo, MatchCase:=False, Orientation:=xlSortColumns
    
End Sub

Public Sub 行降順ソート(Optional ByVal 対象列 As Long = 1)

    If Not (TypeName(cARange) = Range型変数 And 対象列 <= CA2U) Then
    
        MsgBox 無効データ格納時メッセージ & vbCrLf & "または、アウトプット後実行して下さい。"
        Exit Sub
        
    End If
    
    Dim vRange As Range
    
    If cARange Is Nothing Then

        MsgBox "アウトプット後実行して下さい。"

        Exit Sub

    Else

        Set vRange = cARange
        
    End If
    
    vRange.Sort vRange.Cells(1, 対象列), xlDescending, Header:=xlNo, MatchCase:=False, Orientation:=xlSortColumns
    
End Sub

Public Sub 要素反転()

    If Not (cAType = Variant型配列 Or cAType = Range型変数) Then
    
        MsgBox データ未格納時メッセージ
        Exit Sub
        
    End If
    
    'cChArray = WorksheetFunction.Transpose(cChArray)
    
    Dim myArray As Variant
    
    Dim i, j As Long
    
    If UBound(cChArray, 2) = 1 Then
    
        ReDim myArray(1, UBound(cChArray, 1))
        
        j = LBound(cChArray, 1)
        
        Do
            
            myArray(1, j) = cChArray(j, 配列要素初期値)
                
            j = j + 1
                
        Loop While j <= UBound(cChArray, 1)
        
        cChArray = myArray
        
        配列情報更新
        
    Else
    
        'cChArray = WorksheetFunction.Transpose(vArray)
        
        ReDim myArray(UBound(cChArray, 2), UBound(cChArray, 1))
        
        i = LBound(cChArray, 2)
        j = LBound(cChArray, 1)
        
        Do
        
            Do
            
                myArray(i, j) = cChArray(j, i)
                
                j = j + 1
                
            Loop While j <= UBound(cChArray, 1)
            
            j = LBound(cChArray, 1)
            i = i + 1
            
        Loop While i <= UBound(cChArray, 2)
        
        cChArray = myArray
        
        配列情報更新
        
    End If

End Sub


'--------------------------------------------------------------------------------

' オプション機能

'--------------------------------------------------------------------------------

'実行速度検証用プロシージャ

Public Sub 速度検証スタート()
    
    BufferTime = GetTickCount()
    
End Sub

Public Function 速度検証ストップ(Optional ByVal ポップアップ As Boolean = True) As String
    
    Dim 結果 As String
    
    結果 = GetTickCount() - BufferTime & " ms"
    
    If ポップアップ Then MsgBox 結果
    
    速度検証ストップ = 結果
    
End Function

'--------------------------------------------------------------------------------

'速度改善用プロシージャ

Public Sub 画面更新ON()

    Application.ScreenUpdating = True

End Sub

Public Sub 画面更新OFF()

    Application.ScreenUpdating = False

End Sub

