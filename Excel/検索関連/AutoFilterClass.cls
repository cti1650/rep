VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoFilterClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const TestMode As Boolean = False
Private Const 基準行 As String = "1:1"
Private 条件文字列 As String
Private 条件配列 As Variant
Private 基準セル範囲 As Range
Private bCal As Variant

Private Sub Class_Initialize()

    With Application
        bCal = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    Set 基準セル範囲 = ActiveCell
    条件文字列 = ""
    条件配列 = Array("", "")
    
End Sub

Private Sub Class_Terminate()

    条件文字列 = ""
    条件配列 = Array("", "")
    Set 基準セル範囲 = Nothing
    With Application
        .Calculation = bCal
        .ScreenUpdating = True
        .EnableEvents = True
    End With
    
End Sub

Private Function 抽出可否(Optional ByVal 空白確認値 As Variant = "-") As Boolean

    If 基準セル範囲 Is Nothing Or 空白確認値 = "" Then
        抽出可否 = False
        Exit Function
    End If
    
    抽出可否 = True
    Exit Function

End Function

Public Sub 抽出(Optional ByVal フィルターリセット As Boolean = True)

    If フィルターリセット = True Then Me.解除
    
    If 抽出可否(条件文字列) Then
    
        Dim i As Long
        Dim 検索項目 As String
        Dim 検索条件 As String
        Dim 検索条件配列 As Variant
        
        #If TestMode Then
            On Error Resume Next
        #End If
        
        Err.Clear
        
        With 基準セル範囲.Areas(1)
        
            For i = LBound(条件配列, 2) To UBound(条件配列, 2)
                検索項目 = 条件配列(0, i)
                検索条件 = 条件配列(1, i)
                Select Case True
                Case 検索条件 Like "*&&*"
                    検索条件配列 = Split(検索条件, "&&")
                    'DoEvents
                    .AutoFilter Field:=cls項目列(検索項目), Criteria1:=検索条件配列(0), Operator:=xlAnd, Criteria2:=検索条件配列(1)
                Case 検索条件 Like "*||*"
                    検索条件配列 = Split(検索条件, "||")
                    'DoEvents
                    .AutoFilter Field:=cls項目列(検索項目), Criteria1:=検索条件配列(0), Operator:=xlOr, Criteria2:=検索条件配列(1)
                Case 検索条件 Like "*(_)*"
                    検索条件配列 = Split(検索条件, "(_)")
                    'DoEvents
                    .AutoFilter Field:=cls項目列(検索項目), Criteria1:=検索条件配列, Operator:=xlFilterValues
                Case 検索項目 = ""
                Case Else
                    'DoEvents
                    .AutoFilter Field:=cls項目列(検索項目), Criteria1:=検索条件
                End Select
                
            Next
        
        End With
        
        If Err.Number <> 0 Then
        
            MsgBox "処理に失敗しました。 ErrorNo : " & Err.Number
        
        End If
        
        On Error GoTo 0
    
    End If
    

End Sub

Public Function 項目列(ByVal 検索項目名 As String) As Long

    Dim buf As Long

    buf = cls項目列(検索項目名)

    If buf Then
    
        項目列 = 基準セル範囲.Item(1, buf).Column
        
    Else
    
        項目列 = 0
        
    End If

End Function

Private Function cls結合セル値確認(ByRef セル範囲 As Range, Optional ByVal 行 As Long = 1, Optional ByVal 列 As Long = 1, Optional ByVal 代替値 = "") As Variant

    Dim 出力値 As Variant
    
    出力値 = 代替値

    If Not セル範囲 Is Nothing Then
    
        If 行 > 0 And 列 > 0 Then
        
            If 代替値 = "" Then
    
                With セル範囲.Cells(行, 列)
    
                    If .MergeCells = True Then
    
                        出力値 = .MergeArea(1).Value
                
                    End If
                
                End With
                
            End If
            
        End If
        
    End If
    
    cls結合セル値確認 = 出力値

End Function

Private Function cls項目列(ByVal 検索項目名 As String) As Long

    Dim Flag As Long
    
    Flag = 0
    
    If 抽出可否(検索項目名) Then
    
        Dim j As Long
        Dim i As Long
        Dim Rn As Object
        Dim 項目 As String
        Dim 検索項目 As String
        Dim 参照項目 As String
        Dim MRn As Range
        Dim 個別値 As Variant
        
        #If TestMode Then
            On Error Resume Next
        #End If
        
        Set MRn = 基準セル範囲
        
        Dim arRn As Variant
        
        arRn = MRn.Value
        
        For i = LBound(arRn, 2) To UBound(arRn, 2)
            項目 = ""
        
            For j = LBound(arRn, 1) To UBound(arRn, 1)
            
                個別値 = cls結合セル値確認(MRn, j, i, arRn(j, i))
                
                個別値 = replace$(replace$(個別値,vblf,""),vbcr,"")
                
                If Not StrConv(項目, vbNarrow + vbLowerCase) Like "*" & StrConv(個別値, vbNarrow + vbLowerCase) Then _
                        項目 = 項目 & 個別値
                        
                検索項目 = StrConv(検索項目名, vbNarrow + vbLowerCase)
                参照項目 = StrConv(項目, vbNarrow + vbLowerCase)
                
                If 検索項目 Like "*" & 参照項目 & "*" Then
                
                    Flag = i
                        
                    If 検索項目 = 参照項目 Then Exit For
                
                Else
                
                    Flag = 0
                    
                    Exit For
                    
                End If
            
            Next
            
            If 参照項目 = "" Then Exit For
            
            If Flag <> 0 Then Exit For
        
        Next
        
        On Error GoTo 0
    
    End If
    
    cls項目列 = Flag
    

End Function

Public Function 条件追加(ByVal 項目名 As String, Optional ByVal 条件式1 As String = "<>", Optional ByVal 条件式2 As String = "", Optional ByVal AND検索 As Boolean = False)

    Dim 条件式 As String
    
    条件式 = 条件式1
    
    If 条件式2 <> "" Then
    
        If AND検索 = True Then
        
            条件式 = 条件式 & "&&" & 条件式2
        
        Else
        
            条件式 = 条件式 & "||" & 条件式2
        
        End If
    
    End If

    If 項目名 Like "*" & vbCr & "*" or 項目名 Like "*" & vbLf & "*" Then 項目名 = replace$(replace$(項目名 ,vblf,""),vbcr,"")
    If 条件式 Like "*" & vbCr & "*" or 条件式 Like "*" & vbLf & "*" Then 条件式 = replace$(replace$(条件式 ,vblf,""),vbcr,"")
    
    If 条件文字列 <> "" Then 条件文字列 = 条件文字列 & vbCrLf
    
    条件文字列 = 条件文字列 & 項目名 & "^_-" & 条件式
    
    If 条件文字列 <> "" Then
    
        Dim j As Long
        Dim i As Long
        Dim c1 As Long
        Dim c2 As Long
        Dim a As Variant
        Dim b As Variant
        
        a = Split(条件文字列, vbCrLf)
        
        Dim Arr() As Variant
        
        ReDim Arr(2, UBound(a) + 1)
        
        For i = LBound(a) To UBound(a)
        
            b = Split(a(i), "^_-")
            
            If b(0) <> "" Then
            
                For j = LBound(b) To UBound(b)
                    'DoEvents
                    Arr(c2, c1) = b(j)
                
                    c2 = c2 + 1
                
                Next
                
            
                c2 = 0
                
            End If
            
            c1 = c1 + 1
            
        Next
            
        条件配列 = Arr
    
    End If
    
End Function

Public Function 条件値一括追加(ByVal 項目名 As String, ParamArray 条件値() As Variant)

    Dim 条件式 As String
    
    条件式 = Join$(条件値, "(_)")

    If 項目名 Like "*" & vbCr & "*" or 項目名 Like "*" & vbLf & "*" Then 項目名 = replace$(replace$(項目名 ,vblf,""),vbcr,"")
    If 条件式 Like "*" & vbCr & "*" or 条件式 Like "*" & vbLf & "*" Then 条件式 = replace$(replace$(条件式 ,vblf,""),vbcr,"")
    
    If 条件文字列 <> "" Then 条件文字列 = 条件文字列 & vbCrlf
    
    条件文字列 = 条件文字列 & 項目名 & "^_-" & 条件式
    
    If 条件文字列 <> "" Then
    
        Dim j As Long
        Dim i As Long
        Dim c1 As Long
        Dim c2 As Long
        Dim a As Variant
        Dim b As Variant
        
        a = Split(条件文字列, vbCrLf)
        
        Dim Arr() As Variant
        
        ReDim Arr(2, UBound(a) + 1)
        
        For i = LBound(a) To UBound(a)
        
            b = Split(a(i), "^_-")
            
            If b(0) <> "" Then
            
                For j = LBound(b) To UBound(b)
                    'DoEvents
                    Arr(c2, c1) = b(j)
                
                    c2 = c2 + 1
                
                Next
            
                c2 = 0
                
            End If
            
            c1 = c1 + 1
            
        Next
            
        条件配列 = Arr
    
    End If
    
End Function

Public Function 条件_今日(ByVal 項目名 As String, Optional ByVal 過去日 As Boolean = False)

    If 過去日 = True Then

        Me.条件追加 項目名, "<=" & CLng(CDate(Date))
        
    Else
    
        Me.条件追加 項目名, ">=" & CLng(CDate(Date)), "<=" & CLng(CDate(Date)), True
    
    End If
    
End Function

Public Function 条件_未来日(ByVal 項目名 As String)

    Me.条件追加 項目名, ">" & CLng(CDate(Date))
    
End Function

Public Function 条件_過去日(ByVal 項目名 As String)

    Me.条件追加 項目名, "<" & CLng(CDate(Date))
    
End Function

Public Function 条件_解除(ByVal 項目名 As String)

    Me.条件追加 項目名, "<>", "="
    
End Function

Public Function 条件_今週(ByVal 項目名 As String, Optional ByVal 過去日 As Boolean = False)

    If 過去日 = True Then

        Me.条件追加 項目名, "<=" & CLng(CDate(Date + 7))
        
    Else
    
        Me.条件追加 項目名, ">=" & CLng(CDate(Date)), "<=" & CLng(CDate(Date + 7)), True
    
    End If
    
End Function

Public Function 条件_今月(ByVal 項目名 As String, Optional ByVal 過去日 As Boolean = False)

    If 過去日 = True Then

        Me.条件追加 項目名, "<=" & CLng(CDate(DateSerial(Year(Date), Month(Date) + 1, 0)))
        
    Else
    
        Me.条件追加 項目名, ">=" & CLng(CDate(DateSerial(Year(Date), Month(Date), 1))), "<=" & CLng(CDate(DateSerial(Year(Date), Month(Date) + 1, 0))), True
    
    End If
    
End Function

Public Function 条件_選択行値(ByVal 項目名 As String)

    Me.条件追加 項目名, "=" & Cells(Selection.Row, cls項目列(項目名)).Value, "="
    
End Function

Public Function 条件クリア()

    条件文字列 = ""
    条件配列 = Array("", "")
    
End Function

Public Function 解除()

    If 抽出可否 Then

        With 基準セル範囲.Worksheet

            If .AutoFilterMode = True Then
            
                If CInt(Application.Version) > 11 Then
                    If .FilterMode = True Then

                        .ShowAllData
                
                    End If
                Else
                    Dim Flag As Boolean
                    Flag = False
                    For i = 1 To .AutoFilter.Filters.count
                        'DoEvents
                        If .AutoFilter.Filters(i).On Then
                            Flag = True
                            Exit For
                        End If
                    Next i
                    If Flag = True Then .ShowAllData
                End If
        
            Else
        
                基準セル範囲.Rows(基準セル範囲.Rows.count).AutoFilter
        
            End If
        
        End With
    
    End If
    

End Function

Public Function 項目行範囲(ByRef セル As Range)

    Set 基準セル範囲 = セル

End Function
