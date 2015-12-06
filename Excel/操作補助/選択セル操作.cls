VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "選択セル操作"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private セル As Range

Private Sub Class_Initialize()

    セル再定義

End Sub

Private Sub Class_Terminate()

    Set セル = Nothing

End Sub

Private Sub セル再定義()

    Set セル = Selection

End Sub

Public Function 同一行セル(Optional ByVal 取得列番号 As Long = 0) As Range

    セル再定義

    If TypeName(セル) <> "Range" Then
        Set 同一行セル = Nothing
        Exit Function
    End If

    If セル.Count = 1 And 取得列番号 >= 1 And 取得列番号 <= Columns.Count Then
    
        Set 同一行セル = Cells(セル.Row, 取得列番号)
        
    ElseIf セル.Count = 1 And 取得列番号 = 0 Then
    
        Set 同一行セル = セル
        
    Else
    
        Set 同一行セル = Nothing
    
    End If

End Function

Public Function 同一列セル(Optional ByVal 取得行番号 As Long = 0) As Range

    セル再定義

    If TypeName(セル) <> "Range" Then
        Set 同一行セル = Nothing
        Exit Function
    End If

    If セル.Count = 1 And 取得行番号 >= 1 And 取得行番号 <= Rows.Count Then
    
        Set 同一列セル = Cells(取得行番号, セル.Column)
        
    ElseIf セル.Count = 1 And 取得列番号 = 0 Then
    
        Set 同一列セル = セル
        
    Else
    
        Set 同一列セル = Nothing
    
    End If

End Function

Public Sub 上移動(Optional ByVal 非表示行スキップ As Boolean = True)

    セル再定義

    If TypeName(セル) <> "Range" Then Exit Sub
    
    Dim i As Long
    Dim buf As Long
    
    i = 0
    buf = -1

    If セル.Count = 1 And セル.Row > 1 Then
    
        If 非表示行スキップ = True Then
    
            Do
    
                i = i + buf
    
            Loop While セル.Offset(i, 0).EntireRow.Hidden = True
            
        Else
        
            i = i + buf
        
        End If
        
        セル.Offset(i, 0).Select
        
    Else
    
    End If

End Sub

Public Sub 下移動(Optional ByVal 非表示行スキップ As Boolean = True)

    セル再定義

    If TypeName(セル) <> "Range" Then Exit Sub
    
    Dim i As Long
    Dim buf As Long
    
    i = 0
    buf = 1

    If セル.Count = 1 And セル.Row < Rows.Count Then
    
        If 非表示行スキップ = True Then
    
            Do
    
                i = i + buf
    
            Loop While セル.Offset(i, 0).EntireRow.Hidden = True
            
        Else
        
            i = i + buf
        
        End If
        
        セル.Offset(i, 0).Select
        
    Else
    
    End If

End Sub

Public Sub 左移動(Optional ByVal 非表示列スキップ As Boolean = True)

    セル再定義

    If TypeName(セル) <> "Range" Then Exit Sub
    
    Dim i As Long
    Dim buf As Long
    
    i = 0
    buf = -1

    If セル.Count = 1 And セル.Column > 1 Then
    
        If 非表示列スキップ = True Then
    
            Do
    
                i = i + buf
    
            Loop While セル.Offset(0, i).EntireColumn.Hidden = True
            
        Else
        
            i = i + buf
        
        End If
        
        セル.Offset(0, i).Select
        
    Else
    
    End If

End Sub

Public Sub 右移動(Optional ByVal 非表示列スキップ As Boolean = True)

    セル再定義

    If TypeName(セル) <> "Range" Then Exit Sub
    
    Dim i As Long
    Dim buf As Long
    
    i = 0
    buf = 1

    If セル.Count = 1 And セル.Column < Columns.Count Then
    
        If 非表示列スキップ = True Then
    
            Do
    
                i = i + buf
    
            Loop While セル.Offset(0, i).EntireColumn.Hidden = True
            
        Else
        
            i = i + buf
        
        End If
        
        セル.Offset(0, i).Select
        
    Else
    
    End If

End Sub
