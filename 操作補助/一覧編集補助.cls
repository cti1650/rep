VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "一覧編集補助"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private cls開始行 As Long
Private cls終了行 As Long
Private cls対象シート名 As Long
Private WithEvents Sh As Worksheet
Attribute Sh.VB_VarHelpID = -1

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    Set Sh = Nothing
End Sub



Public Function 初期設定(ByVal 開始行 As Long, ByVal 終了行 As Long, Optional ByVal 対象シート名 As String = "") As Boolean

    On Error Resume Next
    
    If 開始行 > 1 And 終了行 > 開始行 Then
    
        cls開始行 = 開始行
        cls終了行 = 終了行
        cls対象シート名 = 対象シート名
    
        初期設定 = True
    
    Else
    
        cls開始行 = 0
        cls終了行 = 0
        cls対象シート名 = ""
    
        初期設定 = False
    
    End If
    
    On Error GoTo 0

End Function

Public Sub 検索(ByVal 検索値 As Variant)

    If (cls開始行 = 0 And cls終了行 = 0) Or ActiveSheet.Name Like cls対象シート名 Then Exit Sub
    
    Dim Rn1 As Range
    Dim Rn2 As Range
    Dim i As Long
    
    Set Sh = ActiveSheet
    
    For i = cls開始行 To cls終了行
    
        If Application.WorksheetFunction.CountIf(Sh.Rows(i), 検索値) > 0 Then
        
            If Rn1 Is Nothing Then
            
                Set Rn1 = Sh.Rows(i)
            
            Else
            
                Set Rn1 = Application.Union(Rn1, Sh.Rows(i))
            
            End If
        
        Else
        
            If Rn2 Is Nothing Then
            
                Set Rn2 = Sh.Rows(i)
            
            Else
            
                Set Rn2 = Application.Union(Rn2, Sh.Rows(i))
            
            End If
        
        End If
    
    Next
    
    If Not Rn1 Is Nothing Then
            
        Rn1.EntireRow.Hidden = False
            
    End If
    
    If Not Rn2 Is Nothing Then
            
        Rn2.EntireRow.Hidden = True
            
    End If
    
    Set Rn1 = Nothing
    Set Rn2 = Nothing

End Sub


Public Sub 連続検索(ByVal 検索値 As Variant)

    If (cls開始行 = 0 And cls終了行 = 0) Or ActiveSheet.Name Like cls対象シート名 Then Exit Sub
    
    Dim Rn2 As Range
    Dim i As Long
    
    Set Sh = ActiveSheet
    
    For i = cls開始行 To cls終了行
    
        If Application.WorksheetFunction.CountIf(Sh.Rows(i), 検索値) = 0 Then
        
            If Rn2 Is Nothing Then
            
                Set Rn2 = Sh.Rows(i)
            
            Else
            
                Set Rn2 = Application.Union(Rn2, Sh.Rows(i))
            
            End If
        
        End If
    
    Next
    
    If Not Rn2 Is Nothing Then
            
        Rn2.EntireRow.Hidden = True
            
    End If
    
    Set Rn2 = Nothing
    
End Sub

Public Sub リセット()

    If (cls開始行 = 0 And cls終了行 = 0) Or ActiveSheet.Name Like cls対象シート名 Then Exit Sub
    
    Set Sh = ActiveSheet
    
    Sh.Range(Sh.Rows(cls開始行), Sh.Rows(cls終了行)).Hidden = False

End Sub

Private Sub Sh_SelectionChange(ByVal Target As Range)

    If TypeName(Target) = "Range" Then
    
        If Target.Columns.Count < Sh.Columns.Count And Target.Columns.Count > 1 Then
        
            Target.SpecialCells(xlCellTypeVisible).Select
        
        End If
    
    End If

End Sub
