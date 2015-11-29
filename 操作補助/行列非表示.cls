VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "行列非表示"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private 行位置 As Boolean
Private cls開始位置 As Long
Private cls終了位置 As Long
Private cls範囲 As Long
Private cls基準Sh As Object


Private Sub Class_Initialize()

    On Error Resume Next
    
    Application.ScreenUpdating = False

    Me.初期化

End Sub

Private Sub Class_Terminate()

    Set cls基準Sh = Nothing
    行位置 = True
    
    Application.ScreenUpdating = True
    
    On Error GoTo 0

End Sub

Public Sub 初期化()

    行位置 = True
    cls開始位置 = 0
    cls終了位置 = 0
    cls範囲 = 0
    Set cls基準Sh = ThisWorkbook.ActiveSheet

End Sub

Public Sub 表示()

    If cls開始位置 <= 0 Then Exit Sub

    Dim i As Long

    If 行位置 = True Then
        Select Case True
        Case cls終了位置 > 0
            cls基準Sh.Range(cls基準Sh.Rows(cls開始位置), cls基準Sh.Rows(cls終了位置)).Hidden = False
        Case cls範囲 >= 0 And cls終了位置 = 0
            cls基準Sh.Rows(cls開始位置).Resize(cls範囲).Hidden = False
        End Select
    Else
        
        Select Case True
        Case cls終了位置 > 0
            cls基準Sh.Range(sh.Columns(cls開始位置), cls基準Sh.Columns(cls終了位置)).Hidden = False
        Case cls範囲 >= 0 And cls終了位置 = 0
            cls基準Sh.Columns(cls開始位置).Resize(cls範囲).Hidden = False
        End Select
        
    End If

End Sub

Public Sub 非表示(Optional ByVal 入力範囲表示 As Boolean = True, Optional ByVal 入力確認位置 As Long = 1)

    If cls開始位置 <= 0 Then Exit Sub

    Dim i As Long

    If 入力範囲表示 = True Then
    
        If 行位置 = True Then
            Select Case True
            Case cls終了位置 > 0
                For i = cls開始位置 To cls終了位置
                    If cls基準Sh.Cells(i, 入力確認位置).Value = "" Then
                        cls基準Sh.Rows(i).Hidden = True
                    Else
                        cls基準Sh.Rows(i).Hidden = False
                    End If
                Next
            Case cls範囲 >= 0 And cls終了位置 = 0
                For i = cls開始位置 To cls開始位置 + cls範囲
                    If cls基準Sh.Cells(i, 入力確認位置).Value = "" Then
                        cls基準Sh.Rows(i).Hidden = True
                    Else
                        cls基準Sh.Rows(i).Hidden = False
                    End If
                Next
            End Select
        Else
        
            Select Case True
            Case cls終了位置 > 0
                For i = cls開始位置 To cls終了位置
                    If cls基準Sh.Cells(入力確認位置, i).Value = "" Then
                        cls基準Sh.Columns(i).Hidden = True
                    Else
                        cls基準Sh.Columns(i).Hidden = False
                    End If
                Next
            Case cls範囲 >= 0 And cls終了位置 = 0
                For i = cls開始位置 To cls開始位置 + cls範囲
                    If cls基準Sh.Cells(入力確認位置, 1).Value = "" Then
                        cls基準Sh.Columns(i).Hidden = True
                    Else
                        cls基準Sh.Columns(i).Hidden = False
                    End If
                Next
            End Select
        
        End If
        
    Else
    
        If 行位置 = True Then
            Select Case True
            Case cls終了位置 > 0
                cls基準Sh.Range(cls基準Sh.Rows(cls開始位置), cls基準Sh.Rows(cls終了位置)).Hidden = True
            Case cls範囲 >= 0 And cls終了位置 = 0
                cls基準Sh.Rows(cls開始位置).Resize(cls範囲).Hidden = True
            End Select
        Else
        
            Select Case True
            Case cls終了位置 > 0
                cls基準Sh.Range(cls基準Sh.Columns(cls開始位置), cls基準Sh.Columns(cls終了位置)).Hidden = True
            Case cls範囲 >= 0 And cls終了位置 = 0
                cls基準Sh.Columns(cls開始位置).Resize(cls範囲).Hidden = True
            End Select
        
        End If
    
    End If

End Sub

Public Sub 行選択()

    行位置 = True

End Sub

Public Sub 列選択()

    行位置 = False

End Sub

Public Property Get シート指定() As Worksheet

    Set シート指定 = cls基準Sh

End Property

Public Property Set シート指定(ByVal シートオブジェクト As Worksheet)

    Set cls基準Sh = シートオブジェクト

End Property

Public Property Get 開始位置() As Long

    開始位置 = cls開始位置

End Property

Public Property Let 開始位置(ByVal 位置番号 As Long)

    cls開始位置 = 位置番号

End Property

Public Property Get 終了位置() As Long

    終了位置 = cls終了位置

End Property

Public Property Let 終了位置(ByVal 位置番号 As Long)

    cls範囲 = 0

    cls終了位置 = 位置番号

End Property

Public Property Get 範囲() As Long

    範囲 = cls範囲

End Property

Public Property Let 範囲(ByVal 行列数 As Long)

    cls終了位置 = 0

    cls範囲 = 行列数

End Property
