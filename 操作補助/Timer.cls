VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'オプション機能用　GetTickCountの定義
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private BufferTime As Long          ' オプション機能用 マクロの動作時間検証用
Private sumTime As Long
Private count As Long
Private sumCount As Long


Private Sub Class_Initialize()
    Application.StatusBar = False
End Sub

Private Sub Class_Terminate()
    Application.StatusBar = False
End Sub

'実行速度検証用プロシージャ

Public Sub スタート()
    
    BufferTime = GetTickCount()
    
End Sub

Public Function ストップ(Optional ByVal ポップアップ As Boolean = False, _
                                    Optional ByVal デバックプリント As Boolean = False) As String
    
    Dim 結果 As String
    
    結果 = GetTickCount() - BufferTime & " ms"
    
    If ポップアップ Then MsgBox 結果
    If デバックプリント Then Debug.Print 結果
    
    ストップ = 結果
    
End Function

Public Function ストップタイム(Optional ByVal ポップアップ As Boolean = False, _
                                    Optional ByVal デバックプリント As Boolean = False) As Long
    
    Dim 結果 As String
    
    結果 = Val(GetTickCount() - BufferTime)
    
    If ポップアップ Then MsgBox 結果
    If デバックプリント Then Debug.Print 結果
    
    ストップタイム = 結果
    
End Function

Public Sub カウンタ設定(ByVal 合計処理回数 As Long, Optional ByVal ステータスバー As Boolean = False)
    
    count = 0
    sumCount = 0
    sumTime = 0
    
    If 合計処理回数 > 0 Then
    
        sumCount = 合計処理回数
        
        If ステータスバー = True Then Application.StatusBar = "【進捗状況】" & Format(0, "@@@@@@") & " / " & Format(sumCount, "@@@@@@")
        
    End If
    
    BufferTime = GetTickCount()
    
End Sub

Public Function 進捗表示(Optional ByVal ステータスバー As Boolean = False, Optional ByVal ポップアップ As Boolean = False, _
                                    Optional ByVal デバックプリント As Boolean = False) As String
    
    Dim 通知内容 As String
    Dim 結果 As String
    
    count = count + 1
    
    結果 = Val(GetTickCount() - BufferTime)
    
    If sumCount > 0 Then
    
        sumTime = sumTime + 結果
    
        通知内容 = "【進捗状況】" & Format(count, "@@@@@@") & " / " & Format(sumCount, "@@@@@@") & " 【終了予定時刻】" & Format(DateAdd("s", (sumTime / count) * (sumCount - count) / 1000, Now()), "h時mm分ss秒") & " ( " & Format(DateAdd("s", (sumTime / count) * (sumCount - count) / 1000, 0), "h時間mm分ss秒") & " )"
        
    Else
    
        通知内容 = "事前に合計処理回数の指定がされていないため、表示できません。"
    
    End If
    
    If ステータスバー Then Application.StatusBar = 通知内容
    If ポップアップ Then MsgBox 通知内容
    If デバックプリント Then Debug.Print 通知内容
    
    進捗表示 = 通知内容
    
End Function
