VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "リスト操作"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub 抽出(ByRef 元セル As Range, ByRef 反映セル As Range)

    Dim v As Variant, w As Variant
    
    v = 元セル.Value
    w = 反映セル.Value
    
    Dim i1 As Long, j1 As Long
    Dim i2 As Long, j2 As Long
    
    For i1 = 1 To UBound(v, 1)
    
        For i2 = 1 To UBound(w, 1)
        
            If v(i1, 1) = w(i2, 1) And v(i1, 1) <> "" And w(i2, 1) <> "" Then
            
                For j1 = 1 To UBound(v, 2)
                
                    For j2 = 1 To UBound(w, 2)
                    
                        If v(1, j1) = w(1, j2) And v(1, j1) <> "" And w(1, j2) <> "" Then
                        
                            If i2 > 1 And j2 > 1 Then
                        
                                w(i2, j2) = ""
                        
                            End If
                        
                            w(i2, j2) = v(i1, j1)
                        
                        End If
                        
                    Next
                    
                Next
                
            End If
            
        Next
        
    Next
    
    反映セル.Value = w

End Sub
