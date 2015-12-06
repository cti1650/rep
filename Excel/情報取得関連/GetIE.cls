VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private mainIE As Object
Private cObj As Object
Private cColl As Variant

Private Sub Class_Initialize()

    Set cObj = Nothing
    Set cColl = Nothing
    Set mainIE = Nothing

End Sub

Private Sub Class_Terminate()

    Set cObj = Nothing
    Set cColl = Nothing
    Set mainIE = Nothing

End Sub

Private Function NewMe(ByVal myObject As Object) As GetIE

    If IsObject(myObject) Then
        Dim my As GetIE
        Set my = New GetIE
        Set NewMe = my.Init(myObject, mainIE)
        Set my = Nothing
    Else
        Set NewMe = Nothing
    End If

End Function

Private Function AllElement(ByVal 検索値 As String, ByVal inObj As Variant, Optional ByVal Tag名 As String = "*") As Collection

Dim i As Long, j As Long
Dim myName As Variant, myID As Variant, myClass As Variant, myTitle As Variant
Dim cll As New Collection
Dim myObj As Variant
Dim myFra As Variant
Dim myDoc As Variant
Dim myCll As Variant
Dim myFlag As Boolean

Select Case TypeName(inObj)
Case "IWebBrowser2"
Case "HTMLWindow2"
Case Else
    Set AllElement = cll
    Exit Function
End Select

    If Tag名 = "" Then Tag名 = "*"

    Set myDoc = inObj.document
    
    If TypeName(myDoc.getElementsByTagName(Tag名)) <> "Nothing" Then

        If 0 < myDoc.getElementsByTagName(Tag名).Length Then
    
            For Each myObj In myDoc.getElementsByTagName(Tag名)
        
                On Error Resume Next
        
                myFlag = False
                myName = ""
                myID = ""
                myClass = ""
                myTitle = ""
            
                Err.Clear
                myName = myObj.Name
                If Err.Number <> 0 Then myFlag = True
            
                Err.Clear
                myID = myObj.ID
                If Err.Number <> 0 Then myFlag = True
                
                Err.Clear
                myClass = myObj.className
                If Err.Number <> 0 Then myFlag = True
                
                Err.Clear
                myTitle = myObj.innerText
                If Err.Number <> 0 Then myFlag = True
            
                On Error GoTo 0

                If myName Like 検索値 Or myID Like 検索値 Or myClass Like 検索値 Or myTitle Like 検索値 Then
                
                    DoEvents

                    If TypeName(myObj) <> "Empty" Then cll.Add myObj
                
                End If
            
            Next
        
        End If
        
    End If

    If TypeName(myDoc.frames) = "HTMLWindow2" Then
    
        On Error Resume Next

        Err.Clear

        If 0 < myDoc.getElementsByTagName("frame").Length Then
    
            If Err.Number = 0 Then
    
                For Each myObj In myDoc.getElementsByTagName("frame")
                
                    DoEvents
            
                    Set myCll = AllElement(検索値, myObj.contentWindow, Tag名)
                    
                    If 0 < myCll.Count Then
        
                        For Each myFra In myCll

                            If TypeName(myFra) <> "Empty" Then cll.Add myFra
                
                        Next
                        
                    End If
            
                Next
                
            End If
        
        End If
        
        Err.Clear
        
        If 0 < myDoc.getElementsByTagName("iframe").Length Then
        
            If Err.Number = 0 Then

                For Each myObj In myDoc.getElementsByTagName("iframe")
                
                    DoEvents

                    Set myCll = AllElement(検索値, myObj.contentWindow, Tag名)
                    
                    If 0 < myCll.Count Then
        
                        For Each myFra In myCll

                            If TypeName(myFra) <> "Empty" Then cll.Add myFra
                
                        Next
                    
                    End If
            
                Next
                
            End If
        
        End If
        
        On Error GoTo 0
        
    End If

    Set AllElement = cll

End Function

Public Function Init(Optional ByVal myObject As Object, Optional ByRef myIE As Object) As GetIE ' 暗黙設定

    If IsObject(myObject) Then
        Set cObj = myObject
        If TypeName(myObject) = "IWebBrowser2" Then
            Set mainIE = myObject
        Else
            Set mainIE = myIE
        End If
        
        Set Init = Me
    Else
        Set mainIE = Nothing
        Set cObj = Nothing
        Set Init = Nothing
    End If

End Function

'フレームのDocumentを取得する関数
Public Function Frame(ByVal Name As String) As GetIE

Dim objFrames As Variant
Dim frameDoc As Object
Dim i As Integer

Select Case TypeName(cObj)
Case "IWebBrowser2"

    Set objFrames = cObj.document.frames

    For i = 0 To objFrames.Length - 1
        Set frameDoc = objFrames(i)
        If frameDoc.Name = Name Then
            Exit For
        End If
    Next
    
Case "HTMLWindow2"

    Set objFrames = cObj.frames

    For i = 0 To objFrames.Length - 1
        Set frameDoc = objFrames(i)
        If frameDoc.Name = Name Then
            Exit For
        End If
    Next
    
Case Else
    
    Set frameDoc = Nothing
    
End Select

Set Frame = NewMe(frameDoc)

End Function

Public Function IE(ByVal 検索値 As String) As GetIE
Attribute IE.VB_UserMemId = 0

    Dim OBJ_FSO As Variant
    Dim OBJ_IE As Variant
    Dim objIE As Object
    Dim STR_ProgramName As String
    Dim Checker As Boolean
    Dim myURL As String
    Dim myTitle As String
    Dim i As Long
    
    If 検索値 Like "http://*" Or 検索値 Like "https://*" Then
        myURL = 検索値
        myTitle = "_[]"
    Else
        myURL = "_[]"
        myTitle = 検索値
    End If

    Set OBJ_FSO = CreateObject("Scripting.FileSystemObject")
    For Each OBJ_IE In CreateObject("Shell.Application").Windows
    
        DoEvents

        STR_ProgramName = OBJ_FSO.GetFileName(OBJ_IE.FullName)
        If LCase(STR_ProgramName) = "iexplore.exe" Then
        
            If OBJ_IE.LocationURL Like myURL & "*" Or OBJ_IE.LocationName Like "*" & myTitle & "*" Then
                
                Set objIE = OBJ_IE
                
                Set IE = NewMe(objIE)
                
                Set OBJ_IE = Nothing
                Set OBJ_FSO = Nothing
            
                Exit Function

            End If

        End If

    Next

    If myURL <> "_[]" Then

        Set OBJ_IE = CreateObject("InternetExplorer.Application")
    
        With OBJ_IE
            .Left = 0
            .Top = 0
            .Visible = True
            .Navigate myURL
            Do

                DoEvents

            Loop Until .ReadyState = 4 And .busy = False
        End With
    
        Set objIE = OBJ_IE
        fraSleep objIE
        Set IE = NewMe(objIE)

    Else

        Set IE = NewMe(Nothing)
        
        ' MsgBox "IEオブジェクトが取得出来ませんでした", vbOKOnly, "取得不可"

    End If

    Set OBJ_IE = Nothing
    Set OBJ_FSO = Nothing

End Function

Public Function Ele(ByVal Keyword As String, Optional ByVal Tag名 As String = "*") As GetIE

    Select Case TypeName(cObj)
    Case "IWebBrowser2"
    Case "HTMLWindow2"
    Case Else
        Set Ele = NewMe(Nothing)
        Exit Function
    End Select
    
    Dim i As Long
    Dim myCll As Collection
    
    If Tag名 = "" Then Tag名 = "*"
    
    Set myCll = AllElement(Keyword, cObj, Tag名)
    
    i = myCll.Count
    
    If 0 < i Then
    
        Set Ele = NewMe(myCll)
        
    Else
    
        Set Ele = NewMe(Nothing)
        
        ' MsgBox Keyword & " エレメントは存在しません"
        
    End If

End Function

Public Property Get Item(Optional ByVal 番号 As Long = -1) As GetIE

    Select Case TypeName(cObj)
    Case "DispHTMLElementCollection"
        If 0 < 番号 And 番号 <= Me.Length Then
            Set Item = NewMe(cObj(番号))
        Else
            Set Item = NewMe(cObj)
        End If
    Case "Collection"
        If 0 < 番号 And 番号 <= Me.Length Then
            Set Item = NewMe(cObj(番号))
        Else
            Set Item = NewMe(cObj)
        End If
    Case Else
        Set Item = NewMe(Nothing)
        ' MsgBox "Val プロパティを使用できません"
        Exit Property
    End Select
    
End Property

Public Function objType() As Variant

    objType = TypeName(cObj)
    
End Function

Public Function Obj() As Object

    Set Obj = cObj
    
End Function

Public Function Doc() As Variant

    Select Case TypeName(cObj)
    Case "IWebBrowser2"
        Set Doc = cObj.document
    Case "HTMLWindow2"
        Set Doc = cObj.document
    Case Else
        Set Doc = Nothing
    End Select

End Function

Public Sub 待機(Optional ByVal 待機時間 As Long)

    Dim i As Long

    If 0 < 待機時間 Then
        Sleep 待機時間
    Else
        If TypeName(mainIE) = "IWebBrowser2" Then
            Do

                DoEvents

            Loop Until mainIE.ReadyState = 4 And mainIE.busy = False
            fraSleep mainIE
        Else
            If TypeName(cObj) = "IWebBrowser2" Then
                Do
                
                    DoEvents
                    
                Loop Until mainIE.ReadyState = 4 And mainIE.busy = False
                fraSleep cObj
            End If
        End If
    End If

End Sub

Private Sub fraSleep(ByRef inObj As Object)

    Dim myDoc As Object
    Dim myObj As Object
    Dim i As Long
    
    Set myDoc = inObj.document

    On Error Resume Next

    Err.Clear

    If 0 < myDoc.getElementsByTagName("frame").Length Then
    
        If Err.Number = 0 Then
    
            For Each myObj In myDoc.getElementsByTagName("frame")
            
                Do
                    
                    DoEvents
                    
                Loop Until myObj.contentWindow.document.ReadyState = "complete"
                
                fraSleep myObj.contentWindow
            
            Next
                
        End If
        
    End If
        
    Err.Clear
        
    If 0 < myDoc.getElementsByTagName("iframe").Length Then
        
        If Err.Number = 0 Then
    
            For Each myObj In myDoc.getElementsByTagName("iframe")
            
                Do
                    
                    DoEvents

                Loop Until myObj.contentWindow.document.ReadyState = "complete"
                
                fraSleep myObj.contentWindow
            
            Next
                
        End If
        
    End If
        
    On Error GoTo 0

End Sub


Public Sub Navigate(ByVal URL As String)

    Dim myObj As Object

    Select Case TypeName(cObj)
    Case "IWebBrowser2"
        Set myObj = cObj
    Case "HTMLWindow2"
        Set myObj = cObj
    Case "HTMLFrameElement"
        Set myObj = cObj.contentWindow
    Case Else
        ' MsgBox "Navigate メソッドを使用できません"
        Exit Sub
    End Select
    
    With myObj
        .Navigate URL
        Do

            DoEvents

        Loop Until .ReadyState = 4 And .busy = False
        fraSleep myObj
    End With

End Sub

Public Sub jQuery(ParamArray code() As Variant)

    Dim myObj As Object

    Select Case TypeName(cObj)
    Case "IWebBrowser2"
        Set myObj = cObj
    Case "HTMLWindow2"
        Set myObj = cObj
    Case "HTMLFrameElement"
        Set myObj = cObj.contentWindow
    Case Else
        ' MsgBox "Navigate メソッドを使用できません"
        Exit Sub
    End Select
    
    myObj.Navigate "javascript:(function(d,f,s){s=d.createElement(""script"");s.src=""//j.mp/1bPoAXq"";s.onload=function(){f(jQuery.noConflict(1))};d.body.appendChild(s)})(document,function($){" & Join$(code, ";") & "})"

End Sub

Public Sub Q(ParamArray code() As Variant)

    Dim myObj As Object

    Select Case TypeName(cObj)
    Case "IWebBrowser2"
        Set myObj = cObj
    Case "HTMLWindow2"
        Set myObj = cObj
    Case "HTMLFrameElement"
        Set myObj = cObj.contentWindow
    Case Else
        ' MsgBox "Navigate メソッドを使用できません"
        Exit Sub
    End Select
    
    myObj.Navigate "javascript:(function(d,f,s){s=d.createElement(""script"");s.src=""//j.mp/1bPoAXq"";s.onload=function(){f(jQuery.noConflict(1))};d.body.appendChild(s)})(document,function($){" & Join$(code, ";") & "})"

End Sub

Public Sub JavaScript(ByVal Run As String)

    Dim myObj As Object

    Select Case TypeName(cObj)
    Case "IWebBrowser2"
        Set myObj = cObj
    Case "HTMLWindow2"
        Set myObj = cObj
    Case "HTMLFrameElement"
        Set myObj = cObj.contentWindow
    Case Else
        ' MsgBox "Navigate メソッドを使用できません"
        Exit Sub
    End Select
    
    With myObj
        .Navigate "javascript:(function(){" & Run & "})()"
        Do

            DoEvents

        Loop Until .ReadyState = 4 And .busy = False
        fraSleep myObj
    End With

End Sub

Public Property Get Length() As Long

    Select Case TypeName(cObj)
    Case "DispHTMLElementCollection"
        Length = cObj.Length
        Exit Property
    Case "Collection"
        Length = cObj.Count
        Exit Property
    Case Else
        Length = 0
        ' MsgBox "Length プロパティを使用できません"
        Exit Property
    End Select

End Property

Public Property Get Name() As Variant

    Dim buf As String
    On Error Resume Next
    
    Err.Clear
    buf = ""
    
    buf = cObj.Name
    If Err.Number <> 0 Then buf = ""
    
    On Error GoTo 0
    
    Name = buf
    
End Property

Public Property Get Tag() As Variant

    Dim buf As String
    On Error Resume Next
    
    Err.Clear
    buf = ""
    
    buf = cObj.tagName
    If Err.Number <> 0 Then buf = ""
    
    On Error GoTo 0
    
    Tag = buf
    
End Property

Public Property Get Class() As Variant

    Dim buf As String
    On Error Resume Next
    
    Err.Clear
    buf = ""
    
    buf = cObj.className
    If Err.Number <> 0 Then buf = ""
    
    On Error GoTo 0
    
    Class = buf
    
End Property

Public Property Get ID() As Variant

    Dim buf As String
    On Error Resume Next
    
    Err.Clear
    buf = ""
    
    buf = cObj.ID
    If Err.Number <> 0 Then buf = ""
    
    On Error GoTo 0
    
    ID = buf
    
End Property

Public Property Get Val() As Variant

    Dim a As Variant
    Dim i As Long
    
    If TypeName(cObj) = "DispHTMLElementCollection" Or TypeName(cObj) = "Collection" Then
    
        For Each a In cObj

            DoEvents

            If outVal(a, Val) <> 0 Then MsgBox "Val プロパティを使用できません"
        Next
        
    Else
    
        If outVal(cObj, Val) <> 0 Then MsgBox "Val プロパティを使用できません"
        
    End If
    
End Property

Public Property Let Val(ByVal vNewValue As Variant)

    Dim a As Variant
    Dim i As Long
    
    If TypeName(cObj) = "DispHTMLElementCollection" Or TypeName(cObj) = "Collection" Then
    
        For Each a In cObj
            DoEvents
            If inVal(a, vNewValue) <> 0 Then MsgBox "Val プロパティを使用できません"
        Next
        
    Else
    
        If inVal(cObj, vNewValue) <> 0 Then MsgBox "Val プロパティを使用できません"
        
    End If
    
End Property

Private Function inVal(ByVal myObj As Object, ByVal myValue As Variant) As Long

    On Error Resume Next
    
    Err.Clear
    
    Select Case TypeName(myObj)
    Case "HTMLOptionElement"
        myObj.Value = myValue
    Case "HTMLButtonElement"
        myObj.Value = myValue
    Case "HTMLSelectElement"
        myObj.Value = myValue
    Case "HTMLInputElement"
        myObj.Value = myValue
    Case "HTMLTableCell"
        myObj.innerText = myValue
    Case "HTMLTextAreaElement"
        myObj.innerText = myValue
    End Select

    inVal = Err.Number
    
    On Error GoTo 0

End Function

Private Function outVal(ByVal myObj As Object, ByRef myValue As Variant) As Long

    On Error Resume Next
    
    Err.Clear
    
    Select Case TypeName(myObj)
    Case "HTMLOptionElement"
        myValue = myObj.Value
    Case "HTMLButtonElement"
        myValue = myObj.Value
    Case "HTMLSelectElement"
        myValue = myObj.Value
    Case "HTMLInputElement"
        myValue = myObj.Value
    Case "HTMLTableCell"
        myValue = myObj.innerText
    Case "HTMLTextAreaElement"
        myValue = myObj.innerText
    End Select

    outVal = Err.Number
    
    On Error GoTo 0

End Function

Public Property Get Index() As Variant

    Dim a As Variant
    Dim i As Long
    
    If TypeName(cObj) = "DispHTMLElementCollection" Or TypeName(cObj) = "Collection" Then
    
        For Each a In cObj
            DoEvents
            If outIndex(a, Index) <> 0 Then MsgBox "Index プロパティを使用できません"
        Next
        
    Else
    
        If outIndex(cObj, Index) <> 0 Then MsgBox "Index プロパティを使用できません"
        
    End If
    
End Property

Public Property Let Index(ByVal vNewIndex As Variant)

    Dim a As Variant
    Dim i As Long
    
    If TypeName(cObj) = "DispHTMLElementCollection" Or TypeName(cObj) = "Collection" Then
    
        For Each a In cObj
            DoEvents
            If inIndex(a, vNewIndex) <> 0 Then MsgBox "Index プロパティを使用できません"
        Next
        
    Else
    
        If inIndex(cObj, vNewIndex) <> 0 Then MsgBox "Index プロパティを使用できません"
        
    End If
    
End Property

Private Function inIndex(ByVal myObj As Object, ByVal myIndex As Variant) As Long

    On Error Resume Next
    
    Err.Clear
    
    Select Case TypeName(myObj)
    Case "HTMLOptionElement"
        myObj.Index = myIndex
    Case "HTMLSelectElement"
        myObj.SelecedIndex = myIndex
    End Select

    inIndex = Err.Number
    
    On Error GoTo 0

End Function

Private Function outIndex(ByVal myObj As Object, ByRef myIndex As Variant) As Long

    On Error Resume Next
    
    Err.Clear
    
    Select Case TypeName(myObj)
    Case "HTMLSelectElement"
        myIndex = myObj.SelecedIndex
    End Select

    outIndex = Err.Number
    
    On Error GoTo 0

End Function

Public Property Get URL() As Variant

    Dim a As Variant
    Dim i As Long
    
    If TypeName(cObj) = "DispHTMLElementCollection" Or TypeName(cObj) = "Collection" Then
    
        For Each a In cObj
            DoEvents
            If outURL(a, URL) <> 0 Then MsgBox "URL プロパティを使用できません"
        Next
        
    Else
    
        If outURL(cObj, URL) <> 0 Then MsgBox "URL プロパティを使用できません"
        
    End If
    
End Property

Public Property Let URL(ByVal vNewURL As Variant)

    Dim a As Variant
    Dim i As Long
    
    If TypeName(cObj) = "DispHTMLElementCollection" Or TypeName(cObj) = "Collection" Then
    
        For Each a In cObj
            DoEvents
            If inURL(a, vNewURL) <> 0 Then MsgBox "URL プロパティを使用できません"
        Next
        
    Else
    
        If inURL(cObj, vNewURL) <> 0 Then MsgBox "URL プロパティを使用できません"
        
    End If
    
End Property

Private Function inURL(ByVal myObj As Object, ByVal myURL As Variant) As Long

    On Error Resume Next
    
    Err.Clear
    
    Select Case TypeName(myObj)
    Case "HTMLAnchorElement"
        myObj.herf = myURL
    Case "HTMLFrameElement"
        myObj.contentWindow.Location.href = myURL
    Case "IWebBrowser2"
        myObj.LocationURL = myURL
    Case "HTMLWindow2"
        myObj.Location.href = myURL
    End Select

    inURL = Err.Number
    
    On Error GoTo 0

End Function

Private Function outURL(ByVal myObj As Object, ByRef myURL As Variant) As Long

    On Error Resume Next
    
    Err.Clear
    
    Select Case TypeName(myObj)
    Case "HTMLAnchorElement"
        myURL = myObj.herf
    Case "HTMLFrameElement"
        myURL = myObj.contentWindow.Location.href
    Case "IWebBrowser2"
        myURL = myObj.LocationURL
    Case "HTMLWindow2"
        myURL = myObj.Location.href
    End Select

    outURL = Err.Number
    
    On Error GoTo 0

End Function

Public Property Get Title() As Variant

    Dim a As Variant
    Dim i As Long
    
    If TypeName(cObj) = "DispHTMLElementCollection" Or TypeName(cObj) = "Collection" Then
    
        For Each a In cObj
            DoEvents
            If outTitle(a, Title) <> 0 Then MsgBox "Title プロパティを使用できません"
        Next
        
    Else
    
        If outTitle(cObj, Title) <> 0 Then MsgBox "Title プロパティを使用できません"
        
    End If
    
End Property

Public Property Let Title(ByVal vNewTitle As Variant)

    Dim a As Variant
    Dim i As Long
    
    If TypeName(cObj) = "DispHTMLElementCollection" Or TypeName(cObj) = "Collection" Then
    
        For Each a In cObj
            DoEvents
            If inTitle(a, vNewTitle) <> 0 Then MsgBox "Title プロパティを使用できません"
        Next
        
    Else
    
        If inTitle(cObj, vNewTitle) <> 0 Then MsgBox "Title プロパティを使用できません"
        
    End If
    
End Property

Private Function inTitle(ByVal myObj As Object, ByVal myTitle As Variant) As Long

    On Error Resume Next
    
    Err.Clear
    
    Select Case TypeName(myObj)
    Case "HTMLAnchorElement"
        myObj.innerText = myTitle
    Case "HTMLFrameElement"
        myObj.contentWindow.document.Title = myTitle
    Case "IWebBrowser2"
        myObj.LocationName = myTitle
    Case "HTMLWindow2"
        myObj.document.Title = myTitle
    End Select

    inTitle = Err.Number
    
    On Error GoTo 0

End Function

Private Function outTitle(ByVal myObj As Object, ByRef myTitle As Variant) As Long

    On Error Resume Next
    
    Err.Clear
    
    Select Case TypeName(myObj)
    Case "HTMLAnchorElement"
        myTitle = myObj.innerText
    Case "HTMLFrameElement"
        myTitle = myObj.contentWindow.document.Title
    Case "IWebBrowser2"
        myTitle = myObj.LocationName
    Case "HTMLWindow2"
        myTitle = myObj.document.Title
    End Select

    outTitle = Err.Number
    
    On Error GoTo 0

End Function

Public Property Get OptionText() As Variant

    Dim a As Variant
    Dim i As Long
    
    If TypeName(cObj) = "DispHTMLElementCollection" Or TypeName(cObj) = "Collection" Then
    
        For Each a In cObj
            DoEvents
            If outOptionText(a, OptionText) <> 0 Then MsgBox "OptionText プロパティを使用できません"
        Next
        
    Else
    
        If outOptionText(cObj, OptionText) <> 0 Then MsgBox "OptionText プロパティを使用できません"
        
    End If
    
End Property

Public Property Let OptionText(ByVal vNewOptionText As Variant)

    Dim a As Variant
    Dim i As Long
    
    If TypeName(cObj) = "DispHTMLElementCollection" Or TypeName(cObj) = "Collection" Then
    
        For Each a In cObj
            DoEvents
            If inOptionText(a, vNewOptionText) <> 0 Then MsgBox "OptionText プロパティを使用できません"
        Next
        
    Else
    
        If inOptionText(cObj, vNewOptionText) <> 0 Then MsgBox "OptionText プロパティを使用できません"
        
    End If
    
End Property

Private Function inOptionText(ByVal myObj As Object, ByVal myOptionText As Variant) As Long

    On Error Resume Next
    
    Err.Clear
    
    Select Case TypeName(myObj)
    Case "HTMLOptionElement"
        myObj.innerText = myOptionText
    End Select

    inOptionText = Err.Number
    
    On Error GoTo 0

End Function

Private Function outOptionText(ByVal myObj As Object, ByRef myOptionText As Variant) As Long

    On Error Resume Next
    
    Err.Clear
    
    Select Case TypeName(myObj)
    Case "HTMLOptionElement"
        myOptionText = myObj.innerText
    End Select

    outOptionText = Err.Number
    
    On Error GoTo 0

End Function

Public Sub Click()

    Dim a As Variant
    Dim i As Long
    
    If TypeName(cObj) = "DispHTMLElementCollection" Or TypeName(cObj) = "Collection" Then
    
        For Each a In cObj
            DoEvents
            If onClick(a) <> 0 Then MsgBox "Click メソッドを使用できません"
        Next
        
    Else
    
        If onClick(cObj) <> 0 Then MsgBox "Click メソッドを使用できません"
        
    End If
    
End Sub

Private Function onClick(ByVal myObj As Object) As Long

    On Error Resume Next
    
    Err.Clear
    
    Select Case TypeName(myObj)
    Case "HTMLAnchorElement"
        myObj.Click
        Me.待機
    Case "HTMLFormElement"
        myObj.submit
        Me.待機
    Case "HTMLButtonElement"
        myObj.Click
        Me.待機
    End Select

    onClick = Err.Number
    
    On Error GoTo 0

End Function

Public Sub Reset()

    Dim a As Variant
    Dim i As Long
    
    If TypeName(cObj) = "DispHTMLElementCollection" Or TypeName(cObj) = "Collection" Then
    
        For Each a In cObj
            DoEvents
            If onReset(a) <> 0 Then MsgBox "Reset メソッドを使用できません"
        Next
        
    Else
    
        If onReset(cObj) <> 0 Then MsgBox "Reset メソッドを使用できません"
        
    End If
    
End Sub

Private Function onReset(ByVal myObj As Object) As Long

    On Error Resume Next
    
    Err.Clear
    
    Select Case TypeName(myObj)
    Case "HTMLFormElement"
        myObj.Reset
        Me.待機
    End Select

    onReset = Err.Number
    
    On Error GoTo 0

End Function

