VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "excQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Declare Sub sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

Private Const jQueryVer1 = "//bit.ly/13gXeNc"
Private Const jQueryVer2 = "//bit.ly/1uwTMVc"
Private Const Word1 = "javascript:(function(d,f,s){s=d.createElement('script');s.src='"
Private Const Word2 = "';s.onload=function(){f(jQuery.noConflict(1))};d.body.appendChild(s)})(document,function($){"
Private Const Word3 = "})"

Private mainIE As Object
Private cObj As Object
Private cSelector As String

Private Sub Class_Initialize()

    Set cObj = Nothing
    Set mainIE = Nothing
    cSelector = "body"

End Sub

Private Sub Class_Terminate()

    cSelector = ""
    Set cObj = Nothing
    Set mainIE = Nothing

End Sub

Private Function NewMe(ByVal myObject As Object, Optional ByVal セレクタ As String = "body") As excQuery

    Dim mySele As String

    If IsObject(myObject) Then
        Dim my As excQuery
        Set my = New excQuery
        Set NewMe = my.Init(myObject, mainIE, セレクタ)
        Set my = Nothing
    Else
        Set NewMe = Nothing
    End If

End Function

Private Function StrUnity(ByVal Str As String) As String

    StrUnity = LCase$(StrConv(Str, vbNarrow))

End Function

Private Sub wait(Optional ByVal 待機時間 As Long)

    Dim i As Long

    If 0 < 待機時間 Then
        sleep 待機時間
    Else
        Select Case "IWebBrowser2"
        Case TypeName(mainIE)
            Do
                DoEvents
            Loop Until mainIE.ReadyState = 4 And mainIE.busy = False
            fraSleep mainIE
        Case TypeName(cObj)
            Do
                DoEvents
            Loop Until mainIE.ReadyState = 4 And mainIE.busy = False
            fraSleep cObj
        End Select
    End If

End Sub

Private Sub fraSleep(ByRef inObj As Object)

    Dim myDoc As Object
    Dim myObj As Object
    Dim i As Long
    Dim frameType As Variant
    Dim frameTypes As Variant
    
    Set myDoc = inObj.document
    
    frameTypes = Array("frame", "iframe")

    On Error Resume Next

    For Each frameType In frameTypes
    
        Err.Clear

        If 0 < myDoc.getElementsByTagName(frameType).length Then
    
            If Err.Number = 0 Then
    
                For Each myObj In myDoc.getElementsByTagName(frameType)
            
                    Do
                    
                        DoEvents
                    
                    Loop Until myObj.contentWindow.document.ReadyState = "complete"
                
                    fraSleep myObj.contentWindow
            
                Next
                
            End If
        
        End If
        
    Next
        
    On Error GoTo 0

End Sub

Private Function doTrav(ByVal 走査名 As String) As excQuery

    Dim mySel As String
    Dim myVal As String

    If cSelector Like "*$*" Then
        mySel = cSelector & "." & 走査名 & "()"
        Set doTrav = NewMe(cObj, mySel)
    Else
        mySel = "$('" & cSelector & "')" & "." & 走査名 & "()"
        Set doTrav = NewMe(cObj, mySel)
    End If

End Function

Private Function setTrav(ByVal 走査名 As String, Optional ByVal セレクタ As String = "") As excQuery

    Dim mySel As String
    Dim myVal As String
    
    If セレクタ <> "" And セレクタ Like "*[^$]*" And Not (IsNumeric(セレクタ)) Then
        myVal = "'" & セレクタ & "'"
    ElseIf IsNumeric(セレクタ) Then
        myVal = セレクタ
    End If

    If セレクタ Like "*$*" Then
        mySel = cSelector & "." & 走査名 & "(" & myVal & ")"
        Set setTrav = NewMe(cObj, mySel)
    Else
        mySel = "$('" & cSelector & "')" & "." & 走査名 & "(" & myVal & ")"
        Set setTrav = NewMe(cObj, mySel)
    End If

End Function

Private Function getLength(ByVal 属性 As String) As Long

    On Error Resume Next
    
    Dim buf As String
    
    If cSelector Like "$*" Then
    
        buf = cSelector & "." & 属性
    
    Else
    
        buf = "$('" & cSelector & "')." & 属性
    
    End If

    Me.jQuery "document.body.removeAttribute('excQueryValue')", "document.body.setAttribute('excQueryValue', (" & buf & ")?" & buf & ":'')"
    getLength = Me.doc().body.getAttribute("excQueryValue")
    
    On Error GoTo 0

End Function

Private Function getAttr(ByVal 属性 As String) As String

    On Error Resume Next
    
    Dim buf As String
    
    If cSelector Like "$*" Then
    
        buf = cSelector & "." & 属性 & "()"
    
    Else
    
        buf = "$('" & cSelector & "')." & 属性 & "()"
    
    End If

    Me.jQuery "document.body.removeAttribute('excQueryValue')", "document.body.setAttribute('excQueryValue', (" & buf & ")?" & buf & ":'')"
    getAttr = Me.doc().body.getAttribute("excQueryValue")
    
    On Error GoTo 0

End Function

Private Function setAttr(ByVal 属性 As String, ByVal 値 As String) As String

    On Error Resume Next
    
    Dim buf As String
    
    If cSelector Like "$*" Then
    
        buf = cSelector & "." & 属性
    
    Else
    
        buf = "$('" & cSelector & "')." & 属性
    
    End If

    Me.jQuery buf & "('" & 値 & "')"
    
    On Error GoTo 0

End Function

Private Function onEvent(ByVal イベント名 As String, Optional ByVal 実行処理 As String = "") As String

    On Error Resume Next
    
    Dim buf As String
    
    If cSelector Like "$*" Then
    
        buf = cSelector & "." & イベント名
    
    Else
    
        buf = "$('" & cSelector & "')." & イベント名
    
    End If
    
    If 実行処理 <> "" Then 実行処理 = "function(){" & 実行処理 & "}"

    Me.jQuery buf & "(" & 実行処理 & ")"
    
    On Error GoTo 0

End Function


Public Function Init(Optional ByVal myObject As Object, Optional ByRef myIE As Object, Optional ByVal セレクタ As String = "body") As excQuery ' 暗黙設定

    If IsObject(myObject) Then
        Set cObj = myObject
        If TypeName(myObject) = "IWebBrowser2" Then
            Set mainIE = myObject
        Else
            Set mainIE = myIE
        End If
        cSelector = セレクタ
        Set Init = Me
    Else
        Set mainIE = Nothing
        Set cObj = Nothing
        Set Init = Nothing
    End If

End Function

Public Function IEVersion() As Long

    Dim FSO As Object
    Dim ver As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ver = FSO.GetFileVersion("C:\Program Files\Internet Explorer\IEXPLORE.EXE")
    
    IEVersion = Left$(ver, InStr(ver, ".") - 1) - 0
    
    Set FSO = Nothing

End Function

Public Function IE(ByVal 検索値 As String) As excQuery
Attribute IE.VB_UserMemId = 0

    Dim OBJ_FSO As Variant
    Dim OBJ_IE As Variant
    Dim objIE As Object
    Dim STR_ProgramName As String
    Dim Checker As Boolean
    Dim myURL As String
    Dim myTitle As String
    Dim i As Long
    
    検索値 = StrUnity(検索値)
    If 検索値 Like "http[s:][:/]*" Then
        myURL = 検索値
        myTitle = "_[]"
    Else
        myURL = "_[]"
        myTitle = 検索値
    End If

    On Error Resume Next

    Set OBJ_FSO = CreateObject("Scripting.FileSystemObject")
    For Each OBJ_IE In CreateObject("Shell.Application").Windows
    
        DoEvents

        STR_ProgramName = OBJ_FSO.GetFileName(OBJ_IE.FullName)
        If StrUnity(STR_ProgramName) = "iexplore.exe" Then
        
            Select Case True
            Case StrUnity(OBJ_IE.LocationURL) Like myURL & "*", StrUnity(OBJ_IE.LocationName) Like "*" & myTitle & "*"
                Set objIE = OBJ_IE
                
                Set IE = NewMe(objIE)
                
                Set OBJ_IE = Nothing
                Set OBJ_FSO = Nothing
            
                Exit Function
            End Select

        End If

    Next
    
    On Error GoTo 0

    If myTitle = "_[]" Then

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

Public Function doc() As Variant

    Select Case TypeName(cObj)
    Case "IWebBrowser2"
        Set doc = cObj.document
    Case "HTMLWindow2"
        Set doc = cObj.document
    Case Else
        Set doc = Nothing
    End Select

End Function

Public Sub Navigate(ByVal URL As String)

    On Error Resume Next

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
    
    On Error GoTo 0

End Sub

Public Sub Version()

    On Error Resume Next

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
        .Navigate "javascript:alert($.fn.jquery)"
        Do

            DoEvents

        Loop Until .busy = False
    End With
    
    On Error GoTo 0

End Sub

Public Sub jQueryVersion()

    On Error Resume Next

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
        Select Case Me.IEVersion
        Case Is < 8
            .Navigate Word1 & jQueryVer1 & Word2 & "alert($.fn.jquery)" & Word3  ' 1.11.1
        Case Else
            .Navigate Word1 & jQueryVer2 & Word2 & "alert($.fn.jquery)" & Word3  ' 2.1.1
        End Select
        Do

            DoEvents

        Loop Until .busy = False
    End With
    
    On Error GoTo 0

End Sub

Public Sub jQuery(ParamArray code() As Variant)

    On Error Resume Next

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
        Select Case Me.IEVersion
        Case Is < 8
            .Navigate Word1 & jQueryVer1 & Word2 & Join$(code, ";") & Word3  ' 1.11.1
        Case Else
            .Navigate Word1 & jQueryVer2 & Word2 & Join$(code, ";") & Word3  ' 2.1.1
        End Select
        Do

            DoEvents

        Loop Until .busy = False
    End With
    
    On Error GoTo 0

End Sub

Public Function selector(ByVal セレクタ As String) As excQuery

    Set selector = NewMe(cObj, セレクタ)

End Function

Public Function S(ByVal セレクタ As String) As excQuery

    Set S = NewMe(cObj, セレクタ)

End Function

Public Function ID_S(ByVal セレクタ As String) As excQuery

    If セレクタ Like "*$*" Then
        Set ID_S = NewMe(Nothing)
    Else
        Set ID_S = NewMe(cObj, "#" & セレクタ)
    End If

End Function

Public Function CLASS_S(ByVal セレクタ As String) As excQuery

    If セレクタ Like "*$*" Then
        Set CLASS_S = NewMe(Nothing)
    Else
        Set CLASS_S = NewMe(cObj, "." & セレクタ)
    End If

End Function

Public Function ATTR_S(ByVal セレクタ As String, Optional ByVal 属性 As String = "", Optional ByVal 属性値 As String = "", Optional ByVal 比較方法 As String = "=") As excQuery

    Dim mySel As String

    If セレクタ Like "*$*" Then
        Set ATTR_S = NewMe(Nothing)
    Else
        mySel = セレクタ
        If 属性 <> "" Then
            mySel = mySel & "[" & 属性
            If 属性値 <> "" And 比較方法 Like "*=*" Then mySel = mySel & " " & 比較方法 & " """ & 属性値 & """"
            mySel = mySel & "]"
        End If
        Set ATTR_S = NewMe(cObj, "." & mySel)
    End If

End Function

Public Function find(Optional ByVal セレクタ As String = "") As excQuery

    Set find = setTrav("find", セレクタ)

End Function

Public Function children(Optional ByVal セレクタ As String = "") As excQuery

    Set children = setTrav("children", セレクタ)

End Function

Public Function closest(Optional ByVal セレクタ As String = "") As excQuery

    Set closest = setTrav("closest", セレクタ)

End Function

Public Function parent(Optional ByVal セレクタ As String = "") As excQuery

    Set parent = setTrav("parent", セレクタ)

End Function

Public Function parents(Optional ByVal セレクタ As String = "") As excQuery

    Set parents = setTrav("parents", セレクタ)

End Function

Public Function sNext(Optional ByVal セレクタ As String = "") As excQuery

    Set sNext = setTrav("next", セレクタ)

End Function

Public Function nextAll(Optional ByVal セレクタ As String = "") As excQuery

    Set nextAll = setTrav("nextAll", セレクタ)

End Function

Public Function prev(Optional ByVal セレクタ As String = "") As excQuery

    Set prev = setTrav("prev", セレクタ)

End Function

Public Function prevAll(Optional ByVal セレクタ As String = "") As excQuery

    Set prevAll = setTrav("prevAll", セレクタ)

End Function

Public Function siblings(Optional ByVal セレクタ As String = "") As excQuery

    Set siblings = setTrav("siblings", セレクタ)

End Function

Public Function sNot(Optional ByVal セレクタ As String = "") As excQuery

    Set sNot = setTrav("not", セレクタ)

End Function


Public Function sEnd() As excQuery

    Set sEnd = doTrav("end")

End Function

Public Function first() As excQuery

    Set first = doTrav("first")

End Function

Public Function last() As excQuery

    Set last = doTrav("last")

End Function

Public Function contents() As excQuery

    Set contents = doTrav("contents")

End Function

Public Function eq(Optional ByVal 番号 As String = "0") As excQuery

    Set eq = setTrav("eq", 番号)

End Function

Public Function slice(Optional ByVal 指定番号範囲 As String = "0") As excQuery

    Set sNot = setTrav("not", 指定番号範囲)

End Function

Public Property Get val() As String

    On Error Resume Next

    val = getAttr("val")
    
    On Error GoTo 0

End Property

Public Property Let val(ByVal 値 As String)

    On Error Resume Next

    setAttr "val", 値
    
    On Error GoTo 0

End Property

Public Property Get text() As String

    On Error Resume Next

    text = getAttr("text")
    
    On Error GoTo 0

End Property

Public Property Let text(ByVal 値 As String)

    On Error Resume Next

    setAttr "text", 値
    
    On Error GoTo 0

End Property

Public Property Get html() As String

    On Error Resume Next

    html = getAttr("html")
    
    On Error GoTo 0

End Property

Public Property Let html(ByVal 値 As String)

    On Error Resume Next

    setAttr "html", 値
    
    On Error GoTo 0

End Property

Public Property Get prop() As String

    On Error Resume Next

    prop = getAttr("prop")
    
    On Error GoTo 0

End Property

Public Property Let prop(ByVal 値 As String)

    On Error Resume Next

    setAttr "prop", 値
    
    On Error GoTo 0

End Property

Public Property Get css() As String

    css = getAttr("css")

End Property

Public Property Let css(ByVal 値 As String)

    On Error Resume Next

    setAttr "css", 値
    
    On Error GoTo 0

End Property

Public Property Get length() As String

    On Error Resume Next

    length = getLength("length")
    
    On Error GoTo 0

End Property

Public Function click(Optional ByVal 実行処理 As String = "")

    On Error Resume Next

    onEvent "click", 実行処理
    
    On Error GoTo 0

End Function

Public Function dblclick(Optional ByVal 実行処理 As String = "")

    On Error Resume Next

    onEvent "dblclick", 実行処理
    
    On Error GoTo 0

End Function

Public Function change(Optional ByVal 実行処理 As String = "")

    On Error Resume Next

    onEvent "change", 実行処理
    
    On Error GoTo 0

End Function

Public Function blur(Optional ByVal 実行処理 As String = "")

    On Error Resume Next

    onEvent "blur", 実行処理
    
    On Error GoTo 0

End Function

Public Function focus(Optional ByVal 実行処理 As String = "")

    On Error Resume Next

    onEvent "focus", 実行処理
    
    On Error GoTo 0

End Function

Public Function keydown(Optional ByVal 実行処理 As String = "")

    On Error Resume Next

    onEvent "keydown", 実行処理
    
    On Error GoTo 0

End Function

Public Function keypress(Optional ByVal 実行処理 As String = "")

    On Error Resume Next

    onEvent "keypress", 実行処理
    
    On Error GoTo 0

End Function

Public Function keyup(Optional ByVal 実行処理 As String = "")

    On Error Resume Next

    onEvent "keyup", 実行処理
    
    On Error GoTo 0

End Function


Public Function submit(Optional ByVal 実行処理 As String = "")

    On Error Resume Next

    onEvent "submit", 実行処理
    
    On Error GoTo 0

End Function
