VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyJpegExif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/////////////////////////////////////
'   JpegExifクラス
'       ver 1.00    読み出し専用
'/////////////////////////////////////

'使用方法
'JpegExif クラス
' プロパティ (すべて読み取り専用)
'  Name String ファイル名 (拡張子含む)
'  Path String フルパス (ファイル名含む)
'  FileSize Long ファイルサイズ (byte単位)
'  DateTimeLastModified Date 最終更新日時
'  DateTimeOriginal Date 撮影日時
'  Make String メーカー名
'  Model String 機種名
'  ExifVersion String Exifバージョン
'  ExposureTime String 露出時間 (秒単位)。分数形式 例) 1/30
'  Flash Boolean 発光したか否
'  FNumber String Fナンバー。分数形式 例) 28/10
'  ImageLength Long 画像縦サイズ (ピクセル単位)
'  ImageWidth Long 画像横サイズ (ピクセル単位)
' メソッド
'  InitSet
'   ファイルフルパスを受け取り､プロパティなどの値取得を行う｡
'   引数 String ファイルフルパス
'   戻り値 Integer
'     0  JpegExifファイルである
'     2  Jpegファイルであるが､Exifではない
'    -1  ファイルが見つからない
'    -2  ファイルが開けない
'    -3  Jpegファイルではない
'    -11 格納データにエラーがあるみたい


Option Explicit

Const TAG_DATETIMEORIGINAL = 36867
Const TAG_MAKE = 271
Const TAG_MODEL = 272
Const TAG_EXIFVERSION = 36864
Const TAG_EXPOSURETIME = 33434
Const TAG_FLASH = 37385
Const TAG_FNUMBER = 33437
Const TAG_EXIF_IFD = 34665

Const サイズ微調整 = 4
Const 枠微調整 = 2

Private cName As String
Private cPath As String         'ファイル名まで含めたフルパス
Private cFileSize As Long       'byte単位
Private cDateTimeLastModified As Date
Private cDateTimeOriginal As Date
Private cMake As String
Private cModel As String
Private cExifVersion As String
Private cExposureTime As String
Private cFlash As Boolean
Private cFNumber As String
Private cImageWidth As Long
Private cImageLength As Long
Private cRet As Long

Private cEndian As Integer      'IFDにおける整数格納形式 1:BigEndian  -1:LittleEndian
Private cAddrTifHead As Long    'Tifヘッドアドレス
                                'IFDのオフセットはすべてここを基準とする
                                
Enum MyCode
    画像なし_0 = 0
    画像情報あり_1 = 1
    画像あり_2 = 2
End Enum
Enum Mypoint
    中央 = 1
    左上 = 2
End Enum


'   インスタンスが作成されるときInitializeイベントが発生する
Private Sub Class_Initialize()
    cEndian = 1
End Sub

'***** プロパティ *****

'***** InstSet結果
Public Property Get Ret() As MyCode
    Ret = cRet
End Property

'***** InstSet結果
Public Property Get ErrCheck() As Boolean
    If cRet > 0 Then
        ErrCheck = True
    Else
        ErrCheck = False
    End If
End Property

'***** ファイル名
Public Property Get ファイル名() As String
    ファイル名 = cName
End Property

'***** フルパス
Public Property Get フルパス() As String
    フルパス = cPath
End Property

'***** ファイルサイズ(byte単位)
Public Property Get ファイルサイズ() As Long
    ファイルサイズ = cFileSize
End Property

'***** ファイル最終更新日付
Public Property Get 最終更新日時() As Date
    最終更新日時 = cDateTimeLastModified
End Property

'***** 撮影日付
Public Property Get 撮影日時() As Date
    撮影日時 = cDateTimeOriginal
End Property

Public Property Get 作成日時() As Date
    If cDateTimeOriginal = "00:00:00" Then
        作成日時 = cDateTimeLastModified
    Else
        作成日時 = cDateTimeOriginal
    End If
End Property

'***** メーカー
Public Property Get メーカー() As String
    メーカー = cMake
End Property

'***** モデル
Public Property Get 機種名() As String
    機種名 = cModel
End Property

'***** Exifバージョン
Public Property Get Exifバージョン() As String
    Exifバージョン = cExifVersion
End Property

'***** 露出時間
Public Property Get 露出時間() As String
    露出時間 = cExposureTime
End Property

'***** 発光
Public Property Get 発光() As Boolean
    発光 = cFlash
End Property

'***** Fナンバー
Public Property Get Fナンバー() As String
    Fナンバー = cFNumber
End Property

'***** 画像サイズたて (ピクセル単位)
Public Property Get 縦幅() As Long
    縦幅 = cImageLength
End Property

'***** 画像サイズよこ (ピクセル単位)
Public Property Get 横幅() As Long
    横幅 = cImageWidth
End Property

'***** メソッド *****

'*********************************************************************
'   ファイル名を渡しそのファイルの実在確認と情報取得を行う。
'   fullPathはコンストラクタの引数にしたいところだが、
'   引数付きコンストラクタを作成できないのでこのメソッドを用意した。
'   ret 0: JpegExifファイルである
'       2: JpegファイルであるがExifではない (画像サイズは取得できる)
'      -1: ファイルが存在しない
'      -2: (存在はするが)オープンできない
'      -3: Jpegファイルではない
'     -11: 格納データのフォーマットにエラーがある
'*********************************************************************
Public Function 画像情報入力(ByVal fullPath As String) As Integer
    Dim Ret As Integer
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(fullPath) Then
        With fso.GetFile(fullPath)
            cName = .Name
            cPath = .Path   'ファイル名も含めたフルパス
            cDateTimeLastModified = .DateLastModified
        End With
        
        Ret = readJPEG()
    Else
        Ret = -1
    End If
    Set fso = Nothing
    
    画像情報入力 = Ret
End Function

'----- プライベート関数 -----

'----------------------------------------------------
'   Jpegファイルを読む
'   ret 0: JpegExifである
'       2: JpegファイルではあるがExif形式ではない
'      -2: ファイルオープン失敗
'      -3: Jpegファイルではない
'     -11: データエラーあり
'----------------------------------------------------
Private Function readJPEG() As Integer
    Dim Ret As Integer
    Dim i As Integer
    Dim fnum As Integer
    Dim b1 As Byte
    Dim b2(0 To 1) As Byte
    Dim b4(0 To 3) As Byte
    Dim ba() As Byte
    Dim fPointer As Long    'ファイル先頭を0とする
    Dim segSize As Long
    Dim segStart As Long
    Dim offset As Long
    
    Ret = 1
    fnum = FreeFile
    On Error GoTo Er_readJPEG
    Open cPath For Binary Access Read As #fnum
    On Error GoTo 0
    cFileSize = LOF(fnum)
    
    '1. SOI(StartOfImage)マーカ の確認
    '   ファイル先頭が ffd8 であること
    fPointer = Seek(fnum) - 1   'seek関数はファイル先頭を1とするので
    If fPointer + 2 > cFileSize Then
        '読み出し不可
        Ret = -3
    Else
        Get #fnum, , b2
        fPointer = fPointer + 2
        If Not (b2(0) = &HFF And b2(1) = &HD8) Then
            Ret = -3
        End If
    End If
    If Ret < 0 Then GoTo Ex_readJPEG
    
    '2. 最初の「データ部を持つJPEGマーカ」を読む
    '   SOIマーカの直後
    '   ff MM SS SS
    ' MM: マーカ識別子。Exif用のAPP1マーカの場合は e1
    ' SS SS: このマーカセグメントのサイズ(上位のSSを基点とする) 符合なしBigEndian
    If fPointer + 4 > cFileSize Then
        '読み出し不可
        Ret = -3
    Else
        ReDim ba(0 To 3)
        Get #fnum, , ba
        fPointer = fPointer + 4
        If ba(0) = &HFF Then
            'JPEGマーカ
            segStart = fPointer - 2
            segSize = CLng(ba(2)) * 256 + CLng(ba(3))
            If ba(1) = &HE1 Then
                '3. Exif証明書の確認
                'APP1マーカであるなら続く6バイトが固定
                ' 45 78 69 66 00 00
                ' E  x  i  f
                If fPointer + 6 > cFileSize Then
                    Ret = -11
                Else
                    ReDim ba(0 To 5)
                    Get #fnum, , ba
                    fPointer = fPointer + 6
                    If Not (ba(0) = &H45 And ba(1) = &H78 And ba(2) = &H69 And ba(3) = &H66 _
                        And ba(4) = 0 And ba(5) = 0) Then
                        Ret = -11
                    End If
                End If
            Else
                'JPEGではあるが、Exif情報は持っていない
                Ret = 2
            End If
        Else
            'JPEGファイルではない
            Ret = -3
        End If
    End If
    If Ret < 0 Then GoTo Ex_readJPEG
    
    If Ret = 1 Then
        'Exif情報の読み出し
        '   Exif情報はTIFFのタグ形式。Exif証明書の直後から始まる
        '4. Tifヘッダの読み出し
        '   NN NN TT TT XX XX XX XX
        '   NN NN : 整数格納形式 4d4d=BigEndian 4949=LittleEndian
        '   TT TT : TIFF番号(整数。何を意味するかは不明) 値は &h2a (42)
        '   XX XX XX XX : 最初のIFD(ImageFileDirectory)へのオフセット
        '                 通常はTifヘッダの直後だから8
        If fPointer + 8 > cFileSize Then
            Ret = -11
        Else
            'Endian取得
            Get #fnum, , b2
            If b2(0) = &H49 And b2(1) = &H49 Then
                cEndian = -1 '既定値は1
            End If
            
            'TIFF番号の確認
            Get #fnum, , b2
            If bytesUnsignedToLong(b2) <> 42 Then
                'TIFFヘッダではない
                Ret = -11
            Else
                'Tiffタグ認定ができたと言うことで
                cAddrTifHead = fPointer
                'IFD0へのオフセット値取得
                Get #fnum, , b4
                offset = bytesUnsignedToLong(b4)
                If offset > 0 Then
                    'IFD0 処理
                    offset = readIFD(fnum, offset)
                    '次のIFD1はサムネイルに関する情報なので読まない
                    'IFD1で連鎖終了
                Else
                    'オフセット値がおかしい。
                    '負数の場合はオフセット値が扱える範囲を超えているということだが、
                    '取得済みのファイルサイズ(符号あり長整数)を超えるというのはおかしい。
                    '0の場合は全く先に進めない。
                    Ret = -11
                End If
            End If
        End If
        If Ret < 0 Then GoTo Ex_readJPEG
    End If
    
    '5. 画像サイズ取得
    '   JPEGマーカ(ffc0)より取得する
    '   APP1タグとしてImageWidth, ImageLengthは定義されているが、
    '   JPEGマーカとの重複になるので記載しないことになっている。
    'ffc0マーカまでマーカ連鎖をたどる
    'JPEGマーカフォーマット
    ' ff MM SS SS
    ReDim ba(0 To 3)
    Do Until Ret < 0 Or segSize = 0
        fPointer = segStart + segSize
        If fPointer + 4 > cFileSize Then
            Ret = -11
        Else
            Get #fnum, fPointer + 1, ba
            fPointer = fPointer + 4
            If ba(0) = &HFF Then
                If ba(1) = &HC0 Then
                    '目的のJPEGマーカに到達
                    'サイズに続く部分はマーカごとに異なる。
                    'ffcoマーカの場合は
                    'PP LL LL WW WW XX XX...
                    ' PP: サンプル精度。通常は08
                    ' LL LL: イメージの縦サイズ、符合なし整数、BigEndian
                    ' WW WW: イメージの横サイズ、同上
                    '5バイトだけ読み出す
                    If fPointer + 5 > cFileSize Then
                        Ret = -11
                    Else
                        Get #fnum, , b1 'PP 空読み
                        Get #fnum, , ba
                        cImageLength = CLng(ba(0)) * 256 + CLng(ba(1))
                        cImageWidth = CLng(ba(2)) * 256 + CLng(ba(3))
                        segSize = 0
                    End If
                Else
                    segStart = fPointer - 2
                    segSize = CLng(ba(2)) * 256 + CLng(ba(3))
                End If
            Else
                Ret = -11
            End If
        End If
    Loop
    
Ex_readJPEG:
    On Error Resume Next
    Close fnum
    On Error GoTo 0
    readJPEG = Ret
    Exit Function
Er_readJPEG:
    'ファイルオープン失敗
    Ret = -2
    Resume Ex_readJPEG
End Function

'---------------------------------------------------------------------------
'   ひとつのIFDを読み出し処理する
'   ret: 次のIFDへのオフセット (次がない場合は0)
'   IFDのフォーマット
'     PP PP [エントリ群] NN NN
'       PP PP: タグ個数
'       [エントリ群]: [エントリ](12byte)×タグ個数
'       NN NN: 次のIFDのオフセット(Tifヘッダ基点)
'   エントリのフォーマット
'     TT TT YY YY UU UU UU UU VV VV VV VV
'       TT TT: タグ(番号)
'       YY YY: データタイプ
'       UU UU UU UU: 値の個数
'       VV VV VV VV: 値または値へのオフセット(Tifヘッダ基点)
'---------------------------------------------------------------------------
Private Function readIFD(ByVal fnum As Integer, ByVal offset As Long) As Long
    Dim Ret As Long
    Dim b2(0 To 1) As Byte
    Dim b4(0 To 3) As Byte
    Dim fPointer As Long    '0基点
    Dim tagCount As Long
    Dim tagInfo_Id() As Long
    Dim tagInfo_ValType() As Long
    Dim tagInfo_ValCount() As Long
    Dim tagInfo_Val() As Long
    Dim i As Long
    Dim nextIFD As Long
    Dim tmpStr As String
    
    fPointer = cAddrTifHead + offset
    
    'タグ数の取得
    Get #fnum, fPointer + 1, b2
    tagCount = bytesUnsignedToLong(b2)
    If tagCount > 0 Then
    
        'タグのエントリ情報取得
        '   seekポインタはエントリ情報ブロック先頭(タグ数の直後)にある
        ReDim tagInfo_Id(0 To tagCount - 1)
        ReDim tagInfo_ValType(0 To tagCount - 1)
        ReDim tagInfo_ValCount(0 To tagCount - 1)
        ReDim tagInfo_Val(0 To tagCount - 1)
        
        For i = 0 To tagCount - 1
            'タグID取得
            Get #fnum, , b2
            tagInfo_Id(i) = bytesUnsignedToLong(b2)
            '値の種類
            Get #fnum, , b2
            tagInfo_ValType(i) = bytesUnsignedToLong(b2)
            '値の個数
            Get #fnum, , b4
            tagInfo_ValCount(i) = bytesUnsignedToLong(b4)
            '値または値へのオフセット
            Get #fnum, , b4
            Select Case tagInfo_ValType(i)
            Case 1  '1バイト符号なし整数
                tagInfo_Val(i) = CLng(b4(0))
            Case 2  'ASCII文字列へのオフセット
                tagInfo_Val(i) = bytesUnsignedToLong(b4)
            Case 3  '2バイト符号なし整数
                b2(0) = b4(0)
                b2(1) = b4(1)
                tagInfo_Val(i) = bytesUnsignedToLong(b2)
            Case 4  '4バイト符号なし整数
                tagInfo_Val(i) = bytesUnsignedToLong(b4)
            Case 5  '4バイト符号なし整数2個。オフセット
                tagInfo_Val(i) = bytesUnsignedToLong(b4)
            Case 7  '未定義のバイト列。タグにより扱い方が異なる
                If tagInfo_Id(i) = TAG_EXIFVERSION Then
                    cExifVersion = Chr(b4(0)) & Chr(b4(1)) & Chr(b4(2)) & Chr(b4(3))
                End If
            Case 9  '4バイト符号あり整数
                tagInfo_Val(i) = bytesSignedToLong(b4)
            Case 10 '4バイト符号あり整数2個。オフセット
                tagInfo_Val(i) = bytesUnsignedToLong(b4)
            Case Else
                tagInfo_Val(i) = bytesUnsignedToLong(b4)
            End Select
        Next i
        
        '次のIFDへのオフセット取得(2byte!!)
        Get #fnum, , b2
        Ret = bytesUnsignedToLong(b2)
        
        'エントリを順次処理
        For i = 0 To tagCount - 1
            
            If tagInfo_Id(i) = TAG_EXIF_IFD Then
                'Exif IFDを処理する
                '取得した値はIFDへのオフセット
                nextIFD = tagInfo_Val(i)
                Do While nextIFD > 0
                    nextIFD = readIFD(fnum, nextIFD)
                Loop
            Else
                '取得の必要なTagのみ処理する
                Select Case tagInfo_Id(i)
                Case TAG_DATETIMEORIGINAL
                    '日付データの格納形式
                    'yyyy:mm:dd HH:MM:SS
                    If tagInfo_ValType(i) = 2 And tagInfo_Val(i) > 0 Then
                        tmpStr = readASCIIAt(cAddrTifHead + tagInfo_Val(i), tagInfo_ValCount(i), fnum)
                        If Len(tmpStr) = 19 Then
                            '   日付の区切りをスラッシュに変更
                            Mid(tmpStr, 5, 1) = "/"
                            Mid(tmpStr, 8, 1) = "/"
                            On Error Resume Next
                            cDateTimeOriginal = CDate(tmpStr)
                            On Error GoTo 0
                        End If
                    End If
                Case TAG_MAKE
                    'メーカー
                    If tagInfo_ValType(i) = 2 And tagInfo_Val(i) > 0 Then
                        cMake = readASCIIAt(cAddrTifHead + tagInfo_Val(i), tagInfo_ValCount(i), fnum)
                    End If
                Case TAG_MODEL
                    'モデル
                    If tagInfo_ValType(i) = 2 And tagInfo_Val(i) > 0 Then
                        cModel = readASCIIAt(cAddrTifHead + tagInfo_Val(i), tagInfo_ValCount(i), fnum)
                    End If
                Case TAG_EXIFVERSION
                    'Exifバージョン
                    'すでに処理済み
                Case TAG_EXPOSURETIME
                    '露出時間
                    If tagInfo_ValType(i) = 5 And tagInfo_Val(i) > 0 Then
                        cExposureTime = readRationalAt(cAddrTifHead + tagInfo_Val(i), fnum)
                    End If
                Case TAG_FLASH
                    '発光 最下位ビットのオン・オフで判断する。
                    If tagInfo_Val(i) Mod 2 <> 0 Then cFlash = True
                Case TAG_FNUMBER
                    'Fナンバー
                    If tagInfo_ValType(i) = 5 And tagInfo_Val(i) > 0 Then
                        cFNumber = readRationalAt(cAddrTifHead + tagInfo_Val(i), fnum)
                    End If
                End Select
            End If
        Next i
    End If
    readIFD = Ret
End Function

'-----------------------------------------------------------------------------
'   タグValue読み出し Type: 5(Rational)対応
'   指定位置から符合なし4バイト整数２個を読み出す。分子・分母に相当。
'   結果は２数をそれぞれ文字列に変換し、スラッシュで連結したもの。
'       ex) 28/10
'   分子>=0かつ分母>0　が成り立たなければ空文字列を返す
'   引数 pos: 読み出し開始位置(ファイル先頭を0とする)
'        fnum: 対象ファイルのオープン番号
'-----------------------------------------------------------------------------
Private Function readRationalAt(ByVal pos As Long, ByVal fnum As Integer) As String
    Dim Ret As String
    Dim b4(0 To 3) As Byte
    Dim tmpLong As Long
    Dim tmpRet As String
    
    If pos + 8 <= cFileSize Then
        Get #fnum, pos + 1, b4
        tmpLong = bytesUnsignedToLong(b4)
        If tmpLong >= 0 Then
            tmpRet = CStr(tmpLong)
            Get #fnum, , b4
            tmpLong = bytesUnsignedToLong(b4)
            If tmpLong > 0 Then
                Ret = tmpRet & "/" & CStr(tmpLong)
            End If
        End If
    End If
    readRationalAt = Ret
End Function

'--------------------------------------------------------------
'   タグValue読み出し Type: 2(ASCII)対応
'   指定位置からASCII文字としてlength個読み出す。
'   ただし、Nullが出現した位置で読み出しを終了し、
'   Nullの直前までを戻り値とする。
'   引数 pos: 読み出し開始位置(ファイル先頭を0とする)
'        length: 読み出しバイト数
'        fnum: 対象ファイルのオープン番号
'--------------------------------------------------------------
Private Function readASCIIAt(ByVal pos As Long, ByVal length As Long, ByVal fnum As Integer) As String
    Dim Ret As String
    Dim ba() As Byte
    Dim i As Long
    
    If length > 0 Then
        If pos + length <= cFileSize Then
            ReDim ba(0 To length - 1)
            Get #fnum, pos + 1, ba
            For i = 0 To length - 1
                If ba(i) = 0 Then Exit For
                Ret = Ret & Chr(ba(i))
            Next i
        End If
    End If
    readASCIIAt = Ret
End Function

'-----------------------------------------------------------------------------
'   符号なし整数とみなすバイト配列を4バイト符号あり整数に変換
'   引数の配列は0基底とする。
'   4次元を超える配列はインデクス0～3までを用いる。
'   4次元の場合、戻り値が符号あり4バイト整数なので最上位ビットがonの場合は、
'   オーバーフローになってしまう。この場合は-1 を返す。
'   すなわち、2^31以上の格納値は扱えないということ。
'-----------------------------------------------------------------------------
Private Function bytesUnsignedToLong(ba() As Byte) As Long
    Dim Ret As Long
    Dim i As Integer
    Dim iMax As Integer
    
    iMax = UBound(ba)
    If iMax < 1 Then Exit Function
    If iMax > 3 Then iMax = 3
    
    On Error GoTo Er_bytesUnsignedToLong
    If cEndian = -1 Then
        For i = 0 To iMax
            Ret = Ret + ba(i) * (256 ^ i)
        Next i
    Else
        For i = 0 To iMax
            Ret = Ret + ba(i) * (256 ^ (iMax - i))
        Next i
    End If
    
Ex_bytesUnsignedToLong:
    On Error GoTo 0
    bytesUnsignedToLong = Ret
    Exit Function
Er_bytesUnsignedToLong:
    Ret = -1
    Resume Ex_bytesUnsignedToLong
End Function

'---------------------------------------------------------------------
'   符合あり整数とみなすバイト配列を4バイト符合あり整数に変換
'   引数の配列は0基底とする。
'   4次元を超える配列はインデクス0～3までを用いる。
'---------------------------------------------------------------------
Private Function bytesSignedToLong(ba() As Byte) As Long
    Dim Ret As Long
    Dim bt() As Integer
    Dim i As Integer
    Dim iMax As Integer
    Dim s As Long
    Dim carry As Integer
    
    iMax = UBound(ba)
    If iMax < 1 Then Exit Function
    If iMax > 3 Then iMax = 3
        
    'Little Endianとして処理を行うために配列を移し替える
    ReDim bt(0 To iMax)
    If cEndian = -1 Then
        For i = 0 To iMax
            bt(i) = CInt(ba(i))
        Next i
    Else
        For i = 0 To iMax
            bt(i) = CInt(ba(iMax - i))
        Next i
    End If
    
    '最上位バイトの最上位ビットを調べ結果値の符号を求める
    If bt(iMax) < 128 Then
        s = 1
    Else
        s = -1
    End If
    
    '負の場合のみ符号反転処理実施
    If s = -1 Then
        '最下位バイトから順に繰り上げを考慮しながら2の補数に変換
        carry = 0 '繰り上げ
        For i = 0 To iMax
            bt(i) = bt(i) + carry   'この操作でByte型の範囲を逸脱する可能性
                                    'があるのでInteger型に変換してある
            If bt(i) > 0 Then
                bt(i) = &H100 - bt(i)
                carry = 1
            End If
        Next i
    End If
    
    For i = 0 To iMax
        Ret = Ret + bt(i) * (256 ^ i)
    Next i
    bytesSignedToLong = Ret * s
    
End Function

Public Function 選択セル画像貼付(Optional ByVal 選択範囲行数 As Long = 0, Optional ByVal 選択範囲列数 As Long = 0, Optional ByVal ポジション As Mypoint) As Boolean
'
' 画像貼付 Macro
' マクロ記録日 : 2013/4/16  ユーザー名 : 阪上 雄一
' 作成：阪上 雄一
'


    '初期設定
    Dim ファイル名 As String
    Dim Cell_Row As Single
    Dim Cell_Column As Single
    Dim Cell_H As Single
    Dim Cell_W As Single
    Dim Cell_T As Single
    Dim Cell_L As Single
    Dim Cell_Ratio As Single
    Dim Obj_Ratio As Single

    '選択されているオブジェクトがセルか判断する
    If TypeName(Selection) = "Range" Or TypeName(Selection) = "Ranges" Then
    
        If 選択範囲行数 <> 0 And 選択範囲列数 <> 0 Then
                If Selection.Rows.Count <> 選択範囲行数 Or Selection.Columns.Count <> 選択範囲列数 Then Exit Function
        ElseIf 選択範囲行数 <> 0 And 選択範囲列数 = 0 Then
                If Selection.Rows.Count <> 選択範囲行数 Then Exit Function
        ElseIf 選択範囲行数 = 0 And 選択範囲列数 <> 0 Then
                If Selection.Columns.Count <> 選択範囲列数 Then Exit Function
        Else
                Exit Function
        End If
        
    
    '画面描画を無効にする
    Application.ScreenUpdating = False
    
    '選択されているセルの情報を取得する
    With Selection
        Cell_Row = .Row
        Cell_Column = .Column
        Cell_H = .Height
        Cell_W = .Width
        Cell_T = .Top
        Cell_L = .Left
        Cell_Ratio = .Height / .Width     '高さ／幅
    End With
    
    ファイル名 = ファイル選択
    
    If ファイル名 <> "" And Right(ファイル名, 3) = "jpg" Or ファイル名 <> "" And Right(ファイル名, 3) = "JPG" Or Right(ファイル名, 4) = "jpeg" Then
        '画像ファイルの情報を書込みExif情報が含まれるか確認する
        cRet = Me.画像情報入力(ファイル名)
    
        With ActiveSheet.Pictures.Insert(ファイル名)
            With .ShapeRange
                '画像をセルの幅または高さに合わせる
                .LockAspectRatio = msoTrue      '一定比率
                Obj_Ratio = .Height / .Width
                If Cell_Ratio > Obj_Ratio Then
                    'セル枠の方が高さの比が大きい（セル幅に合わせる）
                    .Width = Cell_W
                Else
                    'セル枠の方が高さの比が小さい（セル高に合わせる）
                    .Height = Cell_H
                End If
            End With
            '容量少なくするために一度画像を切り取る
            .Cut
        End With
        '画質を落とした状態でシート上に画像を貼り付ける
        ActiveSheet.PasteSpecial Format:="図 (JPEG)"
        '貼り付けた画像を操作する
        With ActiveSheet.DrawingObjects(ActiveSheet.DrawingObjects.Count)
            If Cell_Ratio > Obj_Ratio Then
                'セル枠の中に収まる様に画像位置を調整する
                .Width = .Width - サイズ微調整
                .Left = Cell_L + (Cell_W - .Width) / 2
                If ポジション = 中央 Then
                    .Top = Cell_T + (Cell_H - .Height) / 2
                ElseIf ポジション = 左上 Then
                    .Top = Cell_T + 枠微調整
                Else
                    .Top = Cell_T + (Cell_H - .Height) / 2
                End If
            Else
                'セル枠の中に収まる様に画像位置を調整する
                .Height = .Height - サイズ微調整
                .Top = Cell_T + (Cell_H - .Height) / 2
                If ポジション = 中央 Then
                    .Left = Cell_L + (Cell_W - .Width) / 2
                ElseIf ポジション = 左上 Then
                    .Left = Cell_L + 枠微調整
                Else
                    .Left = Cell_L + (Cell_W - .Width) / 2
                End If
            End If
        End With
    
    End If
    
        '画面描画を有効にする
        Application.ScreenUpdating = True
    
    End If
    
    If cRet > 0 Then
        選択セル画像貼付 = True
    Else
        選択セル画像貼付 = False
    End If
    
End Function

Public Function 選択範囲チェック()

        Dim S_Row As Long
        Dim S_Column As Long
        Dim buf As String
        Dim buf2 As String
        Dim buf3 As String
        Dim buf4 As String
        Dim CB As New DataObject
        
        buf = "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf & vbCrLf & _
                  "    Dim a As New MyJpegExif" & vbCrLf & vbCrLf & _
                  "    a.選択セル画像貼付 "
        buf2 = ", "
        buf3 = ", 中央" & vbCrLf & vbCrLf & _
                   "    Set a = Nothing" & vbCrLf & vbCrLf & _
                   "End Sub"

        If TypeName(Selection) = "Range" Then
                With Selection
                        S_Row = .Rows.Count
                        S_Column = .Columns.Count
                End If
                buf4 = buf & S_Row & buf2 & S_Column & buf3
        With CB
                .SetText buf4        ''変数のデータをDataObjectに格納する
                .PutInClipboard     ''DataObjectのデータをクリップボードに格納する
        End With
        MsgBox "クリップボードへの貼付が完了しました。" & vbCrLf & vbCrLf & buf4
        End If

End Function

Private Function ファイル選択()

    Dim myFilePath As String
    myFilePath = Application.GetOpenFilename(FileFilter:="画像ファイル,*.jpg;*.JPG;*.jpeg", Title:="貼り付ける画像を選択")
    If myFilePath = "False" Then
        ファイル選択 = ""
        Exit Function
    Else
        ファイル選択 = myFilePath
    End If

End Function


