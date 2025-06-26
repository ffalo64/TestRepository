VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImscLib12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'このライブラリに書いてあることはきちんと読んでおいてください
'読まずにやっちゃうと、思わぬバグが発生する恐れがあります。

'DATAFILENAMEには各自で保存用ファイル名を入れても構いません。

'ライブラリについて
'ライブラリは、全てのゲームに共通して付属させるモジュールです。
'ライブラリには、主に音楽ファイルの再生や、データの保存など、多くのゲームで使いたい関数を入れます。


'現在、起こってはならないこと。
'最小のほうがいいスコアで、凄く大きいスコアになってしまう。
'(例:マインスイーパのようなゲームで、ハイスコア99999999とか。)
'普通はありえないことやけど、もしなってしまったらバグるんで、制限付けてください。
'RankKを使わない人には関係ないです。

Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Enum RankKind
    bigger = 0
    smaller = 1
End Enum

Private Type SOUNDDATA
    Number As Long '通し番号
    FileName As String '再生するファイルの名前
    RePlay As Boolean '繰り返し再生をするかどうか
End Type
'音楽再生に使用する構造体です。


Dim Sound() As SOUNDDATA '音楽変数

Private Type REGDATA 'データ用の構造体
    Str As String   '記録するもの
    Key As String   '記録呼び出し用のキー
    Name As String
End Type

Dim Data() As REGDATA 'データ変数

'これがセーブする場所の名前を決める所。
'今の場合は「Data」という名前のimrファイルができる。
Private Const DATAFILENAME As String = "Data.imr"
'Dataを自分のゲームにあった名前に変えてくださいね。

Private Const LAUNCHERNAME As String = "数研ゲーム集Vol.10"
'ゲーム終了時にランチャーを開くために使用します。

Private Const Pi As Double = 3.14159265358979




Private Sub Class_Initialize() 'Form_Loadと同じようなもの

    ReDim Sound(0) '定義しなおし
    Call GetData 'データ読み込み
    
End Sub

Public Sub SoundSet(Number As Long, FileName As String, Optional RePlay As Boolean = False)
    '音楽関連/////////////////////////////////////////////////////////////////////////////////////////////////////////////////音楽関連
    
    ReDim Preserve Sound(Max(UBound(Sound), Number))
    
    With Sound(Number)
        
        .Number = Number
        .FileName = FileName
        .RePlay = RePlay
        If .RePlay = False Then
            Call mciSendString("open """ & Sound(Number).FileName & """ alias " & Sound(Number).Number, "", 0, 0)
        Else
            Call mciSendString("open """ & Sound(Number).FileName & """ alias " & Sound(Number).Number & " type mpegvideo", "", 0, 0)
        End If
    
    End With

End Sub

Public Sub PlaySound(Number As Long)
                                   
    Dim ErrNum As Long
    Dim Buf As String
    Buf = Space(255)
    
    If Sound(Number).RePlay = False Then
        ErrNum = mciSendString("play """ & Sound(Number).Number & """ from 0 notify", "", 0, 0)
    Else
        ErrNum = mciSendString("play """ & Sound(Number).Number & """ repeat", "", 0, 0)
    End If
    
    Call mciGetErrorString(ErrNum, Buf, 255)
    Debug.Print Buf
    
End Sub

Public Sub StopSound(Number As Long)
    
    Call mciSendString("stop """ & Sound(Number).Number & """ notify", "", 0, 0)
    
End Sub

Public Sub AllStopSound()
    
    Dim i As Long
    For i = 0 To UBound(Sound)
        Call mciSendString("stop """ & Sound(i).Number & """", "", 0, 0)
    Next i

End Sub
Public Sub EndSound() 'ゲームを終わるときはこっちをUnloadに入れて、完全に消して下さい

    Dim i As Long
    For i = 0 To UBound(Sound)
        Call mciSendString("stop """ & Sound(i).Number & """", "", 0, 0)
        Call mciSendString("close """ & Sound(i).Number & """", "", 0, 0)
    Next i

End Sub

Public Sub VolumeSound(Number As Long, Volume As Long) '音量を変更できます。値の大きさに注意してください
    Call mciSendString("setaudio """ & Sound(Number).Number & """ volume to " & CStr(Volume), "", 0, 0)
End Sub

Public Sub DeleteSound()
If Sound(0).FileName <> "" Then Call EndSound
ReDim Sound(0) '新たに一連の曲や効果音を読み込みなおしたいとき、再利用するために使えます。
End Sub

Public Function GetLength(Number As String) As Double

    Dim RetBuffer As String * 20 '曲の長さを取得します
    Dim MCICommandString As String

    MCICommandString = "status """ & Sound(Number).Number & """ length"
    Call mciSendString(MCICommandString, RetBuffer, Len(RetBuffer), 0)

    GetLength = Val(RetBuffer) / 1000

End Function



Public Function GetPosition(Number As Long) As Double

    Dim RetBuffer As String * 20 '曲の現在の再生位置を返します
    Dim MCICommandString As String

    MCICommandString = "status """ & Sound(Number).Number & """ position"
    Call mciSendString(MCICommandString, RetBuffer, Len(RetBuffer), 0)

    GetPosition = Val(RetBuffer) / 1000


End Function

Public Sub SetPosition(Number As Long, Position As Double)

    Dim NewPosition As Long '曲を指定した位置から開始できます。

    NewPosition = Position * 1000
    'Call mciSendString("stop """ & Sound(Number).Number & """", "", 0, 0)
    Call mciSendString("seek """ & Sound(Number).Number & """ to " & NewPosition, "", 0, 0)
    Call mciSendString("play """ & Sound(Number).Number & """", "", 0, 0)

End Sub

Public Function Max(x1 As Long, x2 As Long) As Long '大きい数字を返す
Max = IIf(x1 > x2, x1, x2)
End Function

Public Function Min(x1 As Long, x2 As Long) As Long '小さい数字を返す
Min = IIf(x1 < x2, x1, x2)
End Function

Private Function GetData() As Boolean 'セーブデータ系////////////////////////////////////////////////////////////////セーブデータ系

    Dim i As Long 'データ取り込み
    Dim s() As String
    On Error GoTo FirstSet 'エラーが起こるということは
    'まだ何も設定されてないということでして
    'じゃあ、初期データを設定しましょう
    
    'エラーが起こらない→既に一度データが書き込まれている
    'まず、0番目のデータを定義しなおして、1からのデータを読み込む
    'Data(0)という箱は使用していない（多分）
    
    ReDim Data(0)
    
    'imrファイルを開く
    'ファイル番号は1
    
    Open App.Path & "\" & DATAFILENAME For Input As #1
    
    Do Until EOF(1) 'ファイルの読み込みが終わるまでループ
    'EOF(1)はファイル1の読み込みが終わっていなければFalse,終わっていればTrueを返す
    
        i = i + 1
        ReDim Preserve Data(i) '1番以降のDataを設定しなおす。
        Input #1, Data(i).Str 'ファイル1から読み込む
    Loop
    Close #1 'ファイル1を閉じる
    GoTo NormalSet 'FirstSetは飛ばす
    
FirstSet:

    Call FirstSet
    
    GetData = True
    Exit Function
    
NormalSet:

    For i = 1 To UBound(Data)
        s = Split(Data(i).Str, "/")
        'Data(i).Name = Mid(Data(i).Str, InStr(1, Data(i).Str, "@") + 1)
        'Data(i).Key = Mid(Left(Data(i).Str, InStr(1, Data(i).Str, "@") - 1), InStr(1, Data(i).Str, "/") + 1)
        'Data(i).Str = Left(Data(i).Str, InStr(1, Data(i).Str, "/") - 1)
        
        Data(i).Str = s(0): Data(i).Key = s(1): Data(i).Name = s(2)
    Next i
        
    GetData = True
    
End Function

Private Sub FirstSet()
    Dim i As Long
    On Error Resume Next
'/////////////////////////////////////////////////////////////////////////////////////////////
    'この下のData()のなかに使いたい分だけ数字を入れてください。
    '例)10を入れると10まで使えます。
    '20以上使いたかったら適当に増やしてください。
'/////////////////////////////////////////////////////////////////////////////////////////////
    ReDim Data(20)
    
    Data(1).Key = "1":           Data(1).Str = "20":              Data(1).Name = "10"
    Data(2).Key = "2":           Data(2).Str = "20":              Data(2).Name = "10"
    Data(3).Key = "3":           Data(3).Str = "10":              Data(3).Name = "2"
    Data(4).Key = "4":           Data(4).Str = "3":              Data(4).Name = "10"
    Data(5).Key = "5":           Data(5).Str = "0":              Data(5).Name = "1"
    Data(6).Key = "6":            Data(6).Str = "1":              Data(6).Name = "0"
    Data(7).Key = "7":            Data(7).Str = "0":              Data(7).Name = "10"
    Data(8).Key = "8":            Data(8).Str = "0":              Data(8).Name = ""
    Data(9).Key = "9":            Data(9).Str = "0":              Data(9).Name = ""
    Data(10).Key = "10":           Data(10).Str = "0":             Data(10).Name = ""
    Data(11).Key = "11":           Data(11).Str = "0":             Data(11).Name = ""
    Data(12).Key = "12":           Data(12).Str = "0":             Data(12).Name = ""
    Data(13).Key = "13":           Data(13).Str = "0":             Data(13).Name = ""
    Data(14).Key = "14":           Data(14).Str = "200":             Data(14).Name = ""
    Data(15).Key = "15":           Data(15).Str = "2":             Data(15).Name = ""
    Data(16).Key = "16":           Data(16).Str = "":             Data(16).Name = ""
    Data(17).Key = "17":           Data(17).Str = "":             Data(17).Name = ""
    Data(18).Key = "18":           Data(18).Str = "":             Data(18).Name = ""
    Data(19).Key = "19":           Data(19).Str = "":             Data(19).Name = ""
    Data(20).Key = "20":           Data(20).Str = "":             Data(20).Name = ""

    For i = 1 To UBound(Data) 'セーブする文字列の中に、
    '半角/があるとバグる（セーブの区切りに半角/を使用しているので）。
    'よって、半角/を@^_^@kouki@^_^@に書き換える。(Replace)
        
        Data(i).Key = Replace(Data(i).Key, "/", "@^_^@kouki@^_^@")
        Data(i).Str = Replace(Data(i).Str, "/", "@^_^@kouki@^_^@")
        Data(i).Name = Replace(Data(i).Name, "/", "@^_^@kouki@^_^@")
        
    Next i

End Sub

Public Function ReadText(Key As String, Optional Name As Boolean = False) As String

    Dim i As Long
    If Name = False Then
        For i = 1 To UBound(Data)
            If Data(i).Key = Key Then
                ReadText = Data(i).Str
                ReadText = Replace(ReadText, "@^_^@kouki@^_^@", "/")
                Exit For
            End If
        Next i
    Else
        For i = 1 To UBound(Data)
            If Data(i).Key = Key Then
                ReadText = Data(i).Name
                ReadText = Replace(ReadText, "@^_^@kouki@^_^@", "/")
                Exit For
            End If
        Next i

    End If

End Function

Public Sub WriteText(Str As String, Key As String, Optional Name As Boolean = False)

    Dim i As Long
    Str = Replace(Str, "/", "@^_^@kouki@^_^@")
    If Name = False Then
        For i = 1 To UBound(Data)
            If Data(i).Key = Key Then
                Data(i).Str = Str
                Exit Sub
            End If
        Next i
    Else
        For i = 1 To UBound(Data)
            If Data(i).Key = Key Then
                Data(i).Name = Str
                Exit Sub
            End If
        Next i
    End If
    ReDim Preserve Data(i)
    
    If Name = False Then
        Data(i).Str = Str
    Else
        Data(i).Name = Str
    End If
    
    Data(i).Key = Key

End Sub

Private Sub SaveData()

    Dim i As Long
    
    Open App.Path & "\" & DATAFILENAME For Output As #1
    For i = 1 To UBound(Data)
        Write #1, Data(i).Str & "/" & Data(i).Key & "/" & Data(i).Name
    Next i
    Close #1

End Sub

Private Sub OpenLauncher()

    Dim hWnd As Long
    hWnd = FindWindow(vbNullString, LAUNCHERNAME)
    If hWnd <> 0 Then
        Call OpenIcon(hWnd)
        Call SetForegroundWindow(hWnd)
    End If

End Sub


Private Sub Class_Terminate()

    Call SaveData
    Call EndSound
    Call OpenLauncher

End Sub

Public Sub ForBitblt(hDC As Long, Left As Long, Top As Long, pic As PictureBox)
                                                     'ラクに描画する系////////////////ぜひ使うべき！！！////////////////////ラクに描画する系

Call BitBlt(hDC, Left, Top, pic.ScaleWidth / 2, pic.ScaleHeight, pic.hDC, pic.ScaleWidth / 2, 0, vbSrcAnd)
Call BitBlt(hDC, Left, Top, pic.ScaleWidth / 2, pic.ScaleHeight, pic.hDC, 0, 0, vbSrcPaint)

End Sub

Public Sub ForDrawText(hDC As Long, Sentence As String, x1 As Long, y1 As Long, x2 As Long, y2 As Long, Optional wFormat As Long = 37)

Dim dtRect As RECT
With dtRect
    .Left = x1
    .Top = y1
    .Right = x2
    .Bottom = y2
End With

Call DrawText(hDC, Sentence, LenB(StrConv(Sentence, vbFromUnicode)), dtRect, wFormat)

End Sub
Public Sub ForStretchblt(hDC As Long, Left As Long, Top As Long, Percentage As Long, pic As PictureBox, Optional HorizonalTurn As Boolean = False, Optional PerpendicularlyTurn As Boolean = False)

Dim FWidth As Long
Dim FHeight As Long

Dim FLeft As Long
Dim FTop As Long

If HorizonalTurn = False Then
    FWidth = pic.ScaleWidth / 2 / 100 * Percentage
    FLeft = Left
Else:
    FWidth = -pic.ScaleWidth / 2 / 100 * Percentage
    FLeft = Left - FWidth
End If

If PerpendicularlyTurn = False Then
    FHeight = pic.ScaleHeight / 100 * Percentage
    FTop = Top
Else:
    FHeight = -pic.ScaleHeight / 100 * Percentage
    FTop = Top - FHeight
End If



Call StretchBlt(hDC, FLeft, FTop, FWidth, FHeight, pic.hDC, pic.ScaleWidth / 2, 0, pic.ScaleWidth / 2, pic.ScaleHeight, vbSrcAnd)
Call StretchBlt(hDC, FLeft, FTop, FWidth, FHeight, pic.hDC, 0, 0, pic.ScaleWidth / 2, pic.ScaleHeight, vbSrcPaint)

End Sub
Public Sub Rectblt(hDC As Long, Left As Long, Right As Long, Top As Long, Bottom As Long, pic As PictureBox, Optional HorizonalTurn As Boolean = False, Optional PerpendicularlyTurn As Boolean = False)

Dim FWidth As Long
Dim FHeight As Long

Dim FLeft As Long
Dim FTop As Long

If HorizonalTurn = False Then
    FWidth = Right - Left
    FLeft = Left
Else:
    FWidth = Left - Right
    FLeft = Right
End If

If PerpendicularlyTurn = False Then
    FHeight = Bottom - Top
    FTop = Top
Else:
    FHeight = Top - Bottom
    FTop = Bottom
End If


Call StretchBlt(hDC, FLeft, FTop, FWidth, FHeight, pic.hDC, pic.ScaleWidth / 2, 0, pic.ScaleWidth / 2, pic.ScaleHeight, vbSrcAnd)
Call StretchBlt(hDC, FLeft, FTop, FWidth, FHeight, pic.hDC, 0, 0, pic.ScaleWidth / 2, pic.ScaleHeight, vbSrcPaint)


End Sub

Public Sub ForAlphaBlend(hDC As Long, Left As Long, Top As Long, pic As PictureBox, Persent As Long)
'Public Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDCSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal nBlendFunc As Long) As Long
'AlphaBlendをつかうには、この一文をいつものばしょに宣言して下さい
'AlphaBlendをつかうと、半透明に塗ることができます
'なかなか便利ですよ
Dim FPersent As Long
FPersent = Persent
If FPersent >= 256 Then FPersent = 255
If FPersent < 0 Then FPersent = 0

'Call AlphaBlend(hDC, Left, Top, pic.ScaleWidth, pic.ScaleHeight, pic.hDC, pic.hDC, 0, 0, pic.ScaleWidth, pic.ScaleHeight, FPersent * &H10000)


End Sub

'アクションゲーム用//////////////////////////////////////////////////////////////////////////////////////////////////アクションゲーム用

Public Sub Circle1(WAY As Long, A As Long, X As Long, Y As Long)  'WAY　進む距離 A 角度
                                                                   
X = X + Int(WAY * Cos((A * Pi / 180))) 'らじあんだそうです
Y = Y + Int(WAY * Sin((A * Pi / 180)))

End Sub

Public Sub Circle2(X As Long, Y As Long, Xc As Long, Yc As Long, R As Long, A As Long, LR As Long) 'R半径　Ａ　角度
                                                                     
X = Xc + Int(R * Cos((A * Pi / 180))) 'らじあんだそうです       'Xc Yc　中心＝Centerの座標を基準に見て
Y = Yc + LR * Int(R * Sin((A * Pi / 180)))                     '第一、第二因数にいれた変数の位置を角度A向きにR移動することができます
                                
End Sub

Public Sub AIM1(Xm As Long, Ym As Long, Xe As Long, Ye As Long, ByVal Z As Long) '  Xm Ym 自分の今の座標
                                                                           '　Xe Ye　敵の今の座標
Dim Xh As Long                                                             '  Ｚ 進む距離
Dim Yh As Long                             '以上により、当たるまで狙い続ける追尾弾をつくれます

Xh = Xe - Xm
Yh = Ye - Ym

If Xh <> 0 Or Yh <> 0 Then
    Xm = IIf(Xh = 0, Xm, Xm + Int((Xh * Z) / Sqr(Xh ^ 2 + Yh ^ 2)))
    Ym = IIf(Yh = 0, Ym, Ym + Int((Yh * Z) / Sqr(Xh ^ 2 + Yh ^ 2)))
End If

End Sub

Public Sub AIM2(X As Long, Y As Long, Xm As Long, Ym As Long, Xe As Long, Ye As Long, ByVal Z As Long) 'X Y 自分の今の座標
                                                                           'Xm Ym 自分の前の座標　Xe Ye　敵の前の座標
Dim Xh As Long                                                             '  Ｚ 進む距離
Dim Yh As Long     '以上により、自分の最初の座標と敵の最初の座標を放り込んだら、敵めがけて一直線に飛んでいく物体の移動をつくれます
                            
Xh = Xe - Xm 'xh=o yh=0 のとき怖いが、きっと、すぐに辺り判定があるからだいじょうぶだろう
Yh = Ye - Ym

If Xh <> 0 Or Yh <> 0 Then
    X = IIf(Xh = 0, X, X + Int((Xh * Z) / Sqr(Xh ^ 2 + Yh ^ 2)))
    Y = IIf(Yh = 0, Y, Y + Int((Yh * Z) / Sqr(Xh ^ 2 + Yh ^ 2)))
End If

End Sub

Public Function ForATARIHANTEI(MYLeft As Long, MYTop As Long, MYWidth As Long, MYHeight As Long, EnemyLeft As Long, EnemyTop As Long, EnemyWidth As Long, EnemyHeight As Long) As Boolean

ForATARIHANTEI = False '当たり判定が成立していればTrueを,してなければFalseを返します
If MYLeft < EnemyLeft + EnemyWidth And MYTop < EnemyTop + EnemyHeight _
And MYLeft + MYWidth > EnemyLeft And MYTop + MYHeight > EnemyTop Then ForATARIHANTEI = True

End Function

Public Function Between(X0, Y0, x1, y1) '２点間の距離関数　弾幕先輩より

Between = Sqr((X0 - x1) ^ 2 + (Y0 - y1) ^ 2)

End Function

Public Function ArcCos(X As Long, Y As Long, D As Double, EX As Long, EY As Long) As Double
'ある点からの角度を求める関数　　弾幕先輩より
'(EX,EY)から見た(X,Y)の角度
'逆の方が良かった気もしないでもない

'まず０度の場合
If X >= EX And Y = EY Then
    ArcCos = 0
'次に０度より大きく９０度以下の場合
ElseIf X >= EX And Y > EY Then
    ArcCos = Atn(-Abs((EX - X) / D) / Sqr(-((EX - X) / D) ^ 2 + 1)) + Pi / 2
'次に９０度より大きく１８０度より小さい場合
ElseIf X < EX And Y > EY Then
    ArcCos = Abs(Pi / 2 - Atn(-Abs((EX - X) / D) / Sqr(-((EX - X) / D) ^ 2 + 1)) + Pi / 2) - Pi / 2
'次に１８０度の場合
ElseIf X < EX And Y = EY Then
    ArcCos = 180
'次に１８０度より大きく２７０度より小さい場合
ElseIf X < EX And Y < EY Then
    ArcCos = Atn(-Abs((EX - X) / D) / Sqr(-((EX - X) / D) ^ 2 + 1)) + 3 * Pi / 2
'最後に、２７０度以上、３６０度未満の場合
ElseIf X >= EX And Y < EY Then
    ArcCos = Abs(Pi / 2 - Atn(-Abs((EX - X) / D) / Sqr(-((EX - X) / D) ^ 2 + 1)) + Pi / 2) + Pi / 2
End If

End Function

Public Sub PlayMovie(FileName As String, PicBox As PictureBox) '動画再生系////////////////////////////////////////////////////////動画再生系

PicBox.Appearance = 0      'ムービーをピクチャーボックス上に再生できます
PicBox.BorderStyle = 0     'いつものピクチャーボックスが映像用ボックスに早変わり！
PicBox.Visible = True      '念のためピクチャーボックスは、画像を入れないでください
                            '映像の大きさに合わせてピクチャーボックスの大きさを変えれませんので、そこだけは初期設定してください。
                            '更に、再生、終了だけなら大抵の拡張子がいけますが、
                            '確認済みではmpegでしか後述のMovieLengthとMoviePositionを使えません。のでMpeg推奨
                            'MovieLengthとMoviePositionを使えば、動画の長さ、現在の再生位置が分かるので、うまく利用すれば、
                            '再生終了をできます。
Call mciSendString("open """ & FileName & """", "", 0, 0)
Call mciSendString("window """ & FileName & """" & " handle " & PicBox.hWnd, "", 0, 0)
Call mciSendString("play """ & FileName & """", "", 0, 0)
    
End Sub

Public Sub CloseMovie(FileName As String, PicBox As PictureBox)

Call mciSendString("close """ & FileName & """", "", 0, 0) '再生しているムービーを終了します
PicBox.Visible = False

End Sub

Public Function MovieLength(FileName As String) As Double

    Dim RetBuffer As String * 20
    Dim MCICommandString As String

    MCICommandString = "status """ & FileName & """ length"
    Call mciSendString("open """ & FileName & """", "", 0, 0)
    Call mciSendString(MCICommandString, RetBuffer, Len(RetBuffer), 0)
    Call mciSendString("close """ & FileName & """", "", 0, 0)

    'GetLength = Val(RetBuffer) / 1000

End Function

Public Function MoviePosition(FileName As String) As Double

    Dim RetBuffer As String * 20
    Dim MCICommandString As String

    MCICommandString = "status """ & FileName & """ position"
    Call mciSendString(MCICommandString, RetBuffer, Len(RetBuffer), 0)

    'GetPosition = Val(RetBuffer) / 1000

End Function

'スコア各種////////////////////////////////////////////////////////////////////////////////////////////////////////スコア各種

Public Sub ScoreAllDelete(Optional first As Boolean = False)
    Dim i As Long
If first = True Then
 
    For i = 1 To UBound(Data)
        Data(i).Key = "": Data(i).Name = "": Data(i).Str = ""
    Next i
    Call FirstSet
    
Else

Call KeyAllDelete
Call NameAllDelete
Call StrAllDelete

End If
End Sub
Public Sub KeyAllDelete()

Dim i As Long
For i = 0 To UBound(Data)
    Data(i).Key = ""
Next i
End Sub
Public Sub NameAllDelete()

Dim i As Long
For i = 0 To UBound(Data)
    Data(i).Name = ""
Next i
End Sub
Public Sub StrAllDelete()
Dim i As Long
For i = 0 To UBound(Data)
    Data(i).Str = ""
Next i
End Sub

Public Function HighScoreCheckK(NowScore As Long, kind As RankKind, ParamArray RankKey() As Variant) As Boolean

    Dim i As Long
    Dim j As Long
    Dim Memo As Long
    If kind = smaller Then
        j = 0
    Else
        j = 1000000000
    End If
    
    For i = 0 To UBound(RankKey)
        If kind = smaller Then
            If CLng(ReadText(CStr(RankKey(i)))) > j Then
                j = CLng(ReadText(CStr(RankKey(i))))
                Memo = i
            End If
        Else
            If CLng(ReadText(CStr(RankKey(i)))) < j Then
                j = CLng(ReadText(CStr(RankKey(i))))
                Memo = i
            End If
        End If
    Next i
    
    If kind = smaller Then
    
        If CLng(ReadText(CStr(RankKey(Memo)))) > NowScore Then HighScoreCheckK = True
    
    Else
    
        If CLng(ReadText(CStr(RankKey(Memo)))) < NowScore Then HighScoreCheckK = True
        
    End If


End Function
Public Function INRankK(NowScore As Long, kind As RankKind, Name As String, ParamArray RankKey() As Variant)
    Dim i As Long
    Dim j As Long
    Dim Memo As Long
    
    If kind = smaller Then
        j = 0
    Else
        j = 1000000000
    End If
    For i = 0 To UBound(RankKey)
        If kind = smaller Then
            If CLng(ReadText(CStr(RankKey(i)))) > j Then
                j = CLng(ReadText(CStr(RankKey(i))))
                Memo = i
            End If
        Else
            If CLng(ReadText(CStr(RankKey(i)))) < j Then
                j = CLng(ReadText(CStr(RankKey(i))))
                Memo = i
            End If
        End If
    Next i
    Call WriteText(CStr(NowScore), CStr(RankKey(Memo)))
    Call WriteText(Name, CStr(RankKey(Memo)), True)

End Function

Public Function RankK(num As Long, Name As Boolean, kind As RankKind, ParamArray RankKey() As Variant) As String
Dim i As Long
Dim j As Long
Dim k As Long
Dim DeleteNum() As Long
Dim RankNum As Long
Dim CheckScore() As Long
Dim Memo As Long

ReDim CheckScore(0) 'ランキングで比較されるスコアを取得する。
ReDim DeleteNum(0) '順位取得のため、自分以上の順位を消すために利用。
DeleteNum(0) = -1

On Error Resume Next

For j = 1 To num '順位が出るまで続く
    ReDim Preserve DeleteNum(UBound(DeleteNum) + 1) '消す番号が一つ増える
    DeleteNum(UBound(DeleteNum)) = -1
    If kind = bigger Then '大きい方をとるとき。
        RankNum = 0
        Else: RankNum = 1000000000 'それ以外
    End If
    
    For i = 0 To UBound(RankKey) '一つ一つ
    
        For k = 0 To UBound(DeleteNum)
            If i = DeleteNum(k) Then GoTo Nextloop '順位が上のデータなら次へ。
        Next k
    
            
            ReDim Preserve CheckScore(i)
            CheckScore(i) = CLng(ReadText(CStr(RankKey(i)))) '実際に調べるためのスコアを取得
        
        If kind = bigger Then '大きい方をとるとき。
            If Max(CheckScore(i), RankNum) = CheckScore(i) Then
                RankNum = CheckScore(i) '段段大きい方をとっていく。
                Memo = i
            End If
        Else
            If Min(CheckScore(i), RankNum) = CheckScore(i) Then '小さいほうをとるとき
                RankNum = CheckScore(i) '段段小さいほうを取っていく。
                Memo = i
            End If
        End If
Nextloop:
    Next i
    
    DeleteNum(UBound(DeleteNum)) = Memo '現在の順位のやつ。
    
    
Next j

If Name = False Then 'Forを抜けたなら、そのままMemoを使用。
    RankK = ReadText(CStr(RankKey(Memo)))
Else
    RankK = ReadText(CStr(RankKey(Memo)), True)
End If


End Function
Public Function HighScoreCheckN(NowScore As Long, kind As RankKind, Optional FirstNum As Long = 1, Optional LastNum As Long = 10, Optional StepNum As Long = 1) As Boolean
        
    Dim i As Long
    Dim j As Long
    Dim Memo As Long
    If kind = smaller Then
        j = 0
    Else
        j = RankN(1, bigger, FirstNum, LastNum, StepNum)
    End If
    
    For i = FirstNum To LastNum Step StepNum
        If kind = smaller Then
            If CLng(Data(i).Str) > j Then
                j = CLng(Data(i).Str)
                Memo = i
            End If
        Else
            If CLng(Data(i).Str) < j Then
                j = CLng(Data(i).Str)
                Memo = i
            End If
        End If
    Next i
    
    If kind = smaller Then
    
        If CLng(Data(Memo).Str) > NowScore Then HighScoreCheckN = True
    
    Else
    
        If CLng(Data(Memo).Str) < NowScore Then HighScoreCheckN = True
        
    End If

End Function
Public Sub InRankN(NowScore As Long, kind As RankKind, Name As String, Optional FirstNum As Long = 1, Optional LastNum As Long = 10, Optional StepNum As Long = 1)
    Dim i As Long
    Dim j As Long
    Dim Memo As Long
        
    If kind = smaller Then
        j = 0
    Else
        j = RankN(1, bigger, FirstNum, LastNum, StepNum)
    End If
    For i = FirstNum To LastNum Step StepNum
        If kind = smaller Then
            If CLng(Data(i).Str) > j Then
                j = CLng(Data(i).Str)
                Memo = i
            End If
        Else
            If CLng(Data(i).Str) < j Then
                j = CLng(Data(i).Str)
                Memo = i
            End If
        End If
    Next i
    If Memo = 0 Then
        Exit Sub
    Else
        Data(Memo).Str = NowScore
        Data(Memo).Name = Name
    End If
End Sub
Public Function RankN(num As Long, kind As RankKind, Optional FirstNum As Long = 1, Optional LastNum As Long = 10, Optional StepNum As Long = 1, Optional Name As Boolean = False) As String
'データ配列番号を利用したランキング
'第一引数:必要とする順位、第二引数:ランクを求める配列の最初の数、第三引数:ランクを求める配列の最後の数、第四引数:一つ飛ばしや、二つ飛ばしの時のため、第五引数:スコアを返すか名前を返すか
'Data(i).strが数字じゃないとエラーになるから注意。
Dim i As Long
Dim j As Long
Dim k As Long
Dim DeleteNum() As Long
Dim RankNum As Long
Dim CheckScore() As Long
Dim Memo As Long

On Error Resume Next

ReDim CheckScore(0) 'ランキングで比較されるスコアを取得する。
ReDim DeleteNum(0) '順位取得のため、自分以上の順位を消すために利用。

For j = 1 To num '順位が出るまで続く
    ReDim Preserve DeleteNum(UBound(DeleteNum) + 1) '消す番号が一つ増える
    
    If kind = bigger Then '大きい方をとるとき。
        RankNum = 0
        Else: RankNum = RankN(1, bigger, FirstNum, LastNum, StepNum) + 1 'それ以外
    End If
    
    For i = FirstNum To LastNum Step StepNum '一つ一つ
    
        For k = 0 To UBound(DeleteNum)
            If i = DeleteNum(k) Then GoTo Nextloop '順位が上のデータなら次へ。
        Next k
    
            
            ReDim Preserve CheckScore(i)
            CheckScore(i) = CLng(ReadText(CStr(i))) '実際に調べるためのスコアを取得
        
        If kind = bigger Then '大きい方をとるとき。
            If Max(CheckScore(i), RankNum) = CheckScore(i) Then
                RankNum = CheckScore(i) '段段大きい方をとっていく。
                Memo = i
            End If
        Else
            If Min(CheckScore(i), RankNum) = CheckScore(i) Then '小さいほうをとるとき
                RankNum = CheckScore(i) '段段小さいほうを取っていく。
                Memo = i
            End If
        End If
Nextloop:
    Next i
    
    DeleteNum(UBound(DeleteNum)) = Memo '現在の順位のやつ。
    
    
Next j

If Name = False Then 'Forを抜けたなら、そのままMemoを使用。
    RankN = ReadText(CStr(Memo))
Else
    RankN = ReadText(CStr(Memo), True)
End If

End Function

