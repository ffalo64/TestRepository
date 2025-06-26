Attribute VB_Name = "BasicModule"

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_CHARSTREAM = 4          '  Character-stream, PLP
Public Const DT_DISPFILE = 6            '  Display-file
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_INTERNAL = &H1000
Public Const DT_LEFT = &H0
Public Const DT_METAFILE = 5            '  Metafile, VDM
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_PLOTTER = 0             '  Vector plotter
Public Const DT_RASCAMERA = 3           '  Raster camera
Public Const DT_RASDISPLAY = 1          '  Raster display
Public Const DT_RASPRINTER = 2          '  Raster printer
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10
Public Const Entrance = 0
Public Const Dungeon = 1
Public Const HowtoPlay = 2
Public Const Options = 3
Public Const GameOver = 4
Public Const GameClear = 5
Public Const Museum = 6
Public Const PlayerConst = 10 ^ 9

Public Type RECT

    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    
End Type

Public Type Status

    Left As Long
    Top As Long
    Oleft As Long
    Otop As Long
    Width As Long
    Height As Long
    Alive As Boolean
    Hp As Long
    MaxHp As Long
    Exp As Long         '経験値
    ATK As Long
    Ability As Long
    DEF As Long
    Level As Long
    Condition As Long
    Direction As Long
    Name As String
    Explanation As String  '画面に表示される説明

End Type

'変数の定義

Public Player As Status
Public HpBar As Status
Public Music() As Status
Public SpecBar As RECT
Public StatusBar As RECT
Public MessageBar As RECT
Public GameMode As Long
Public Floor As Long
Public Turn As Long
Public SumDamage As Long
Public Lib As ImscLib12

Public Sub Main() '一番最初に通る関数。
'右のPrpjectで右クリック
'→Project1のプロパティをクリック
'→スタートアップの設定をSub Mainに変更
'→モジュールの中にPublic Sub Mainと書く

'まずライブラリを使うためには下の一行を書いてやる
'これはコード中一回しか通らんところ　例)Sub Mainとか　に書くのが無難
Set Lib = New ImscLib12

Form1.Show 'ゲーム開始時のFormをShow

Call Lib.SoundSet(0, App.Path & "\Sounds" & "\n31.mp3") 'メニュー
Call Lib.SoundSet(1, App.Path & "\Sounds" & "\n11.mp3") 'ゲームクリア
Call Lib.SoundSet(2, App.Path & "\Sounds" & "\c26.mp3") 'ゲームオーバー
Call Lib.SoundSet(3, App.Path & "\Sounds" & "\c1.mp3")  'ダンジョン

End Sub

Public Sub Movement()

    Select Case GameMode
    
        Case Entrance
        
            Player.Direction = 1

        Case Dungeon
                                  
            If Player.Direction > 1 Then
            
                Dim i As Long
                        
                With Player
                              
                    If (.Direction Mod 2 = 0) And (.Top > 0) Then .Top = .Top - .Height
                    If (.Direction Mod 3 = 0) And (.Top < LandNumber * Sc - .Height) Then .Top = .Top + .Height
                    If (.Direction Mod 5 = 0) And (.Left < LandNumber * Sc - .Width) Then .Left = .Left + .Width
                    If (.Direction Mod 7 = 0) And (.Left > 0) Then .Left = .Left - .Width
                    
                    Call PositionCheck(PlayerConst)
                    
                    If .Condition <> Slow Then
                    
                        .Direction = 1
                        Turn = Turn - 1
                        
                    ElseIf (.Condition = Slow) And (Turn Mod 2 = 0) Then
                    
                        .Direction = 11
                        Turn = Turn - 1
                        
                    Else
                    
                        .Direction = 1
                        Turn = Turn - 1
                        
                    End If
        
                End With
                
                For i = 0 To MonsterNumber - 1
        
                    With Monster(i)
                    
                        If (.Alive = True) And (.Ability <> Slow) And (Player.Condition <> Stealth) Then
                                                                           
                            If .Top > Player.Top Then
                            
                                .Top = .Top - .Height
                                .Direction = .Direction * 2
                                
                            End If
                                                             
                            If .Top < Player.Top Then
                            
                                .Top = .Top + .Height
                                .Direction = .Direction * 3
          
                            End If
                                
                            If .Left < Player.Left Then
                            
                                .Left = .Left + .Width
                                .Direction = .Direction * 5
                                
                            End If
                                    
                            If .Left > Player.Left Then
                            
                                .Left = .Left - .Width
                                .Direction = .Direction * 7
                                
                            End If
                            
                            Call PositionCheck(i)
                            .Direction = 1
                            
                        ElseIf (.Alive = True) And (.Ability = Slow) And (Turn Mod 2 = 0) And (Player.Condition <> Stealth) Then
                        
                            If .Top > Player.Top Then
                            
                                .Top = .Top - .Height
                                .Direction = .Direction * 2
                                
                            End If
                                                             
                            If .Top < Player.Top Then
                            
                                .Top = .Top + .Height
                                .Direction = .Direction * 3
          
                            End If
                                
                            If .Left < Player.Left Then
                            
                                .Left = .Left + .Width
                                .Direction = .Direction * 5
                                
                            End If
                                    
                            If .Left > Player.Left Then
                            
                                .Left = .Left - .Width
                                .Direction = .Direction * 7
                                
                            End If
                            
                            Call PositionCheck(i)
                            .Direction = 1

                        End If
                            
                    End With
                
                Next i
                
                If SumDamage > 0 Then
                
                    Words(3) = "プレイヤーは" & CStr(SumDamage) & "ダメージを受けた"
                    SumDamage = 0
                
                End If

            End If
            
        Case HowtoPlay
        
            With Player
        
                If (.Direction Mod 5 = 0) Then
                
                    .Condition = (.Condition + 1) Mod 6
                
                ElseIf (.Direction Mod 7 = 0) And (.Condition = 0) Then
                
                    .Condition = 5
                    
                ElseIf (.Direction Mod 7 = 0) And (.Condition > 0) Then
                
                   .Condition = .Condition - 1
                
                End If
                
                .Direction = 1
                
                Select Case .Condition
                
                    Case 0
                    
                        For i = 0 To UBound(Landsquare)
            
                            Landsquare(i, 0).Alive = False
            
                        Next i
                
                    Case 1
                    
                        For i = 0 To UBound(Landsquare)
                        
                            With Landsquare(i, 0)
                            
                                .Alive = True

                                Select Case i
                                
                                    Case 0 To 4
                                    
                                        .Left = Form1.ScaleWidth / 2 + Sc * 2 * i
                                        .Top = Player.Top
                                        .Condition = Mine
                                        .Ability = i
                                
                                    Case 5
                                
                                        .Left = Form1.ScaleWidth / 2
                                        .Top = Player.Top + Form1.FontSize * 4
                                        .Condition = Stair
                                        
                                    Case 6 To 15
                                    
                                        .Left = (Form1.ScaleWidth / 2) + Sc * 2 * (i - 10)
                                        .Top = Player.Top + Form1.FontSize * 14
                                        .Condition = Wall
                                        .Ability = i - 6
            
                                    Case 16 To 25
                                    
                                        .Left = (Form1.ScaleWidth / 2) + Sc * 2 * (i - 20)
                                        .Top = Player.Top + Form1.FontSize * 16
                                        .Condition = Room
                                        .Ability = i - 16
                                        
                                End Select
                            
                            End With
            
                        Next i
                        
                    Case 2 To 5
                    
                        For i = 0 To UBound(Landsquare)
            
                            With Landsquare(i, 0)
                            
                                Select Case i
                                
                                    Case 0 To 4
                            
                                        .Left = (Form1.ScaleWidth / 2) + Sc * 2 * (i - 2)
                                        .Top = Player.Top - Form1.FontSize * 2
                                        .Condition = i
                                        .Alive = True
                                    
                                    Case 5 To UBound(Landsquare)
                                    
                                        .Alive = False
                                
                                End Select
                            
                            End With
            
                        Next i
                        
                End Select
            
            End With
            
        Case Options
        
            With Player
        
                If (.Direction Mod 5 = 0) Then
                
                    .Ability = (.Ability + 1) Mod 5
                
                ElseIf (.Direction Mod 7 = 0) And (Player.Ability = 0) Then
                
                    .Ability = 4
                    
                ElseIf (.Direction Mod 7 = 0) And (Player.Ability > 0) Then
                
                    .Ability = .Ability - 1
                
                End If
                
                .Direction = 1
            
            End With
            
        Case Museum
        
            With Player
        
                If (.Direction Mod 5 = 0) Then
                
                    .Condition = (.Condition + 1) Mod 6
                
                ElseIf (.Direction Mod 7 = 0) And (.Condition = 0) Then
                
                    .Condition = 5
                    
                ElseIf (.Direction Mod 7 = 0) And (.Condition > 0) Then
                
                   .Condition = .Condition - 1
                
                End If
                
                .Direction = 1
            
            End With
    
    End Select

End Sub

Public Sub StatusCheck(Levelup As Boolean)

Dim i As Long
Dim j As Long
        
    With Player
    
        If (.Exp >= .Level ^ 3) And (.Level < 999) And (.Alive = True) Then
        
            Dim DHp As Long
    
            DHp = Int(5 + Rnd() * 3)
            .Level = .Level + 1
            .MaxHp = .MaxHp + DHp
            .Hp = .Hp + DHp
            .ATK = (.ATK * 1.1) + 1
            .DEF = .DEF + 1
                            
        End If
        
        If .ATK > 10 ^ 9 - 1 Then .ATK = 10 ^ 9 - 1
        If .DEF > 10 ^ 9 - 1 Then .DEF = 10 ^ 9 - 1
        If .DEF = 0 Then .DEF = 1
        If Turn > 10 ^ 4 - 1 Then Turn = 10 ^ 4 - 1
        If .Exp > 10 ^ 9 - 1 Then .Exp = 10 ^ 9 - 1
        If .MaxHp > 10 ^ 7 - 1 Then .MaxHp = 10 ^ 7 - 1
        
        If .Hp > .MaxHp Then
        
            .Hp = .MaxHp
        
        ElseIf .Hp > .MaxHp / 4 Then
 
            Form1.ForeColor = vbWhite
        
        ElseIf (.Hp > 0) And (.Hp <= .MaxHp / 4) Then
        
            Form1.ForeColor = RGB(255, 127, 39)
            
        End If
        
        If (.Hp <= 0) Or (Turn <= 0) Then
        
            If AbilityHp(5) = 0 Then
                
                .Hp = 0
                .Direction = 1
                .Alive = False
                GameMode = GameOver
                Call LandSet
                Call WordSet
                
            ElseIf AbilityHp(5) > 0 Then
            
                .Alive = False
                Call AbilityEffect(vbKeyD)
            
            End If
                
        End If

    End With
    
    For i = 0 To UBound(Monster)

        With Monster(i)
        
            If (.Level < 999) And (Levelup = True) Then
        
                .Level = .Level + 1
                .MaxHp = (.MaxHp * 1.1) + 1
                .Hp = .MaxHp
                .ATK = (.ATK * 1.2) + 1
                .DEF = .DEF + 1
                .Exp = .Exp + 10
                                    
                If i = 0 Then Words(0) = "ヘキサスライム達はLevel" & CStr(.Level) & "になった。"
    
            End If
            
            If .ATK > 10 ^ 9 - 1 Then .ATK = 10 ^ 9 - 1
            If .DEF > 10 ^ 9 - 1 Then .DEF = 10 ^ 9 - 1
            If .DEF = 0 Then .DEF = 1
            If .Exp > 10 ^ 4 - 1 Then .Exp = 10 ^ 4 - 1
            If .Hp > .MaxHp Then .Hp = .MaxHp
            If .MaxHp > 10 ^ 6 - 1 Then .MaxHp = 10 ^ 6 - 1
    
        End With
        
    Next i
    
    For i = 0 To LandNumber - 1
    For j = 0 To LandNumber - 1
    
        With Landsquare(i, j)
    
            If .DEF > 10 ^ 9 - 1 Then .DEF = 10 ^ 9 - 1
            If .DEF = 0 Then .DEF = 1
            If .Hp > .MaxHp Then .Hp = .MaxHp
            If .MaxHp > 10 ^ 6 - 1 Then .MaxHp = 10 ^ 6 - 1
            
        End With
        
    Next j
    Next i

End Sub

Public Sub FloorSet()

    Dim i As Long
    Dim j As Long
    
    Floor = Floor + 1
    RandomPosition.Hp = LandNumber ^ 2
        
    Select Case Floor
    
        Case 1
    
            For i = 0 To UBound(Words)
    
                Words(i) = ""
                    
            Next i
      
            With SpecBar
            
                .Left = LandNumber * Sc
                .Right = Form1.ScaleWidth
                .Top = 0
                .Bottom = LandNumber * Sc
        
            End With
            
            With StatusBar
            
                .Left = 0
                .Right = Form1.ScaleWidth
                .Top = LandNumber * Sc
                .Bottom = Form1.ScaleHeight
            
            End With
            
            With MessageBar
            
                .Left = 0
                .Right = Form1.ScaleWidth
                .Top = StatusBar.Top + Form1.FontSize * 2
                .Bottom = Form1.ScaleHeight
            
            End With
            
            With Player
            
                .Hp = 20
                .MaxHp = 20
                .Level = 1
                .Exp = 0
                .Height = Sc
                .Width = Sc
                .Left = 0
                .Top = 0
                .ATK = 10
                .DEF = 3
                .Direction = 1
                .Condition = 0
            
            End With
            
            Turn = 200
            MonsterNumber = 10
            
        Case 20, 40, 60, 80, 100
        
            MonsterNumber = MonsterNumber + 10
            
        Case 200, 300, 400, 500, 600, 700, 800
        
            MonsterNumber = MonsterNumber + 100
            
        Case 900
        
            MonsterNumber = 1000
            
        Case 1000
        
            Player.Direction = 1
            Player.Alive = False
            GameMode = GameClear
            Call LandSet
            Call WordSet
                    
    End Select
    
    Call Form1.MapSet
    
    Do
        
        With Player
    
            .Left = Int(Rnd() * LandNumber) * Sc
            .Top = Int(Rnd() * LandNumber) * Sc
            .Oleft = .Left
            .Otop = .Top
    
            If Landsquare(.Left / Sc, .Top / Sc).Condition = Room Then
            
                .Alive = True
                Landsquare(.Left / Sc, .Top / Sc).Alive = False
                RandomPosition.Hp = RandomPosition.Hp - 1
                .Condition = Noability
                Exit Do
            
            End If
    
        End With
            
    Loop
    
    Do
        
        With RandomPosition
    
            .Left = Int(Rnd() * LandNumber) * Sc
            .Top = Int(Rnd() * LandNumber) * Sc
    
            If (Landsquare(.Left / Sc, .Top / Sc).Alive = True) And (Landsquare(.Left / Sc, .Top / Sc).Condition = Room) Then
            
                Landsquare(.Left / Sc, .Top / Sc).Condition = Stair
                .Hp = .Hp - 1
                .Condition = 0
                Exit Do
            
            End If
    
        End With
            
    Loop
    
    Do
        
        With RandomPosition
    
            .Oleft = Int(Rnd() * LandNumber) * Sc
            .Otop = Int(Rnd() * LandNumber) * Sc
            .Condition = .Condition + 1
    
            If (Landsquare(.Oleft / Sc, .Otop / Sc).Alive = True) And (Landsquare(.Oleft / Sc, .Otop / Sc).Condition = Room) Then
            
                Landsquare(.Oleft / Sc, .Otop / Sc).Condition = PurpleBox
                .Hp = .Hp - 1
                .Condition = 0
                Exit Do
                
            ElseIf .Condition > LandNumber ^ 2 Then
                    
                .Condition = 0
                Exit Do
            
            End If
    
        End With
            
    Loop
    
    For i = 0 To UBound(Monster)
    
        With Monster(i)
                
            Do
                             
                .Left = Int(Rnd() * LandNumber) * Sc
                .Top = Int(Rnd() * LandNumber) * Sc
                .Oleft = .Left
                .Otop = .Top
                
                If i > MonsterNumber - 1 Then
                
                    .Alive = False
                    .Condition = 0
                    Exit Do
                
                ElseIf (Landsquare(.Left / Sc, .Top / Sc).Alive = True) And (Landsquare(.Left / Sc, .Top / Sc).Condition = Room) Then
        
                    .Alive = True
                    .Hp = .MaxHp
                    Landsquare(.Left / Sc, .Top / Sc).Condition = Enemy
                    RandomPosition.Hp = RandomPosition.Hp - 1
                    .Condition = 0
                    
                    Exit Do
                    
                ElseIf RandomPosition.Hp = 0 Then
                    
                    .Alive = False
                    .Condition = 0
                    Exit Do
        
                End If
                
            Loop
        
        End With
        
    Next i
    
    If Floor >= 2 Then Call StatusCheck(True)
    
End Sub

Public Sub PositionCheck(Character As Long)

Dim i As Long
Dim j As Long

    Select Case Character
    
        Case PlayerConst
    
            With Player
            
                If Landsquare(.Left / Sc, .Top / Sc).Condition = Stair Then

                    If AbilityHp(6) = 0 Then
                    
                        Call FloorSet
                        
                    Else
                    
                        Call AbilityEffect(vbKeyS)
                    
                    End If
                    
                ElseIf Landsquare(.Left / Sc, .Top / Sc).Condition <= PurpleBox Then
                
                    Call BattleCheck(PlayerConst, .Left / Sc, .Top / Sc)

                    '移動を打ち消す処理
                                    
                    .Left = .Oleft
                    .Top = .Otop

                ElseIf Landsquare(.Left / Sc, .Top / Sc).Condition = Wall Then
                    
                    If (.Condition = WallBreak) And (.Left <> 0) And (.Top <> 0) And (.Left <> (LandNumber - 1) * Sc) And (.Top <> (LandNumber - 1) * Sc) Then
                    
                        Call BattleCheck(PlayerConst, .Left / Sc, .Top / Sc)
                    
                    End If

                    .Left = .Oleft
                    .Top = .Otop
                    
                End If
                
                For i = 0 To UBound(Monster)
               
                    If (.Left = Monster(i).Left) And (.Top = Monster(i).Top) And (Monster(i).Alive = True) Then
                                                                             
                        Call BattleCheck(PlayerConst, i, LandNumber)
                        
                        .Left = .Oleft
                        .Top = .Otop
         
                    End If
                
                Next i
                
                .Oleft = .Left
                .Otop = .Top

            End With
            
        Case 0 To MonsterNumber - 1
        
            With Monster(Character)
            
                If .Alive = True Then
            
                    If (.Left = Player.Left) And (.Top = Player.Top) And (Player.Alive = True) Then
                        
                        .Left = .Oleft
                        .Top = .Otop
                        
                        Call BattleCheck(Character, PlayerConst, LandNumber)
                        
                    ElseIf Landsquare(.Left / Sc, .Top / Sc).Condition <> Room Then
                        
                        If (Landsquare(.Left / Sc, .Top / Sc).Condition = Wall) And (.Ability = WallBreak) And (.Left <> 0) And (.Top <> 0) And (.Left <> (LandNumber - 1) * Sc) And (.Top <> (LandNumber - 1) * Sc) Then
                        
                            Call BattleCheck(Character, .Left / Sc, .Top / Sc)
                            
                        ElseIf (Landsquare(.Left / Sc, .Top / Sc).Condition <= PurpleBox) And (.Ability = Boxattack) Then
                        
                            Call BattleCheck(Character, .Left / Sc, .Top / Sc)

                        End If

                        Select Case .Direction
                        
                            Case 2
                            
                                If Landsquare((.Oleft / Sc) - 1, (.Otop / Sc) - 1).Condition = Room Then
                                
                                    .Left = .Oleft - Sc
                                    .Top = .Otop - Sc
                                    
                                ElseIf Landsquare((.Oleft / Sc) + 1, (.Otop / Sc) - 1).Condition = Room Then
                                
                                    .Left = .Oleft + Sc
                                    .Top = .Otop - Sc
                                    
                                Else
                                
                                    .Left = .Oleft
                                    .Top = .Otop
                            
                                End If
                            
                            Case 3
                            
                                If Landsquare((.Oleft / Sc) + 1, (.Otop / Sc) + 1).Condition = Room Then
                                
                                    .Left = .Oleft + Sc
                                    .Top = .Otop + Sc
                                    
                                ElseIf Landsquare((.Oleft / Sc) - 1, (.Otop / Sc) + 1).Condition = Room Then
                                
                                    .Left = .Oleft - Sc
                                    .Top = .Otop + Sc
                                    
                                Else
                                
                                    .Left = .Oleft
                                    .Top = .Otop
                            
                                End If
                            
                            Case 5
                            
                                If Landsquare((.Oleft / Sc) + 1, (.Otop / Sc) - 1).Condition = Room Then
                                
                                    .Left = .Oleft + Sc
                                    .Top = .Otop - Sc
                                    
                                ElseIf Landsquare((.Oleft / Sc) + 1, (.Otop / Sc) + 1).Condition = Room Then
                                
                                    .Left = .Oleft + Sc
                                    .Top = .Otop + Sc
                                    
                                Else
                                
                                    .Left = .Oleft
                                    .Top = .Otop
                            
                                End If
                            
                            Case 7
                            
                                If Landsquare((.Oleft / Sc) - 1, (.Otop / Sc) + 1).Condition = Room Then
                                
                                    .Left = .Oleft - Sc
                                    .Top = .Otop + Sc
                                    
                                ElseIf Landsquare((.Oleft / Sc) - 1, (.Otop / Sc) - 1).Condition = Room Then
                                
                                    .Left = .Oleft - Sc
                                    .Top = .Otop - Sc
                                    
                                Else
                                
                                    .Left = .Oleft
                                    .Top = .Otop
                            
                                End If
                        
                            Case 10
                            
                                If Landsquare(.Oleft / Sc, (.Otop / Sc) - 1).Condition = Room Then
                                    
                                    .Left = .Oleft
                                    .Top = .Otop - Sc
                                    
                                ElseIf Landsquare((.Oleft / Sc) + 1, .Otop / Sc).Condition = Room Then
                                
                                    .Left = .Oleft + Sc
                                    .Top = .Otop
                                    
                                Else
                                
                                    .Left = .Oleft
                                    .Top = .Otop
                            
                                End If
                            
                            Case 15
                            
                                If Landsquare((.Oleft / Sc) + 1, .Otop / Sc).Condition = Room Then
                                
                                    .Left = .Oleft + Sc
                                    .Top = .Otop
                                    
                                ElseIf Landsquare(.Oleft / Sc, (.Otop / Sc) + 1).Condition = Room Then
                                    
                                    .Left = .Oleft
                                    .Top = .Otop + Sc
                                    
                                Else
                                
                                    .Left = .Oleft
                                    .Top = .Otop
                            
                                End If
                            
                            Case 21
                            
                                If Landsquare(.Oleft / Sc, (.Otop / Sc) + 1).Condition = Room Then
                                    
                                    .Left = .Oleft
                                    .Top = .Otop + Sc
                                    
                                ElseIf Landsquare((.Oleft / Sc) - 1, .Otop / Sc).Condition = Room Then
                                
                                    .Left = .Oleft - Sc
                                    .Top = .Otop
                                    
                                Else
                                
                                    .Left = .Oleft
                                    .Top = .Otop
                            
                                End If
                            
                            Case 14
                            
                                If Landsquare((.Oleft / Sc) - 1, .Otop / Sc).Condition = Room Then
                                
                                    .Left = .Oleft - Sc
                                    .Top = .Otop
                                    
                                ElseIf Landsquare(.Oleft / Sc, (.Otop / Sc) - 1).Condition = Room Then
                                
                                    .Left = .Oleft
                                    .Top = .Otop - Sc
                                    
                                Else
                                
                                    .Left = .Oleft
                                    .Top = .Otop
                            
                                End If
                        
                        End Select
                           
                    End If
                    
                    Landsquare(.Oleft / Sc, .Otop / Sc).Condition = Room
                    Landsquare(.Left / Sc, .Top / Sc).Condition = Enemy
                    .Oleft = .Left
                    .Otop = .Top
                
                End If
                            
            End With
            
    End Select
    
End Sub

Public Sub BattleCheck(Attacker As Long, i As Long, j As Long)

Dim Damage As Long

    Select Case Attacker
    
        Case PlayerConst

            With Player
                
                If j < LandNumber Then
                
                    With Landsquare(i, j)
                
                        Damage = Int(Player.ATK * (Rnd() * 0.2 + 0.9) / .DEF) + 1
                        
                        If Damage > 10 ^ 7 - 1 Then Damage = 10 ^ 7 - 1

                        .Hp = .Hp - Damage
                        Words(2) = .Name & "は" & CStr(Damage) & "ダメージを受けた"
                        
                        If .Hp <= 0 Then
                        
                            .Hp = 0
                            .Ability = Int(Rnd() * 15)
                            If Player.Condition <> TurnConst Then Turn = Turn + 10
    
                            Call BoxEffect(i, j)
                        
                        End If
                    
                    End With
                
                ElseIf j >= LandNumber Then
                
                    With Monster(i)
                    
                        If .Alive = True Then
                
                            Damage = Int(Player.ATK * (Rnd() * 0.2 + 0.9) / .DEF) + 1
                            
                            If Damage > 10 ^ 7 - 1 Then Damage = 10 ^ 7 - 1

                            .Hp = .Hp - Damage
                            Words(2) = .Name & "は" & CStr(Damage) & "ダメージを受けた"
                            
                            If .Hp <= 0 Then
                            
                                .Hp = 0
                                .Alive = False
                                Landsquare(.Left / Sc, .Top / Sc).Condition = Room
                                Player.Exp = Player.Exp + .Exp
                            
                            End If
                        
                        End If
                    
                    End With
                
                End If
        
            End With
            
        Case 0 To UBound(Monster)
        
            If j < LandNumber Then
            
                With Landsquare(i, j)
                
                    If .Condition <> Stair Then
            
                        Damage = Int(Monster(Attacker).ATK * (Rnd() * 0.2 + 0.9) / .DEF) + 1
                        
                        If Damage > 10 ^ 7 - 1 Then Damage = 10 ^ 7 - 1

                        .Hp = .Hp - Damage
                        Words(2) = .Name & "は" & CStr(Damage) & "ダメージを受けた"
                        
                        If .Hp <= 0 Then
                        
                            .Hp = 0
                            .Condition = Room
                        
                        End If
                    
                    End If
                
                End With
        
            ElseIf j >= LandNumber Then
        
                With Monster(Attacker)
        
                    Damage = Int(.ATK * (Rnd() * 0.2 + 0.9) / Player.DEF) + 1
                    
                    If Damage > 10 ^ 7 - 1 Then Damage = 10 ^ 7 - 1
                    
                    Player.Hp = Player.Hp - Damage
                    SumDamage = SumDamage + Damage
       
                End With
            
            End If

    End Select
    
End Sub

Public Sub MusicCheck()

    Select Case GameMode
    
        Case Entrance, HowtoPlay, Options, Museum
        
            With Music(0)
        
                If .Alive = False Then
                
                    Call Lib.AllStopSound
                    Call Lib.PlaySound(0)
                    .Alive = True
                    
                ElseIf (.Alive = True) And (.Hp < .Width) Then
                
                    .Hp = .Hp + 1
                    
                Else
                
                    .Alive = False
                    .Hp = 0
                
                End If
            
            End With
            
        Case Dungeon
        
            With Music(3)
        
                If .Alive = False Then
                
                    Call Lib.AllStopSound
                    Call Lib.PlaySound(3)
                    .Alive = True
                    
                ElseIf (.Alive = True) And (.Hp < .Width) Then
                
                    .Hp = .Hp + 1
                    
                Else
                
                    .Alive = False
                    .Hp = 0
                
                End If
            
            End With

    End Select

End Sub

