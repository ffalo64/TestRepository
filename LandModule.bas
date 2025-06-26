Attribute VB_Name = "LandModule"
Public Landsquare() As Status
Public RandomPosition As Status
Public Map() As Long
Public Const BlueBox = 0
Public Const RedBox = 1
Public Const YellowBox = 2
Public Const GreenBox = 3
Public Const PurpleBox = 4
Public Const Stair = 5
Public Const Wall = 6
Public Const Room = 7
Public Const Enemy = 8
Public Const Mine = 9
Public Const LandNumber = 40
Public Const Sc = 10 'Square Const


Public Sub LandSet()

Dim i As Long
Dim j As Long

    Select Case GameMode
    
        Case Entrance
                
            ReDim AbilityHp(6)
            
            For i = 0 To UBound(AbilityHp)
            
                AbilityHp(i) = 0
            
            Next i
                                
            With Player
                
                .Alive = False
                .Height = Sc
                .Width = Sc
                .Left = Form1.ScaleWidth / 4
                .Top = Form1.FontSize * 2.5
                .Level = 1
                .MaxHp = 20
                .Hp = .MaxHp
                .Condition = 0
                .Direction = 1
            
            End With
            
            With MessageBar
                
                .Left = 0
                .Right = Form1.ScaleWidth
                .Top = Form1.FontSize * 2
                .Bottom = Form1.FontSize * 32
                
            End With
            
            With StatusBar
            
                .Left = 0
                .Right = Form1.ScaleWidth
                .Top = MessageBar.Bottom
                .Bottom = Form1.ScaleHeight
            
            End With
            
            ReDim Landsquare(LandNumber - 1, LandNumber - 1)
            
            For i = 0 To LandNumber - 1
            For j = 0 To LandNumber - 1
            
                With Landsquare(i, j)
                
                    .Width = Sc
                    .Height = Sc
                    .Left = .Width * i
                    .Top = .Height * j
                    .Condition = Room
                    .Alive = True
                
                End With
                
            Next j
            Next i
            
            Call Form1.Colorset

            Form1.BackColor = vbBlack
            Form1.ForeColor = vbWhite
            
            Floor = 0
                        
        Case Dungeon
            
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
            
            Player.Hp = Player.MaxHp
            Player.Alive = True
            
        Case HowtoPlay
        
            With Player
            
                .Width = Sc
                .Height = Sc
                .Left = Form1.ScaleWidth / 2
                .Top = Form1.FontSize * 10.5
            
            End With
                  
            With MessageBar
            
                .Left = 0
                .Right = Form1.ScaleWidth
                .Top = Form1.FontSize * 2
                .Bottom = Form1.FontSize * 22
            
            End With
            
            With StatusBar
            
                .Left = 0
                .Right = Form1.ScaleWidth
                .Top = MessageBar.Bottom
                .Bottom = Form1.ScaleHeight
            
            End With
            
            ReDim Landsquare(25, 0)
            
            For i = 0 To UBound(Landsquare)
            
                With Landsquare(i, 0)
                
                    .Width = Sc
                    .Height = Sc
                
                End With

            Next i
            
        Case GameOver
        
            Form1.ForeColor = vbWhite
        
            With MessageBar
                
                .Left = 0
                .Right = Form1.ScaleWidth
                .Top = Form1.FontSize * 2
                .Bottom = Form1.ScaleHeight
                
            End With
            
            Call Lib.AllStopSound
            Call Lib.PlaySound(2)
        
        Case GameClear
                
            With MessageBar
                
                .Left = 0
                .Right = Form1.ScaleWidth
                .Top = Form1.FontSize * 2
                .Bottom = Form1.ScaleHeight
                
            End With
            
            Call Lib.AllStopSound
            Call Lib.PlaySound(1)
            
        Case Museum
        
            With Player
                
                .Level = Lib.ReadText(6)
                .Hp = Lib.ReadText(1)
                .MaxHp = Lib.ReadText(2)
                .ATK = Lib.ReadText(3)
                .DEF = Lib.ReadText(4)
                .Exp = Lib.ReadText(5)
                .Level = Lib.ReadText(6)
                .Ability = Lib.ReadText(15)
            
            End With
            
            For i = 0 To 5
            
                AbilityHp(i) = Lib.ReadText(i + 7)
            
            Next i
            
            Floor = Lib.ReadText(13)
            Turn = Lib.ReadText(14)
            MonsterNumber = Lib.ReadText(7, True)
            
            Call StatusCheck(False)

    End Select

End Sub

Public Sub BoxEffect(i As Long, j As Long)

    Dim X As Long
    Dim Y As Long

    With Landsquare(i, j)
    
        Select Case .Condition
        
            Case BlueBox
                
                Player.Hp = Player.Hp + Player.MaxHp / 10
                
                For X = 0 To UBound(Monster)
                             
                    Monster(X).Ability = Noability
                        
                Next X
                
                .Explanation = "HPが少し回復し、モンスターの状態異常が解除された。"
                
            Case RedBox
            
                Select Case .Ability
                
                    Case 0
                        
                        Player.Hp = 1
                        .Explanation = "Hpが1になってしまった"
                    
                    Case 1
                        
                        Player.ATK = Player.ATK * 0.9
                        .Explanation = "攻撃力が下がった"
                    
                    Case 2
                    
                        Player.Condition = Slow
                        .Explanation = "プレイヤーはこのフロアにいる間、動きが鈍くなった。"
                    
                    Case 3
                        
                        Player.Condition = TurnConst
                        .Explanation = "このフロアにいる間、ターン数が増えなくなった。"
                    
                    Case 4
                        
                        Player.MaxHp = Player.MaxHp * 0.9
                        .Explanation = "最大Hpが下がった"
            
                    Case 5
                        
                        Player.Hp = Int(Player.Hp / 2) + 1
                        .Explanation = "Hpが半分になった"
                    
                    Case 6
                    
                        For X = 0 To UBound(Monster)
                        
                            With Monster(X)
                            
                                If .Alive = True Then
                                    
                                    .Hp = .MaxHp
                                    .Condition = 0
 
                                End If
                                                
                            End With
                        
                        Next X
                    
                        .Explanation = "モンスターが全快した。"
                    
                    Case 7
                    
                        For X = 0 To UBound(Monster)

                            Monster(X).ATK = Monster(X).ATK * 1.1

                        Next X
                        
                        .Explanation = "全てのモンスターの攻撃力が少し上がった"
                    
                    Case 8

                        Call StatusCheck(True)
                        
                        .Explanation = "全てのモンスターのレベルが上った"
                    
                    Case 9
                    
                        For X = 0 To LandNumber - 1
                        For Y = 0 To LandNumber - 1
                        
                            With Landsquare(X, Y)
                            
                                If .Condition <= PurpleBox Then
                                
                                    .DEF = .DEF * 2
                                
                                End If
                            
                            End With
                            
                        Next Y
                        Next X
                    
                        .Explanation = "このフロアの全ての箱の守備力が2倍になった。"
                        
                    Case 10
                    
                        For X = 0 To UBound(Monster)

                            Monster(X).DEF = Monster(X).DEF * 1.1

                        Next X
                        
                        .Explanation = "全てのモンスターの守備力が少し上がった"
                    
                    Case 11
                    
                        For X = 0 To LandNumber - 1
                        For Y = 0 To LandNumber - 1
                        
                            With Landsquare(X, Y)
                            
                                If .Condition <= GreenBox Then
                                
                                    .Condition = RedBox
                                
                                End If
                            
                            End With
                            
                        Next Y
                        Next X
                        
                        .Explanation = "全ての箱が赤色になった"
                    
                    Case 12
                    
                        For X = 0 To UBound(Monster)

                            Monster(X).MaxHp = Monster(X).MaxHp * 1.1

                        Next X
                        
                        .Explanation = "全てのモンスターの最大Hpが少し上がった"
                    
                    Case 13
                    
                        For X = 0 To UBound(Monster)

                            Monster(X).Exp = Monster(X).Exp * 0.9

                        Next X
                        
                        .Explanation = "全てのモンスターの経験値が少し下がった"
                    
                    Case 14
                    
                        For X = 0 To LandNumber - 1
                        For Y = 0 To LandNumber - 1
                        
                            With Landsquare(X, Y)
                            
                                If .Condition <= PurpleBox Then
                                
                                    .Hp = .Hp * 2
                                    .MaxHp = .MaxHp * 2
                                
                                End If
                            
                            End With
                            
                        Next Y
                        Next X
                    
                        .Explanation = "このフロアの全ての箱のHpが2倍になった。"
                
                End Select
            
            Case YellowBox
            
                Select Case .Ability
                
                    Case 0
                    
                        For X = 0 To UBound(Monster)
                             
                            Monster(X).Ability = Stealth
                        
                        Next X
            
                        .Explanation = "モンスターが透明になった。"
                    
                    Case 1
                        
                        For X = 0 To UBound(AbilityHp) - 1
                        
                            AbilityHp(X) = AbilityHp(X) + 1
                        
                        Next X
                        
                        .Explanation = "全てのコマンドの使用回数が1増えた。"
                    
                    Case 2
                        
                        Player.Condition = TurnConst
                        .Explanation = "このフロアにいる間、ターン数が増えなくなった。"
                    
                    Case 3
                    
                        For X = 0 To UBound(Monster)
                        
                            If Monster(X).Alive = True Then

                                Monster(X).Alive = False
                                Monster(X).Hp = 0
                                Monster(X).Condition = 0
                                Landsquare(Monster(X).Left / Sc, Monster(X).Top / Sc).Condition = Room
                                
                            End If

                        Next X
                    
                        .Explanation = "モンスターが全滅した。"
                    
                    Case 4
                    
                        Player.DEF = Player.DEF * 2
                        .Explanation = "守備力が2倍になった"
            
                    Case 5
                    
                        Player.ATK = Player.ATK * 2
                        .Explanation = "攻撃力が2倍になった"
                    
                    Case 6

                        Turn = 1000
                        .Explanation = "ターン数が残り1000になった。"
                    
                    Case 7
                        
                        Player.ATK = Player.ATK / 2
                        .Explanation = "攻撃力が半分になった"
                    
                    Case 8
                    
                        For X = 0 To UBound(Monster)
                             
                            Monster(X).Ability = WallBreak
                        
                        Next X
            
                        .Explanation = "モンスターが壁を掘れるようになった。(一番外側を除く)"
                    
                    Case 9
                    
                        For X = 0 To UBound(AbilityHp) - 1
                        
                            AbilityHp(X) = 5
                        
                        Next X
                        
                        .Explanation = "全てのコマンドの使用回数が5になった。"
                        
                    Case 10
                    
                        For X = 0 To LandNumber - 1
                        For Y = 0 To LandNumber - 1
                        
                            With Landsquare(X, Y)
                            
                                If .Condition <= GreenBox Then
                                
                                    .Condition = Room
                                
                                End If
                            
                            End With
                            
                        Next Y
                        Next X
                        
                        .Explanation = "全ての箱が消滅した。"
                    
                    Case 11
                    
                        For X = 0 To LandNumber - 1
                        For Y = 0 To LandNumber - 1
                        
                            With Landsquare(X, Y)
                            
                                If .Condition <= GreenBox Then
                                
                                    .Condition = BlueBox
                                
                                End If
                            
                            End With
                            
                        Next Y
                        Next X
                        
                        .Explanation = "全ての箱が青色になった"
                    
                    Case 12
                    
                        For X = 0 To LandNumber - 1
                        For Y = 0 To LandNumber - 1
                        
                            With Landsquare(X, Y)
                            
                                If .Condition <= GreenBox Then
                                
                                    .Condition = Int(Rnd() * 4)
                                
                                End If
                            
                            End With
                            
                        Next Y
                        Next X
                        
                        .Explanation = "全ての箱がランダムに変化した。"
                    
                    Case 13
                    
                        For X = 0 To LandNumber - 1
                        For Y = 0 To LandNumber - 1
                        
                            With Landsquare(X, Y)
                            
                                If .Condition <= GreenBox Then
                                
                                    .Condition = YellowBox
                                
                                End If
                            
                            End With
                            
                        Next Y
                        Next X
                        
                        .Explanation = "全ての箱が黄色になった"
                    
                    Case 14
                    
                        For X = 0 To UBound(Monster)
                             
                            Monster(X).Ability = Boxattack
                        
                        Next X
            
                        .Explanation = "モンスターが箱を壊せるようになった。"
                    
                End Select
            
            Case GreenBox
            
                Select Case .Ability
                
                    Case 0
                        
                        Player.Hp = Player.MaxHp
                        .Explanation = "Hpが全快した"
                    
                    Case 1
                        
                        Player.ATK = (Player.ATK * 1.1) + 1
                        .Explanation = "攻撃力が上がった"
                    
                    Case 2
                    
                        For X = 0 To UBound(Monster)
                        
                            If Monster(X).Alive = True Then

                                Monster(X).Alive = False
                                Monster(X).Hp = 0
                                Monster(X).Condition = 0
                                Landsquare(Monster(X).Left / Sc, Monster(X).Top / Sc).Condition = Room
                                
                            End If

                        Next X
                    
                        .Explanation = "モンスターが全滅した。"
                    
                    Case 3
                        
                        For X = 0 To UBound(Monster)
                             
                            Monster(X).Ability = Slow
                        
                        Next X
                        
                        .Explanation = "モンスターの動きが遅くなった。"
                    
                    Case 4
                        
                        Player.MaxHp = (Player.MaxHp * 1.1) + 1
                        .Explanation = "最大Hpが上がった"
                        
                    Case 5
                        
                        For X = 1 To LandNumber - 2
                        For Y = 1 To LandNumber - 2
                        
                            With Landsquare(X, Y)
                            
                                If .Condition = Wall Then
                                
                                    .Condition = Room
                                
                                End If
                            
                            End With
                            
                        Next Y
                        Next X
                    
                        .Explanation = "一番外側以外の壁が全て崩れた。"
                    
                    Case 6
                        
                        Player.Condition = Stealth
                        .Explanation = "プレイヤーは透明になった。(このフロアのみ)"
                    
                    Case 7
                        
                        For X = 0 To UBound(Monster)

                            Monster(X).Hp = 1

                        Next X
                        
                        .Explanation = "全てのモンスターのHpが残り1になった"
                        
                    Case 8
                    
                        For X = 0 To UBound(Monster)

                            Monster(X).ATK = Monster(X).ATK * 0.9

                        Next X
                        
                        .Explanation = "全てのモンスターの攻撃力が少し下がった"
                    
                    Case 9
                        
                        For X = 0 To UBound(Monster)

                            Monster(X).Exp = Monster(X).Exp + 5

                        Next X
                        
                        .Explanation = "全てのモンスターの経験値が少し上がった"
                        
                    Case 10
                    
                        For X = 0 To UBound(Monster)

                            Monster(X).MaxHp = Monster(X).MaxHp * 0.9

                        Next X
                        
                        .Explanation = "全てのモンスターの最大Hpが少し下がった"
                    
                    Case 11
                        
                        For X = 0 To LandNumber - 1
                        For Y = 0 To LandNumber - 1
                        
                            With Landsquare(X, Y)
                            
                                If .Condition <= PurpleBox Then
                                
                                    .Hp = 1
                                
                                End If
                            
                            End With
                            
                        Next Y
                        Next X

                        .Explanation = "全ての箱のHpが残り1になった。"
                    
                    Case 12
                    
                        For X = 0 To LandNumber - 1
                        For Y = 0 To LandNumber - 1
                        
                            With Landsquare(X, Y)
                            
                                If .Condition = RedBox Then
                                
                                    .Condition = Room
                                
                                End If
                            
                            End With
                            
                        Next Y
                        Next X
                        
                        .Explanation = "赤箱が消滅した。"
                    
                    Case 13
                        
                        Player.DEF = Player.DEF + 1
                        .Explanation = "守備力が少し上がった"
                    
                    Case 14
                        
                        Player.Condition = WallBreak
                        .Explanation = "このフロアにいる間、壁を掘れるようになった(一番外側を除く)"
                    
                End Select
                
            Case PurpleBox
            
                Y = Int(Rnd() * (5 - Player.Ability)) + 1

                Select Case .Ability
                
                    Case 0 To 2
                                            
                        AbilityHp(0) = AbilityHp(0) + Y
                        .Explanation = "全体攻撃の使用回数が" & CStr(Y) & "増えた。"
                    
                    Case 3 To 6
                                            
                        AbilityHp(1) = AbilityHp(1) + Y
                        .Explanation = "Hp全快の使用回数が" & CStr(Y) & "増えた。"
                    
                    Case 7
                    
                        AbilityHp(2) = AbilityHp(2) + Y
                        .Explanation = "全消去の使用回数が" & CStr(Y) & "増えた。"
                    
                    Case 8
                    
                        AbilityHp(3) = AbilityHp(3) + Y
                        .Explanation = "モンスター箱化の使用回数が" & CStr(Y) & "増えた。"
                    
                    Case 9 To 11
                        
                        AbilityHp(4) = AbilityHp(4) + Y
                        .Explanation = "次の階へ行く効果の使用回数が" & CStr(Y) & "増えた。"
                    
                    Case 12, 13, 14
                    
                        AbilityHp(5) = AbilityHp(5) + Y
                        .Explanation = "復活の珠が" & CStr(Y) & "個手に入った。"
                        
                End Select
                
            Case Wall
            
                .Explanation = ""
        
        End Select
        
        Call StatusCheck(False)

        Words(1) = .Explanation
        .Ability = 0
        .Condition = Room
        
    End With

End Sub
