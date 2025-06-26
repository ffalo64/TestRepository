Attribute VB_Name = "AbilityModule"
Public Const VeryEasy = 0
Public Const Easy = 1
Public Const Normal = 2
Public Const Hard = 3
Public Const VeryHard = 4
Public AbilityHp() As Long

Public Sub AbilityEffect(keycode As Integer)

Dim i As Long
Dim j As Long
Dim Damage As Long

    Select Case keycode
    
        Case vbKeyA
        
            For i = 0 To LandNumber - 1
            For j = 0 To LandNumber - 1
            
                With Landsquare(i, j)
                
                    If (.Condition = GreenBox) Then

                        .Ability = Int(Rnd() * 15)
                        Turn = Turn + 10
                        Call BoxEffect(i, j)
                    
                    End If
                
                End With
                
            Next j
            Next i
            
            Call FloorSet
    
        Case vbKeyZ
        
            If AbilityHp(0) > 0 Then
        
                AbilityHp(0) = AbilityHp(0) - 1
                
                Damage = Int(Player.ATK * (Rnd() * 0.2 + 0.9) / Monster(0).DEF)
                If Damage > 10 ^ 7 - 1 Then Damage = 10 ^ 7 - 1

                For i = 0 To MonsterNumber - 1
                
                    With Monster(i)
                    
                        If .Alive = True Then
                
                            .Hp = .Hp - Damage
                            
                            If .Hp <= 0 Then
                            
                                .Hp = 0
                                .Alive = False
                                Landsquare(.Left / Sc, .Top / Sc).Condition = Room
                                Player.Exp = Player.Exp + .Exp
                            
                            End If
                            
                            Words(2) = "全てのモンスターは" & CStr(Damage) & "ダメージを受けた"

                        End If
                    
                    End With
            
                Next i
                
                Player.Direction = 11
                            
            End If
            
        Case vbKeyX
        
            If AbilityHp(1) > 0 Then
        
                AbilityHp(1) = AbilityHp(1) - 1
            
                Player.Hp = Player.MaxHp
            
            End If
            
        Case vbKeyC
        
            If AbilityHp(2) > 0 Then

                AbilityHp(2) = AbilityHp(2) - 1
                
                For i = 0 To LandNumber - 1
                For j = 0 To LandNumber - 1
                
                    With Landsquare(i, j)
                    
                        If (.Condition <= GreenBox) Or (.Condition = Wall) Then
                                                        
                            .Condition = Room
                        
                        End If
                    
                    End With
                    
                Next j
                Next i
                
                For X = 0 To MonsterNumber - 1
                
                    If Monster(X).Alive = True Then

                        Monster(X).Alive = False
                        Monster(X).Hp = 0
                        Monster(X).Condition = 0
                        Landsquare(Monster(X).Left / Sc, Monster(X).Top / Sc).Condition = Room
                        
                    End If

                Next X
                
            End If
            
        Case vbKeyD
        
            If (AbilityHp(3) > 0) And (Player.Alive = True) Then
            
                AbilityHp(3) = AbilityHp(3) - 1
            
                For i = 0 To MonsterNumber - 1
                
                    With Monster(i)
                
                        If .Alive = True Then
    
                            .Alive = False
                            .Hp = 0
                            .Condition = 0
                            Landsquare(.Left / Sc, .Top / Sc).Condition = Int(Rnd() * 4)
                            
                        End If
                    
                    End With

                Next i
                
            ElseIf (AbilityHp(5) > 0) And (Player.Alive = False) Then
            
                AbilityHp(5) = AbilityHp(5) - 1
                
                Player.Alive = True
                Player.Hp = Player.MaxHp
                If Player.Condition <> TurnConst Then Turn = Turn + 100
                Player.Condition = Noability
                Words(3) = "復活の珠の効果で復活した。"
                
                Call StatusCheck(False)

            End If
            
        Case vbKeyS
        
            If AbilityHp(6) = 0 Then
            
                AbilityHp(6) = AbilityHp(6) + 1
                Words(3) = "次に階段を降りた時、セーブしてメニュー画面に戻ります。"
        
            ElseIf (AbilityHp(6) > 0) And (Landsquare(Player.Left / Sc, Player.Top / Sc).Condition = Stair) Then
            
                Call StatusCheck(False)

                With Player
                
                    Call Lib.WriteText(CStr(.Hp), CStr(1))
                    Call Lib.WriteText(CStr(.MaxHp), CStr(2))
                    Call Lib.WriteText(CStr(.ATK), CStr(3))
                    Call Lib.WriteText(CStr(.DEF), CStr(4))
                    Call Lib.WriteText(CStr(.Exp), CStr(5))
                    Call Lib.WriteText(CStr(.Level), CStr(6))
                    Call Lib.WriteText(CStr(AbilityHp(0)), CStr(7))
                    Call Lib.WriteText(CStr(AbilityHp(1)), CStr(8))
                    Call Lib.WriteText(CStr(AbilityHp(2)), CStr(9))
                    Call Lib.WriteText(CStr(AbilityHp(3)), CStr(10))
                    Call Lib.WriteText(CStr(AbilityHp(4)), CStr(11))
                    Call Lib.WriteText(CStr(AbilityHp(5)), CStr(12))
                    Call Lib.WriteText(CStr(Floor), CStr(13))
                    Call Lib.WriteText(CStr(Turn), CStr(14))
                    Call Lib.WriteText(CStr(.Ability), CStr(15))
                
                End With
                
                With Monster(0)
                
                    Call Lib.WriteText(CStr(.MaxHp), CStr(1), True)
                    Call Lib.WriteText(CStr(.ATK), CStr(2), True)
                    Call Lib.WriteText(CStr(.DEF), CStr(3), True)
                    Call Lib.WriteText(CStr(.Exp), CStr(4), True)
                    Call Lib.WriteText(CStr(.Level), CStr(5), True)
                    Call Lib.WriteText(CStr(.Ability), CStr(6), True)
                    Call Lib.WriteText(CStr(MonsterNumber), CStr(7), True)
                
                End With
                
                GameMode = Entrance
                Call LandSet
                Call WordSet
                
                For i = 0 To UBound(Music)
            
                    Music(i).Hp = 0
                    Music(i).Alive = False
            
                Next i
            
            End If
            
        Case vbKeyL
        
            GameMode = Dungeon
            
            Call LandSet
            Call MonsterSet
            
            For i = 0 To UBound(Monster)
            
                With Monster(i)
                    
                    .MaxHp = Lib.ReadText(1, True)
                    .Hp = .MaxHp
                    .ATK = Lib.ReadText(2, True)
                    .DEF = Lib.ReadText(3, True)
                    .Exp = Lib.ReadText(4, True)
                    .Level = Lib.ReadText(5, True)
                    .Ability = Lib.ReadText(6, True)
                
                End With
            
            Next i

            Call FloorSet
            
        Case vbKeyReturn
        
            If AbilityHp(4) > 0 Then

                AbilityHp(4) = AbilityHp(4) - 1
                
                Call FloorSet
            
            End If
            
    End Select

End Sub
