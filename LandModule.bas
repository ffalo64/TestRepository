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
                
                .Explanation = "HP�������񕜂��A�����X�^�[�̏�Ԉُ킪�������ꂽ�B"
                
            Case RedBox
            
                Select Case .Ability
                
                    Case 0
                        
                        Player.Hp = 1
                        .Explanation = "Hp��1�ɂȂ��Ă��܂���"
                    
                    Case 1
                        
                        Player.ATK = Player.ATK * 0.9
                        .Explanation = "�U���͂���������"
                    
                    Case 2
                    
                        Player.Condition = Slow
                        .Explanation = "�v���C���[�͂��̃t���A�ɂ���ԁA�������݂��Ȃ����B"
                    
                    Case 3
                        
                        Player.Condition = TurnConst
                        .Explanation = "���̃t���A�ɂ���ԁA�^�[�����������Ȃ��Ȃ����B"
                    
                    Case 4
                        
                        Player.MaxHp = Player.MaxHp * 0.9
                        .Explanation = "�ő�Hp����������"
            
                    Case 5
                        
                        Player.Hp = Int(Player.Hp / 2) + 1
                        .Explanation = "Hp�������ɂȂ���"
                    
                    Case 6
                    
                        For X = 0 To UBound(Monster)
                        
                            With Monster(X)
                            
                                If .Alive = True Then
                                    
                                    .Hp = .MaxHp
                                    .Condition = 0
 
                                End If
                                                
                            End With
                        
                        Next X
                    
                        .Explanation = "�����X�^�[���S�������B"
                    
                    Case 7
                    
                        For X = 0 To UBound(Monster)

                            Monster(X).ATK = Monster(X).ATK * 1.1

                        Next X
                        
                        .Explanation = "�S�Ẵ����X�^�[�̍U���͂������オ����"
                    
                    Case 8

                        Call StatusCheck(True)
                        
                        .Explanation = "�S�Ẵ����X�^�[�̃��x���������"
                    
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
                    
                        .Explanation = "���̃t���A�̑S�Ă̔��̎���͂�2�{�ɂȂ����B"
                        
                    Case 10
                    
                        For X = 0 To UBound(Monster)

                            Monster(X).DEF = Monster(X).DEF * 1.1

                        Next X
                        
                        .Explanation = "�S�Ẵ����X�^�[�̎���͂������オ����"
                    
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
                        
                        .Explanation = "�S�Ă̔����ԐF�ɂȂ���"
                    
                    Case 12
                    
                        For X = 0 To UBound(Monster)

                            Monster(X).MaxHp = Monster(X).MaxHp * 1.1

                        Next X
                        
                        .Explanation = "�S�Ẵ����X�^�[�̍ő�Hp�������オ����"
                    
                    Case 13
                    
                        For X = 0 To UBound(Monster)

                            Monster(X).Exp = Monster(X).Exp * 0.9

                        Next X
                        
                        .Explanation = "�S�Ẵ����X�^�[�̌o���l��������������"
                    
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
                    
                        .Explanation = "���̃t���A�̑S�Ă̔���Hp��2�{�ɂȂ����B"
                
                End Select
            
            Case YellowBox
            
                Select Case .Ability
                
                    Case 0
                    
                        For X = 0 To UBound(Monster)
                             
                            Monster(X).Ability = Stealth
                        
                        Next X
            
                        .Explanation = "�����X�^�[�������ɂȂ����B"
                    
                    Case 1
                        
                        For X = 0 To UBound(AbilityHp) - 1
                        
                            AbilityHp(X) = AbilityHp(X) + 1
                        
                        Next X
                        
                        .Explanation = "�S�ẴR�}���h�̎g�p�񐔂�1�������B"
                    
                    Case 2
                        
                        Player.Condition = TurnConst
                        .Explanation = "���̃t���A�ɂ���ԁA�^�[�����������Ȃ��Ȃ����B"
                    
                    Case 3
                    
                        For X = 0 To UBound(Monster)
                        
                            If Monster(X).Alive = True Then

                                Monster(X).Alive = False
                                Monster(X).Hp = 0
                                Monster(X).Condition = 0
                                Landsquare(Monster(X).Left / Sc, Monster(X).Top / Sc).Condition = Room
                                
                            End If

                        Next X
                    
                        .Explanation = "�����X�^�[���S�ł����B"
                    
                    Case 4
                    
                        Player.DEF = Player.DEF * 2
                        .Explanation = "����͂�2�{�ɂȂ���"
            
                    Case 5
                    
                        Player.ATK = Player.ATK * 2
                        .Explanation = "�U���͂�2�{�ɂȂ���"
                    
                    Case 6

                        Turn = 1000
                        .Explanation = "�^�[�������c��1000�ɂȂ����B"
                    
                    Case 7
                        
                        Player.ATK = Player.ATK / 2
                        .Explanation = "�U���͂������ɂȂ���"
                    
                    Case 8
                    
                        For X = 0 To UBound(Monster)
                             
                            Monster(X).Ability = WallBreak
                        
                        Next X
            
                        .Explanation = "�����X�^�[���ǂ��@���悤�ɂȂ����B(��ԊO��������)"
                    
                    Case 9
                    
                        For X = 0 To UBound(AbilityHp) - 1
                        
                            AbilityHp(X) = 5
                        
                        Next X
                        
                        .Explanation = "�S�ẴR�}���h�̎g�p�񐔂�5�ɂȂ����B"
                        
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
                        
                        .Explanation = "�S�Ă̔������ł����B"
                    
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
                        
                        .Explanation = "�S�Ă̔����F�ɂȂ���"
                    
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
                        
                        .Explanation = "�S�Ă̔��������_���ɕω������B"
                    
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
                        
                        .Explanation = "�S�Ă̔������F�ɂȂ���"
                    
                    Case 14
                    
                        For X = 0 To UBound(Monster)
                             
                            Monster(X).Ability = Boxattack
                        
                        Next X
            
                        .Explanation = "�����X�^�[�������󂹂�悤�ɂȂ����B"
                    
                End Select
            
            Case GreenBox
            
                Select Case .Ability
                
                    Case 0
                        
                        Player.Hp = Player.MaxHp
                        .Explanation = "Hp���S������"
                    
                    Case 1
                        
                        Player.ATK = (Player.ATK * 1.1) + 1
                        .Explanation = "�U���͂��オ����"
                    
                    Case 2
                    
                        For X = 0 To UBound(Monster)
                        
                            If Monster(X).Alive = True Then

                                Monster(X).Alive = False
                                Monster(X).Hp = 0
                                Monster(X).Condition = 0
                                Landsquare(Monster(X).Left / Sc, Monster(X).Top / Sc).Condition = Room
                                
                            End If

                        Next X
                    
                        .Explanation = "�����X�^�[���S�ł����B"
                    
                    Case 3
                        
                        For X = 0 To UBound(Monster)
                             
                            Monster(X).Ability = Slow
                        
                        Next X
                        
                        .Explanation = "�����X�^�[�̓������x���Ȃ����B"
                    
                    Case 4
                        
                        Player.MaxHp = (Player.MaxHp * 1.1) + 1
                        .Explanation = "�ő�Hp���オ����"
                        
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
                    
                        .Explanation = "��ԊO���ȊO�̕ǂ��S�ĕ��ꂽ�B"
                    
                    Case 6
                        
                        Player.Condition = Stealth
                        .Explanation = "�v���C���[�͓����ɂȂ����B(���̃t���A�̂�)"
                    
                    Case 7
                        
                        For X = 0 To UBound(Monster)

                            Monster(X).Hp = 1

                        Next X
                        
                        .Explanation = "�S�Ẵ����X�^�[��Hp���c��1�ɂȂ���"
                        
                    Case 8
                    
                        For X = 0 To UBound(Monster)

                            Monster(X).ATK = Monster(X).ATK * 0.9

                        Next X
                        
                        .Explanation = "�S�Ẵ����X�^�[�̍U���͂�������������"
                    
                    Case 9
                        
                        For X = 0 To UBound(Monster)

                            Monster(X).Exp = Monster(X).Exp + 5

                        Next X
                        
                        .Explanation = "�S�Ẵ����X�^�[�̌o���l�������オ����"
                        
                    Case 10
                    
                        For X = 0 To UBound(Monster)

                            Monster(X).MaxHp = Monster(X).MaxHp * 0.9

                        Next X
                        
                        .Explanation = "�S�Ẵ����X�^�[�̍ő�Hp��������������"
                    
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

                        .Explanation = "�S�Ă̔���Hp���c��1�ɂȂ����B"
                    
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
                        
                        .Explanation = "�Ԕ������ł����B"
                    
                    Case 13
                        
                        Player.DEF = Player.DEF + 1
                        .Explanation = "����͂������オ����"
                    
                    Case 14
                        
                        Player.Condition = WallBreak
                        .Explanation = "���̃t���A�ɂ���ԁA�ǂ��@���悤�ɂȂ���(��ԊO��������)"
                    
                End Select
                
            Case PurpleBox
            
                Y = Int(Rnd() * (5 - Player.Ability)) + 1

                Select Case .Ability
                
                    Case 0 To 2
                                            
                        AbilityHp(0) = AbilityHp(0) + Y
                        .Explanation = "�S�̍U���̎g�p�񐔂�" & CStr(Y) & "�������B"
                    
                    Case 3 To 6
                                            
                        AbilityHp(1) = AbilityHp(1) + Y
                        .Explanation = "Hp�S���̎g�p�񐔂�" & CStr(Y) & "�������B"
                    
                    Case 7
                    
                        AbilityHp(2) = AbilityHp(2) + Y
                        .Explanation = "�S�����̎g�p�񐔂�" & CStr(Y) & "�������B"
                    
                    Case 8
                    
                        AbilityHp(3) = AbilityHp(3) + Y
                        .Explanation = "�����X�^�[�����̎g�p�񐔂�" & CStr(Y) & "�������B"
                    
                    Case 9 To 11
                        
                        AbilityHp(4) = AbilityHp(4) + Y
                        .Explanation = "���̊K�֍s�����ʂ̎g�p�񐔂�" & CStr(Y) & "�������B"
                    
                    Case 12, 13, 14
                    
                        AbilityHp(5) = AbilityHp(5) + Y
                        .Explanation = "�����̎삪" & CStr(Y) & "��ɓ������B"
                        
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
