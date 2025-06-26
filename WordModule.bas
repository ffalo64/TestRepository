Attribute VB_Name = "WordModule"
Public Words() As String
Public Oword() As Status

Public Sub WordSet()

    ReDim Preserve Words(11)
    ReDim Preserve Oword(3)
    Dim i As Long
    Dim j As Long
    
    Select Case GameMode
    
        Case Entrance
        
            Words(0) = "�_���W�����ƕs�v�c�̔�" & vbCrLf & ""
            Words(1) = "�_���W�����ɓ���(Enter�L�[)" & vbCrLf & "�s�v�c�̔�����g����`�����n�܂�܂��B" & vbCrLf & ""
            Words(2) = "How to play(z�L�[)" & vbCrLf & "���������T�v�����Ȃ�" & vbCrLf & ""
            Words(3) = "Option(c�L�[)" & vbCrLf & "��Փx��ݒ�" & vbCrLf & ""
            Words(4) = "Museum(d�L�[)" & vbCrLf & "�L�^�̕���" & vbCrLf & "" & vbCrLf & "(�������͔͂��p�ɂ��ĉ�����)"
            Words(5) = "" & vbCrLf & "���y�f�ޒ񋟌�"
            Words(6) = "�y�T�C�g���z�t���[���y�f�� H/MIX GALLERY" & vbCrLf & "�y�Ǘ��ҁz�@�H�R�T�a"
            Words(7) = "�y�A�h���X�zhttp://www.hmix.net/" & vbCrLf & ""
            Words(8) = "" & vbCrLf & ""

        Case Dungeon
            
            Words(4) = "�S�̍U��(z)" & vbCrLf & CStr(AbilityHp(0))
            Words(5) = "Hp�S��(x)" & vbCrLf & CStr(AbilityHp(1))
            Words(6) = "�S����(c)" & vbCrLf & CStr(AbilityHp(2))
            Words(7) = "�����X�^�[����(d)" & vbCrLf & CStr(AbilityHp(3))
            Words(8) = "���̊K��(Enter)" & vbCrLf & CStr(AbilityHp(4)) & vbCrLf & "�����̎�" & vbCrLf & CStr(AbilityHp(5))
            Words(9) = CStr(Floor) & "F Lv" & CStr(Player.Level) & " HP" & CStr(Int(Player.Hp)) & "/" & CStr(Player.MaxHp) & "  Turn " & Turn
            
            For i = 0 To UBound(Oword)
            
                With Oword(i)
            
                    If .Explanation <> Words(i) Then
                    
                       .Explanation = Words(i)
                       .Hp = 40
                       
                    ElseIf (.Explanation = Words(i)) And (.Hp > 0) Then
                    
                        .Hp = .Hp - 1
                        
                    ElseIf (.Explanation = Words(i)) And (.Hp = 0) Then

                        Words(i) = ""
                    
                    End If
                
                End With
            
            Next i
            
            For i = 0 To LandNumber - 1
            For j = 0 To LandNumber - 1
            
                With Landsquare(i, j)
                
                    Select Case .Condition
                    
                        Case BlueBox
                        
                            .Name = "��"
                        
                        Case RedBox
                        
                            .Name = "�Ԕ�"
                        
                        Case YellowBox
                        
                            .Name = "����"
                        
                        Case GreenBox
                        
                            .Name = "�Δ�"
                        
                        Case PurpleBox
                        
                            .Name = "����"
                            
                        Case Wall
                        
                            .Name = "��"
                    
                    End Select
                
                End With
                
            Next j
            Next i
            
        Case HowtoPlay
        
            Select Case Player.Condition
            
                Case 0
        
                    Words(0) = "How to play" & vbCrLf & "1/6"
                    Words(1) = "" & vbCrLf & "���̃Q�[���͊��ł����Ƃ��Ȃ�Q�[���ł��B"
                    Words(2) = "�Ȃ̂ŁA������ǂނ̂������Ȑl�́A" & vbCrLf & "�����͓ǂݔ�΂��Ă����v�ł��B"
                    Words(3) = "" & vbCrLf & ""
                    Words(4) = "" & vbCrLf & ""
                    Words(5) = "" & vbCrLf & ""
                    Words(6) = "" & vbCrLf & ""
                    Words(7) = "" & vbCrLf & ""
                    Words(8) = "" & vbCrLf & "(x�L�[�Ń��j���[��ʂɖ߂�,���E�L�[�Ńy�[�W�I��)"
            
                Case 1
        
                    Words(0) = "How to play" & vbCrLf & "2/6"
                    Words(1) = "���̃Q�[���̓����X�^�[��|���Ȃ���A" & vbCrLf & "���̊K��ڎw���ĊK�i���~��Ă����Q�[���ł��B"
                    Words(2) = "�v���C���[�@�@�@                     " & vbCrLf & "�����X�^�[�@�@�@                     "
                    Words(3) = "�K�i�@�@�@                     " & vbCrLf & "(�K�i�̓}�E�X�|�C���^�ƌ��ԈႦ�₷���̂Œ��ӂ��ĉ������B)"
                    Words(4) = "����͊�{�I�ɏ\���L�[�ɂ��ړ������ł��B" & vbCrLf & "�U��������̋���}�X�ɐi�����Ƃ��邾���ŏo���܂��B"
                    Words(5) = "" & vbCrLf & "�ǁ@�@�@                     �@"
                    Words(6) = "���@�@�@                     �@" & vbCrLf & ""
                    Words(7) = "�_���W�����̒n�`�͏��2��ނ����ł��B" & vbCrLf & "���̂����ǂ̏�ɂ͐i�ނ��Ƃ��o���܂���B"
                    Words(8) = "" & vbCrLf & "(x�L�[�Ń��j���[��ʂɖ߂�,���E�L�[�Ńy�[�W�I��)"
                    
                Case 2
        
                    Words(0) = "How to play" & vbCrLf & "3/6"
                    Words(1) = "�_���W�����ɏo������5�̔�" & vbCrLf & ""
                    Words(2) = "���E�E�E�󂷂�HP���񕜂��A�����X�^�[�����ʂ̏�Ԃɖ߂�B" & vbCrLf & "�Ԕ��E�E�E�󂷂ƃ}�C�i�X���ʂ���������B"
                    Words(3) = "�����E�E�E�󂷂ƐF�X�Ȍ��ʂ���������" & vbCrLf & "�Δ��E�E�E�󂷂ƃv���X���ʂ���������B"
                    Words(4) = "�����E�E�E�󂷂ƃR�}���h�L�[�̎g�p�񐔂�������B" & vbCrLf & ""
                    Words(5) = "�����ɂ���" & vbCrLf & ""
                    Words(6) = "���5�̔��̂����A�����͏�������ł��B" & vbCrLf & "���̔��͌��ʂȂǂŕʂ̔��ɕς�炸�A"
                    Words(7) = "��Փx�ɂ���Č��ʂ��Ⴂ�܂��B" & vbCrLf & ""
                    Words(8) = "" & vbCrLf & "(x�L�[�Ń��j���[��ʂɖ߂�,���E�L�[�Ńy�[�W�I��)"
            
                Case 3
        
                    Words(0) = "How to play" & vbCrLf & "4/6"
                    Words(1) = "�_���W�������Ŏg����5�̃R�}���h" & vbCrLf & ""
                    Words(2) = "z�L�[...�G�S���ɍU��" & vbCrLf & "x�L�[...Hp��S��������"
                    Words(3) = "c�L�[...�����ȊO�̑S�Ă̔��A�ǁA�����X�^�[����������B" & vbCrLf & "d�L�[...�����X�^�[�𔠂ɕω�������"
                    Words(4) = "Enter�L�[...����̃t���A�ɍ~���" & vbCrLf & ""
                    Words(5) = "(�����̂��Ƃ̓Q�[�����\������Ă���̂ŁA" & vbCrLf & "�ǂ̃L�[���ǂ�Ȍ��ʂ��͊o���Ȃ��đ��v�ł��B)"
                    Words(6) = "" & vbCrLf & ""
                    Words(7) = "" & vbCrLf & ""
                    Words(8) = "" & vbCrLf & "(x�L�[�Ń��j���[��ʂɖ߂�,���E�L�[�Ńy�[�W�I��)"
                    
                Case 4
        
                    Words(0) = "How to play" & vbCrLf & "5/6"
                    Words(1) = "�Z�[�u/���[�h�ɂ���" & vbCrLf & ""
                    Words(2) = "�Q�[���𒆒f�������Ȃ�������S�{�^���������ƃZ�[�u���o���܂��B" & vbCrLf & "�ĊJ���������͋L�^�̕������烍�[�h�ł��܂��B"
                    Words(3) = "�Z�[�u����ƁA�ȑO�̃f�[�^�͏����Ă��܂��̂ŁA���ӂ��ĉ������B" & vbCrLf & ""
                    Words(4) = "�^�[�����ɂ���" & vbCrLf & ""
                    Words(5) = "���̃Q�[���ł�1�^�[�����ƂɃ^�[������1�����Ă����܂��B" & vbCrLf & "�^�[������0�ɂȂ�ƃQ�[���I�[�o�[�ł��B"
                    Words(6) = "���̂��Ƃ�����ƁA�^�[�����͉񕜂��܂��B" & vbCrLf & ""
                    Words(7) = "�����󂷂��A�ǂ��@��(�\�͂���ɓ��ꂽ���̂�)�B" & vbCrLf & "�����̎�ŕ�������B"
                    Words(8) = "" & vbCrLf & "(x�L�[�Ń��j���[��ʂɖ߂�,���E�L�[�Ńy�[�W�I��)"
                    
                Case 5
        
                    Words(0) = "How to play" & vbCrLf & "6/6"
                    Words(1) = "�_���W�����ő厖��6�̃X�e�[�^�X" & vbCrLf & ""
                    Words(2) = "���x��" & vbCrLf & "���ꂪ�オ��ƁA�S�̓I�ɋ����Ȃ�܂��B"
                    Words(3) = "�U����" & vbCrLf & "���̐��l���傫���قǁA�^����_���[�W���傫���Ȃ�܂��B"
                    Words(4) = "�����" & vbCrLf & "���̐��l���傫���قǁA�󂯂�_���[�W�����Ȃ��Ȃ�܂��B"
                    Words(5) = "Hp" & vbCrLf & "�̗͂ł��B���ꂪ0�ɂȂ�Ɠ|����܂��B"
                    Words(6) = "�ő�Hp" & vbCrLf & "Hp�̍ő�l�ł��BHp�͂���ȏ�ɂ͉񕜂��܂���B"
                    Words(7) = "�o���l" & vbCrLf & "���ꂪ���܂�ƃ��x���A�b�v���Ă����܂��B"
                    Words(8) = "�����X�^�[��|���ƁA���̌o���l�������̂��̂ɂȂ�܂��B" & vbCrLf & "(x�L�[�Ń��j���[��ʂɖ߂�,���E�L�[�Ńy�[�W�I��)"
            
            End Select
            
        Case Options
        
            Words(0) = "Options" & vbCrLf & ""
            
            Select Case Player.Ability
            
                Case 0
                
                    Words(1) = "Very Easy" & vbCrLf & ""
                
                Case 1
                
                    Words(1) = "Easy" & vbCrLf & ""
            
                Case 2
            
                    Words(1) = "Normal" & vbCrLf & ""
                    
                Case 3
                
                    Words(1) = "Hard" & vbCrLf & ""
                
                Case 4
                
                    Words(1) = "Very Hard" & vbCrLf & ""
            
            End Select
            
            Words(2) = "�Ⴂ��Փx�قǎ����ŗǂ����ʂ��o�₷���ł��B" & vbCrLf & ""
            Words(3) = "" & vbCrLf & ""
            Words(4) = "x�L�[�Ń��j���[��ʂɖ߂�,���E�L�[�œ�Փx�I��" & vbCrLf & "Enter�L�[�Ń_���W�����ɓ���B"
            
        Case GameOver
        
            Form1.ForeColor = vbWhite
        
            Words(0) = "GameOver"
            Words(1) = ""
            Words(2) = "�v���C���[�͗͐s����"
            Words(3) = ""
            Words(4) = "x�L�[�Ń��j���[��ʂɖ߂�"
            
        Case GameClear
        
            Form1.BackColor = vbWhite
            Form1.ForeColor = vbBlue
                
            Words(0) = "GameClear" & vbCrLf & "�������ŉ��w��1000F�ł��B"
            Words(1) = "" & vbCrLf & "���̃Q�[���������܂ŗV��ł��ꂽ���Ȃ��͉ʕ�҂ł��B"
            Words(2) = "�N���A���̃X�e�[�^�X" & vbCrLf & ""
            Words(3) = "HP " & CStr(Int(Player.Hp)) & "/" & CStr(Player.MaxHp) & vbCrLf & "Lv " & CStr(Player.Level)
            Words(4) = "" & vbCrLf & "x�L�[�Ń��j���[��ʂɖ߂�"
            
        Case Museum
        
            Select Case Player.Condition
            
                Case 0
        
                    Words(0) = "���݂̋L�^" & vbCrLf & ""
                    Words(1) = CStr(Floor + 1) & "F Lv" & CStr(Player.Level) & vbCrLf & "HP" & CStr(Int(Player.Hp)) & "/" & CStr(Player.MaxHp) & "  Turn " & Turn
                    Words(2) = "�S�̍U�� " & CStr(AbilityHp(0)) & vbCrLf & "Hp�S�� " & CStr(AbilityHp(1)) & vbCrLf & "�S���� " & CStr(AbilityHp(2))
                    Words(3) = "�����X�^�[���� " & CStr(AbilityHp(3)) & vbCrLf & "���̊K�� " & CStr(AbilityHp(4)) & vbCrLf & "�����̎� " & CStr(AbilityHp(5))
                    Words(4) = "L�{�^���������ƁA���̃Z�[�u�f�[�^����n�߂��܂��B" & vbCrLf & "x�L�[�Ń��j���[��ʂɖ߂�"
            
            End Select

    End Select

End Sub

