Attribute VB_Name = "MonsterModule"
Public Monster() As Status
Public MonsterNumber As Long
Public Const Noability = 0
Public Const WallBreak = 1
Public Const Slow = 2
Public Const Boxattack = 3
Public Const TurnConst = 3
Public Const Stealth = 4

Public Sub MonsterSet()

    Select Case GameMode
    
        Case Entrance, Dungeon

            ReDim Preserve Monster(999)
            
            Dim i As Long
            
            For i = 0 To UBound(Monster)
            
                With Monster(i)
                
                    .Name = "ヘキサスライム" & CStr(i)
                    .Explanation = ""
                    .Hp = 10
                    .MaxHp = 10
                    .Level = 1
                    .Exp = 10
                    .Ability = Noability
                    .ATK = 10
                    .DEF = 1
                    .Height = Sc
                    .Width = Sc
                    .Direction = 1
            
                End With
        
            Next i
            
        Case HowtoPlay
        
            ReDim Preserve Monster(4)
            
            For i = 0 To UBound(Monster)
            
                With Monster(i)
                
                    .Name = "ヘキサスライム" & CStr(i)
                    .Left = (Form1.ScaleWidth / 2) + Sc * 2 * i
                    .Top = Player.Top + Form1.FontSize * 2
                    .Explanation = ""
                    .Ability = i
                    .Height = Sc
                    .Width = Sc
            
                End With
        
            Next i
    
    End Select
    
End Sub





