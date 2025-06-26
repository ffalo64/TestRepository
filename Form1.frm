VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'å≈íË(é¿ê¸)
   Caption         =   "É_ÉìÉWÉáÉìÇ∆ïsévãcÇÃî†"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "ÉÅÉCÉäÉI"
      Size            =   14.25
      Charset         =   128
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Àﬂ∏æŸ
   ScaleWidth      =   600
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Ç»Çµ
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   1
      Left            =   480
      Picture         =   "Form1.frx":058A
      ScaleHeight     =   50
      ScaleMode       =   3  'Àﬂ∏æŸ
      ScaleWidth      =   20
      TabIndex        =   6
      Top             =   6000
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Ç»Çµ
      BeginProperty Font 
         Name            =   "ÇlÇr Çoñæí©"
         Size            =   20.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4800
      Index           =   0
      Left            =   0
      Picture         =   "Form1.frx":1184
      ScaleHeight     =   320
      ScaleMode       =   3  'Àﬂ∏æŸ
      ScaleWidth      =   320
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   4800
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Ç»Çµ
      BeginProperty Font 
         Name            =   "ÇlÇr Çoñæí©"
         Size            =   24
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   0
      Left            =   960
      Picture         =   "Form1.frx":4C1C6
      ScaleHeight     =   100
      ScaleMode       =   3  'Àﬂ∏æŸ
      ScaleWidth      =   20
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Ç»Çµ
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   75
      Index           =   1
      Left            =   -240
      Picture         =   "Form1.frx":4D978
      ScaleHeight     =   5
      ScaleMode       =   3  'Àﬂ∏æŸ
      ScaleWidth      =   800
      TabIndex        =   3
      Top             =   7920
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   600
      Top             =   9480
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Ç»Çµ
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   0
      Left            =   2520
      Picture         =   "Form1.frx":5089A
      ScaleHeight     =   100
      ScaleMode       =   3  'Àﬂ∏æŸ
      ScaleWidth      =   20
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Ç»Çµ
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":514AC
      ScaleHeight     =   60
      ScaleMode       =   3  'Àﬂ∏æŸ
      ScaleWidth      =   20
      TabIndex        =   1
      Top             =   5760
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Ç»Çµ
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Index           =   0
      Left            =   1560
      Picture         =   "Form1.frx":522FE
      ScaleHeight     =   50
      ScaleMode       =   3  'Àﬂ∏æŸ
      ScaleWidth      =   20
      TabIndex        =   0
      Top             =   6000
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(keycode As Integer, Shift As Integer)
    
     With Player
     
        Select Case GameMode
        
            Case Entrance
            
                Select Case keycode
                        
                    Case vbKeyZ
                    
                        GameMode = HowtoPlay
                        Call LandSet
                        Call MonsterSet
                                     
                    Case vbKeyC
                    
                        GameMode = Options
                        
                    Case vbKeyD
                    
                        GameMode = Museum
                        Call LandSet
                                                            
                    Case vbKeyReturn
                    
                        GameMode = Dungeon
                        Call LandSet
                        Call MonsterSet
                        Call FloorSet
        
                End Select
     
            Case Dungeon
    
                Select Case keycode
                
                    Case vbKeyUp
                        
                        .Direction = .Direction * 2
                    
                    Case vbKeyDown
                        
                        .Direction = .Direction * 3
                        
                    Case vbKeyRight
                        
                        .Direction = .Direction * 5
                    
                    Case vbKeyLeft
                    
                        .Direction = .Direction * 7
        
                    Case vbKeyZ, vbKeyX, vbKeyC, vbKeyD, vbKeyS, vbKeyReturn, vbKeyA
                        
                        Call AbilityEffect(keycode)
                
                End Select
            
            Case HowtoPlay
        
                Select Case keycode
                
                    Case vbKeyRight
                        
                        .Direction = .Direction * 5
                    
                    Case vbKeyLeft
                    
                        .Direction = .Direction * 7
                                            
                    Case vbKeyX
                    
                        GameMode = Entrance
                        Call LandSet
                        Call WordSet
        
                End Select
                
            Case Options
            
                Select Case keycode
                                            
                    Case vbKeyRight
                        
                        .Direction = .Direction * 5
                    
                    Case vbKeyLeft
                    
                        .Direction = .Direction * 7
                                            
                    Case vbKeyX
                    
                        GameMode = Entrance
                        Call LandSet
                        Call WordSet
                        
                    Case vbKeyReturn
                    
                        GameMode = Dungeon
                        Call LandSet
                        Call MonsterSet
                        Call FloorSet
        
                End Select
                
            Case GameOver
            
                Select Case keycode
                               
                    Case vbKeyX
                        
                        Call MusicSet
                        GameMode = Entrance
                        Call LandSet
                        Call WordSet
                        
                End Select
                
            Case GameClear
            
                Select Case keycode
                               
                    Case vbKeyX
                        
                        Call MusicSet
                        GameMode = Entrance
                        Call LandSet
                        Call WordSet
                        
                End Select
                
            Case Museum
            
                Select Case keycode
                
                    Case vbKeyRight
                        
                        .Direction = .Direction * 5
                    
                    Case vbKeyLeft
                    
                        .Direction = .Direction * 7
                        
                    Case vbKeyX
                    
                        GameMode = Entrance
                        Call LandSet
                        Call WordSet
                        
                    Case vbKeyL

                        Call AbilityEffect(keycode)
                        
                End Select
     
        End Select
        
    End With
    
End Sub

Private Sub Form_Load()

    Randomize
    
    GameMode = Entrance
    
    With HpBar
    
        .Alive = True
        .Height = Picture1(1).ScaleHeight
        .Width = Picture1(1).ScaleWidth / 2
        .Left = Form1.ScaleWidth - .Width
        .Top = LandNumber * Sc + Form1.FontSize * 1.5
    
    End With
        
    Call MusicSet
    Call MonsterSet
    Call LandSet
    Call WordSet
   
End Sub
Private Sub MusicSet()

    Select Case GameMode
    
        Case Entrance

            ReDim Music(3)
        
            Music(0).Width = Int(185 * (1000 / Timer1.Interval))
            Music(3).Width = Int(255 * (1000 / Timer1.Interval) * 0.8)
            
        Case GameOver, GameClear
        
            For i = 0 To UBound(Music)
        
                Music(i).Hp = 0
                Music(i).Alive = False
        
            Next i
    
    End Select

End Sub
Public Sub Colorset()

Dim i As Long

    For i = 1 To 6
    
        Landsquare(i, 1).Direction = GetPixel(Picture5(0).hDC, i, 1)
    
    Next i

End Sub
Public Sub MapSet()

    Dim X As Long
    Dim Y As Long
    Dim Z As Long
    
    X = Int(Rnd() * Int(Picture5(0).Width / LandNumber)) * LandNumber
    Y = Int(Rnd() * Int(Picture5(0).Height / LandNumber)) * LandNumber
    
    For i = 0 To LandNumber - 1
    For j = 0 To LandNumber - 1
    
        With Landsquare(i, j)
    
            Z = GetPixel(Picture5(0).hDC, X + i, Y + j)
            
            Select Case Z
            
                Case Landsquare(1, 1).Direction
                
                    .Condition = BlueBox
                
                Case Landsquare(2, 1).Direction
                
                    .Condition = RedBox

                Case Landsquare(3, 1).Direction
                
                    .Condition = YellowBox

                Case Landsquare(4, 1).Direction
                
                    .Condition = GreenBox

                Case Landsquare(5, 1).Direction
                
                    .Condition = Wall
                
                Case Landsquare(6, 1).Direction
                
                    .Condition = Room

            End Select
            
            .MaxHp = Monster(0).MaxHp
            .Hp = .MaxHp
            .DEF = Monster(0).DEF
            .Alive = True

            If .Condition <> Room Then RandomPosition.Hp = RandomPosition.Hp - 1
            
        End With
        
    Next j
    Next i

End Sub
Private Sub Form_Unload(Cancel As Integer)

    Call Lib.EndSound

End Sub

Private Sub Timer1_Timer()

Me.Cls
Dim i As Long
Dim j As Long

    Select Case GameMode
    
        Case Entrance
                        
            Call DrawText(Form1.hDC, Words(0) & vbCrLf & Words(1) & vbCrLf & Words(2) & vbCrLf & Words(3) & vbCrLf & Words(4), -1, MessageBar, DT_CENTER)
            Call DrawText(Form1.hDC, Words(5) & vbCrLf & Words(6) & vbCrLf & Words(7) & vbCrLf & Words(8), -1, StatusBar, DT_CENTER)
            Call WordSet
            Call MusicCheck
            Call Movement
    
        Case Options, GameOver, GameClear, Museum
        
            Call DrawText(Form1.hDC, Words(0) & vbCrLf & Words(1) & vbCrLf & Words(2) & vbCrLf & Words(3) & vbCrLf & Words(4), -1, MessageBar, DT_CENTER)
            Call WordSet
            Call MusicCheck
            Call Movement

        Case Dungeon
                
            With HpBar
                            
                For i = 0 To .Height - 1
                
                    If Player.MaxHp <= .Width - 1 Then
                
                        For j = 0 To Player.Hp - 1
                        
                            Call SetPixel(Picture1(1).hDC, j, i, RGB(34, 177, 76))
                        
                        Next j
                        
                        For j = Player.Hp To Player.MaxHp - 1
                        
                            Call SetPixel(Picture1(1).hDC, j, i, vbRed)
                        
                        Next j
                        
                        For j = Player.MaxHp To .Width - 1
                        
                            Call SetPixel(Picture1(1).hDC, j, i, vbBlack)
                        
                        Next j
                        
                    ElseIf Player.MaxHp > .Width - 1 Then
                    
                        For j = 0 To Int((Player.Hp / Player.MaxHp) * .Width)
                        
                            Call SetPixel(Picture1(1).hDC, j, i, RGB(34, 177, 76))
                        
                        Next j

                        For j = Int((Player.Hp / Player.MaxHp) * .Width) + 1 To .Width - 1
                        
                            Call SetPixel(Picture1(1).hDC, j, i, vbRed)
                        
                        Next j
       
                    End If
                
                Next i
                
                Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture1(1).hDC, .Width, 0, vbSrcAnd)
                Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture1(1).hDC, 0, 0, vbSrcPaint)
    
            End With
                        
            For i = 0 To LandNumber - 1
            For j = 0 To LandNumber - 1
        
                 With Landsquare(i, j)
                    
                    If .Condition = Wall Then
                    
                        Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture4(0).hDC, .Width, .Height * Int((Floor Mod 200) / 20), vbSrcAnd)
                        Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture4(0).hDC, 0, .Height * Int((Floor Mod 200) / 20), vbSrcPaint)
                   
                    ElseIf (.Condition = Room) Or (.Condition = Enemy) Then
                    
                        Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture3(0).hDC, .Width, .Height * Int((Floor Mod 200) / 20), vbSrcAnd)
                        Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture3(0).hDC, 0, .Height * Int((Floor Mod 200) / 20), vbSrcPaint)
                        
                    ElseIf .Condition <= Stair Then
                    
                        Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture2(0).hDC, .Width, .Condition * .Height, vbSrcAnd)
                        Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture2(0).hDC, 0, .Condition * .Height, vbSrcPaint)
                    
                    End If
                                                                                     
                End With
    
            Next j
            Next i
            
            If Player.Alive = True Then
                
                With Player
                                        
                    Call Movement
                    Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture1(0).hDC, .Width, .Condition * .Height, vbSrcAnd)
                    Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture1(0).hDC, 0, .Condition * .Height, vbSrcPaint)
        
                End With
            
            End If
                        
            For i = 0 To MonsterNumber - 1
            
                If Monster(i).Alive = True Then
                
                    With Monster(i)
                        
                        Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture2(1).hDC, .Width, .Ability * .Height, vbSrcAnd)
                        Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture2(1).hDC, 0, .Ability * .Height, vbSrcPaint)
                           
                    End With
                
                End If
            
            Next i
            
            Call StatusCheck(False)
            Call WordSet
            Call MusicCheck
            Call DrawText(Form1.hDC, Words(9), -1, StatusBar, DT_CENTER)
            Call DrawText(Form1.hDC, Words(0) & vbCrLf & Words(1) & vbCrLf & Words(2) & vbCrLf & Words(3), -1, MessageBar, DT_CENTER)
            Call DrawText(Form1.hDC, Words(4) & vbCrLf & Words(5) & vbCrLf & Words(6) & vbCrLf & Words(7) & vbCrLf & Words(8), -1, SpecBar, DT_LEFT)
            
        Case HowtoPlay
        
            Call DrawText(Form1.hDC, Words(0) & vbCrLf & Words(1) & vbCrLf & Words(2) & vbCrLf & Words(3) & vbCrLf & Words(4), -1, MessageBar, DT_CENTER)
            Call DrawText(Form1.hDC, Words(5) & vbCrLf & Words(6) & vbCrLf & Words(7) & vbCrLf & Words(8), -1, StatusBar, DT_CENTER)
            Call WordSet
            Call Movement
            
            For i = 0 To UBound(Monster)
            
                If Player.Condition = 1 Then
                
                    With Monster(i)
                        
                        Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture2(1).hDC, .Width, .Ability * .Height, vbSrcAnd)
                        Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture2(1).hDC, 0, .Ability * .Height, vbSrcPaint)
                           
                    End With
                
                End If
            
            Next i
            
            For i = 0 To UBound(Landsquare)
        
                 With Landsquare(i, 0)
                 
                    If .Alive = True Then
                    
                        If .Condition = Wall Then
                         
                            Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture4(0).hDC, .Width, .Ability * .Height, vbSrcAnd)
                            Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture4(0).hDC, 0, .Ability * .Height, vbSrcPaint)
                        
                        ElseIf .Condition = Room Then
                         
                            Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture3(0).hDC, .Width, .Ability * .Height, vbSrcAnd)
                            Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture3(0).hDC, 0, .Ability * .Height, vbSrcPaint)
                            
                        ElseIf .Condition = Mine Then
                         
                            Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture1(0).hDC, .Width, .Ability * .Height, vbSrcAnd)
                            Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture1(0).hDC, 0, .Ability * .Height, vbSrcPaint)
                             
                        ElseIf .Condition <= Stair Then
                         
                            Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture2(0).hDC, .Width, .Condition * .Height, vbSrcAnd)
                            Call BitBlt(Me.hDC, .Left, .Top, .Width, .Height, Picture2(0).hDC, 0, .Condition * .Height, vbSrcPaint)
                         
                        End If
                    
                    End If
                                                                                     
                End With

            Next i

    End Select
            
    Me.Refresh
        
End Sub
