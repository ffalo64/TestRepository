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
        
            Words(0) = "ダンジョンと不思議の箱" & vbCrLf & ""
            Words(1) = "ダンジョンに入る(Enterキー)" & vbCrLf & "不思議の箱を駆使する冒険が始まります。" & vbCrLf & ""
            Words(2) = "How to play(zキー)" & vbCrLf & "操作説明や概要説明など" & vbCrLf & ""
            Words(3) = "Option(cキー)" & vbCrLf & "難易度を設定" & vbCrLf & ""
            Words(4) = "Museum(dキー)" & vbCrLf & "記録の部屋" & vbCrLf & "" & vbCrLf & "(文字入力は半角にして下さい)"
            Words(5) = "" & vbCrLf & "音楽素材提供元"
            Words(6) = "【サイト名】フリー音楽素材 H/MIX GALLERY" & vbCrLf & "【管理者】　秋山裕和"
            Words(7) = "【アドレス】http://www.hmix.net/" & vbCrLf & ""
            Words(8) = "" & vbCrLf & ""

        Case Dungeon
            
            Words(4) = "全体攻撃(z)" & vbCrLf & CStr(AbilityHp(0))
            Words(5) = "Hp全快(x)" & vbCrLf & CStr(AbilityHp(1))
            Words(6) = "全消去(c)" & vbCrLf & CStr(AbilityHp(2))
            Words(7) = "モンスター箱化(d)" & vbCrLf & CStr(AbilityHp(3))
            Words(8) = "次の階へ(Enter)" & vbCrLf & CStr(AbilityHp(4)) & vbCrLf & "復活の珠" & vbCrLf & CStr(AbilityHp(5))
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
                        
                            .Name = "青箱"
                        
                        Case RedBox
                        
                            .Name = "赤箱"
                        
                        Case YellowBox
                        
                            .Name = "黄箱"
                        
                        Case GreenBox
                        
                            .Name = "緑箱"
                        
                        Case PurpleBox
                        
                            .Name = "紫箱"
                            
                        Case Wall
                        
                            .Name = "壁"
                    
                    End Select
                
                End With
                
            Next j
            Next i
            
        Case HowtoPlay
        
            Select Case Player.Condition
            
                Case 0
        
                    Words(0) = "How to play" & vbCrLf & "1/6"
                    Words(1) = "" & vbCrLf & "このゲームは勘でも何とかなるゲームです。"
                    Words(2) = "なので、説明を読むのが嫌いな人は、" & vbCrLf & "ここは読み飛ばしても大丈夫です。"
                    Words(3) = "" & vbCrLf & ""
                    Words(4) = "" & vbCrLf & ""
                    Words(5) = "" & vbCrLf & ""
                    Words(6) = "" & vbCrLf & ""
                    Words(7) = "" & vbCrLf & ""
                    Words(8) = "" & vbCrLf & "(xキーでメニュー画面に戻る,左右キーでページ選択)"
            
                Case 1
        
                    Words(0) = "How to play" & vbCrLf & "2/6"
                    Words(1) = "このゲームはモンスターを倒しながら、" & vbCrLf & "下の階を目指して階段を降りていくゲームです。"
                    Words(2) = "プレイヤー　　　                     " & vbCrLf & "モンスター　　　                     "
                    Words(3) = "階段　　　                     " & vbCrLf & "(階段はマウスポインタと見間違えやすいので注意して下さい。)"
                    Words(4) = "操作は基本的に十字キーによる移動だけです。" & vbCrLf & "攻撃も相手の居るマスに進もうとするだけで出来ます。"
                    Words(5) = "" & vbCrLf & "壁　　　                     　"
                    Words(6) = "床　　　                     　" & vbCrLf & ""
                    Words(7) = "ダンジョンの地形は上の2種類だけです。" & vbCrLf & "このうち壁の上には進むことが出来ません。"
                    Words(8) = "" & vbCrLf & "(xキーでメニュー画面に戻る,左右キーでページ選択)"
                    
                Case 2
        
                    Words(0) = "How to play" & vbCrLf & "3/6"
                    Words(1) = "ダンジョンに出現する5つの箱" & vbCrLf & ""
                    Words(2) = "青箱・・・壊すとHPが回復し、モンスターが普通の状態に戻る。" & vbCrLf & "赤箱・・・壊すとマイナス効果が発生する。"
                    Words(3) = "黄箱・・・壊すと色々な効果が発生する" & vbCrLf & "緑箱・・・壊すとプラス効果が発生する。"
                    Words(4) = "紫箱・・・壊すとコマンドキーの使用回数が増える。" & vbCrLf & ""
                    Words(5) = "紫箱について" & vbCrLf & ""
                    Words(6) = "上の5つの箱のうち、紫箱は少し特殊です。" & vbCrLf & "この箱は効果などで別の箱に変わらず、"
                    Words(7) = "難易度によって効果が違います。" & vbCrLf & ""
                    Words(8) = "" & vbCrLf & "(xキーでメニュー画面に戻る,左右キーでページ選択)"
            
                Case 3
        
                    Words(0) = "How to play" & vbCrLf & "4/6"
                    Words(1) = "ダンジョン内で使える5つのコマンド" & vbCrLf & ""
                    Words(2) = "zキー...敵全員に攻撃" & vbCrLf & "xキー...Hpを全快させる"
                    Words(3) = "cキー...紫箱以外の全ての箱、壁、モンスターを消去する。" & vbCrLf & "dキー...モンスターを箱に変化させる"
                    Words(4) = "Enterキー...一つ下のフロアに降りる" & vbCrLf & ""
                    Words(5) = "(これらのことはゲーム中表示されているので、" & vbCrLf & "どのキーがどんな効果かは覚えなくて大丈夫です。)"
                    Words(6) = "" & vbCrLf & ""
                    Words(7) = "" & vbCrLf & ""
                    Words(8) = "" & vbCrLf & "(xキーでメニュー画面に戻る,左右キーでページ選択)"
                    
                Case 4
        
                    Words(0) = "How to play" & vbCrLf & "5/6"
                    Words(1) = "セーブ/ロードについて" & vbCrLf & ""
                    Words(2) = "ゲームを中断したくなった時はSボタンを押すとセーブが出来ます。" & vbCrLf & "再開したい時は記録の部屋からロードできます。"
                    Words(3) = "セーブすると、以前のデータは消えてしまうので、注意して下さい。" & vbCrLf & ""
                    Words(4) = "ターン数について" & vbCrLf & ""
                    Words(5) = "このゲームでは1ターンごとにターン数が1減っていきます。" & vbCrLf & "ターン数が0になるとゲームオーバーです。"
                    Words(6) = "次のことをすると、ターン数は回復します。" & vbCrLf & ""
                    Words(7) = "箱を壊すか、壁を掘る(能力を手に入れた時のみ)。" & vbCrLf & "復活の珠で復活する。"
                    Words(8) = "" & vbCrLf & "(xキーでメニュー画面に戻る,左右キーでページ選択)"
                    
                Case 5
        
                    Words(0) = "How to play" & vbCrLf & "6/6"
                    Words(1) = "ダンジョンで大事な6つのステータス" & vbCrLf & ""
                    Words(2) = "レベル" & vbCrLf & "これが上がると、全体的に強くなります。"
                    Words(3) = "攻撃力" & vbCrLf & "この数値が大きいほど、与えるダメージが大きくなります。"
                    Words(4) = "守備力" & vbCrLf & "この数値が大きいほど、受けるダメージが少なくなります。"
                    Words(5) = "Hp" & vbCrLf & "体力です。これが0になると倒されます。"
                    Words(6) = "最大Hp" & vbCrLf & "Hpの最大値です。Hpはこれ以上には回復しません。"
                    Words(7) = "経験値" & vbCrLf & "これが貯まるとレベルアップしていきます。"
                    Words(8) = "モンスターを倒すと、その経験値が自分のものになります。" & vbCrLf & "(xキーでメニュー画面に戻る,左右キーでページ選択)"
            
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
            
            Words(2) = "低い難易度ほど紫箱で良い効果が出やすいです。" & vbCrLf & ""
            Words(3) = "" & vbCrLf & ""
            Words(4) = "xキーでメニュー画面に戻る,左右キーで難易度選択" & vbCrLf & "Enterキーでダンジョンに入る。"
            
        Case GameOver
        
            Form1.ForeColor = vbWhite
        
            Words(0) = "GameOver"
            Words(1) = ""
            Words(2) = "プレイヤーは力尽きた"
            Words(3) = ""
            Words(4) = "xキーでメニュー画面に戻る"
            
        Case GameClear
        
            Form1.BackColor = vbWhite
            Form1.ForeColor = vbBlue
                
            Words(0) = "GameClear" & vbCrLf & "ここが最下層の1000Fです。"
            Words(1) = "" & vbCrLf & "このゲームをここまで遊んでくれたあなたは果報者です。"
            Words(2) = "クリア時のステータス" & vbCrLf & ""
            Words(3) = "HP " & CStr(Int(Player.Hp)) & "/" & CStr(Player.MaxHp) & vbCrLf & "Lv " & CStr(Player.Level)
            Words(4) = "" & vbCrLf & "xキーでメニュー画面に戻る"
            
        Case Museum
        
            Select Case Player.Condition
            
                Case 0
        
                    Words(0) = "現在の記録" & vbCrLf & ""
                    Words(1) = CStr(Floor + 1) & "F Lv" & CStr(Player.Level) & vbCrLf & "HP" & CStr(Int(Player.Hp)) & "/" & CStr(Player.MaxHp) & "  Turn " & Turn
                    Words(2) = "全体攻撃 " & CStr(AbilityHp(0)) & vbCrLf & "Hp全快 " & CStr(AbilityHp(1)) & vbCrLf & "全消去 " & CStr(AbilityHp(2))
                    Words(3) = "モンスター箱化 " & CStr(AbilityHp(3)) & vbCrLf & "次の階へ " & CStr(AbilityHp(4)) & vbCrLf & "復活の珠 " & CStr(AbilityHp(5))
                    Words(4) = "Lボタンを押すと、このセーブデータから始められます。" & vbCrLf & "xキーでメニュー画面に戻る"
            
            End Select

    End Select

End Sub

