VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm登録種類選択 
   Caption         =   "登録"
   ClientHeight    =   1005
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4620
   OleObjectBlob   =   "UserForm登録種類選択.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm登録種類選択"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' 2010.7.3　原本ファイルの名簿シートの列の並びの見直しにより修正
Private Sub CommandButton1_Click()

  'オプションボタン1（入会）をクリックした場合
  If OptionButton1.Value = True Then
    
    '処理に必要なファイルが開かれているか確認する
    If IsBookOpen("郵便番号ﾃﾞｰﾀ【全国版】.xls") = False Then
      MsgBox "「郵便番号ﾃﾞｰﾀ【全国版】.xls」が開かれていません！" _
          & vbNewLine & "開いてからやり直して下さい。"
      End
    End If
    If IsBookOpen("東京東筑会名簿【入退会者一覧】.xls") = False Then
      MsgBox "「東京東筑会名簿【入退会者一覧】.xls」が開かれていません！" _
          & vbNewLine & "開いてからやり直して下さい。"
      End
    End If
   
    With UserForm登録            '登録画面を表示
      .Caption = "入会登録"      'キャプション名を｢入会登録｣に変更
      .TextBox1.Enabled = True   'テキストボックス1を入力可能に
      .Show
    End With
 
  'オプションボタン2（変更）をクリックした場合
  ElseIf OptionButton2.Value = True Then
    With UserForm検索       '検索画面を表示
      .Caption = "検索"     '名称を｢検索｣に
      .CommandButton3.Caption = "変更登録"   'コマンドボタン3の名称を｢変更登録｣に
      .CommandButton3.Enabled = True         'コマンドボタン3を使用可能に
      .Show
    End With
  
  'オプションボタン3（退会）をクリックした場合
  ElseIf OptionButton3.Value = True Then
   
    '処理に必要なファイルが開かれているか確認する
    If IsBookOpen("東京東筑会名簿【入退会者一覧】.xls") = False Then
       MsgBox "「東京東筑会名簿【入退会者一覧】.xls」が開かれていません！" _
          & vbNewLine & "開いてからやり直して下さい。"
       End
    End If
   
    With UserForm検索       '検索画面を表示
      .Caption = "検索"     '名称を｢検索｣に
      .CommandButton3.Caption = "退会登録"  'コマンドボタン3の名称を｢退会登録｣に
      .CommandButton3.Enabled = True        'コマンドボタン3を使用可能に
      .Show
    End With
  
  'オプションボタン4（住所一括照会）をクリックした場合
  ElseIf OptionButton4.Value = True Then
    Call ChkAllAddr
  End If

End Sub

'オプションボタン4（住所一括照会）を押されたときのプロシージャ
Private Sub ChkAllAddr()

  Dim msg As String
  Dim nrow As Long   '行1
  Dim zip1 As String '〒
  Dim zip2 As String '〒2
  Dim obj As Range   '検索1
  Dim zrow As Long   '行2
  Dim addrs(3) As String
   
  '処理に必要なファイルが開かれているか確認する
  If IsBookOpen("郵便番号ﾃﾞｰﾀ【全国版】.xls") = False Then
    MsgBox "「郵便番号ﾃﾞｰﾀ【全国版】.xls」が開かれていません！" _
       & vbNewLine & "開いてからやり直して下さい。"
    End
  End If
 
  msg = "選択された行から、郵便番号に基く、住所一括照合を行いますか？"  'このメッセージを表示し、
  If MsgBox(msg, 4 + 32, "〒⇒住所一括照合") = 6 Then  '｢はい｣をクリックすれば、
  
    nrow = ActiveCell.Row    'アクティブセルの行数を表す変数を｢行1｣とする。
   
    Do '以下の手順を繰り返す。
      zip1 = Left(Cells(nrow, COL_ZIP), 3) & Right(Cells(nrow, COL_ZIP), 4)   '選択されている行に記載されている郵便番号の｢-｣を削除した文字列を変数｢Zip1｣とする。
      zip2 = Cells(nrow, COL_ZIP)     '選択されている行に記載されている文字を表す変数を｢〒2｣とする。
    
      If zip2 = "－" Or zip2 = "" Then         '変数｢Zip2｣が｢－｣、または空白であれば、何もしない。
      Else    '変数｢Zip2｣が｢－｣、または空白でなければ、以下の処理を行う。
    
        '郵便番号を検索する。変数｢Obj｣は郵便番号が検索できたかどうかを表す。
        With Workbooks("郵便番号ﾃﾞｰﾀ【全国版】.xls").Sheets("郵便番号1").Range("B1:B65001")
          Set obj = .Find(zip1, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                    MatchCase:=False, MatchByte:=False)
        End With
   
        If Not obj Is Nothing Then      'シート｢郵便番号1｣で郵便番号が検索されれば、
          With Workbooks("郵便番号ﾃﾞｰﾀ【全国版】.xls").Sheets("郵便番号1")
            zrow = .Range(obj.Address).Row
            addrs(0) = .Cells(zrow, 3)
            addrs(1) = .Cells(zrow, 4)
            addrs(2) = .Cells(zrow, 5)
          End With
      
        Else                             'シート｢郵便番号1｣で郵便番号が検索されなければ、
          With Workbooks("郵便番号ﾃﾞｰﾀ【全国版】.xls").Sheets("郵便番号2").Range("B1:B65001")
            Set obj = .Find(zip1, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                      MatchCase:=False, MatchByte:=False)
          End With
      
          If obj Is Nothing Then     'シート｢郵便番号2｣で郵便番号が検索さなければ、該当行を選択し、以下のメッセージを表示し、照合を中断する。
            Cells(nrow, 7).Select
            MsgBox "該当の郵便番号は検索されませんでした！"
            Exit Do
      
          Else                         'シート｢郵便番号2｣で郵便番号が検索されれば、
            With Workbooks("郵便番号ﾃﾞｰﾀ【全国版】.xls").Sheets("郵便番号2")
              zrow = .Range(obj.Address).Row
              addrs(0) = .Cells(zrow, 3)
              addrs(1) = .Cells(zrow, 4)
              addrs(2) = .Cells(zrow, 5)
            End With
          End If
        End If
        
        'H列と変数｢住所1｣、I列と変数｢住所2｣、J列と変数｢住所3｣が等しければ、何もしない。
        If Cells(nrow, COL_ADDR1) = addrs(0) And Cells(nrow, COL_ADDR2) = addrs(1) _
            And Cells(nrow, COL_ADDR3) = addrs(2) Then
        
        Else  '等しくなければ、該当行を選択し、以下のメッセージを表示し、照合を中断する。
          Cells(nrow, COL_ZIP).Select
          MsgBox "住所が一致しません！『" & addrs(0) & addrs(1) & addrs(2) & "』"
          Exit Do
        End If
      End If
     
      nrow = nrow + 1
            
      If Cells(nrow, 1) = "" Then  'もし、A列が空白であれば、以下のメッセージを表示して、作業を中止する。
        Cells(nrow, COL_ZIP).Select
        MsgBox "照合が終了しました！"
        Exit Do
      Else
      End If
     
    Loop Until nrow = 5000   'ループを5000回繰り返す。
  End If
End Sub

Private Sub CommandButton2_Click()
  UserForm登録種類選択.Hide
End Sub


