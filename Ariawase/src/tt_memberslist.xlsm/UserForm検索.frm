VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm検索 
   Caption         =   "検索"
   ClientHeight    =   1800
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5268
   OleObjectBlob   =   "UserForm検索.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm検索"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'検索ボタンを押したときの動作
Private Sub 検索ボタン_Click()
 
  Dim ki As String      'ユーザフォームの「期」に入力された文字
  Dim txt As String     'ユーザフォームの「氏名等」に入力された文字
  Dim kil1 As String    'ユーザフォームの「期」に入力された文字の左から1文字目の文字
  Dim kir1 As String    'ユーザフォームの「期」に入力された文字の右から1文字目の文字
  Dim kir2 As String    'ユーザフォームの「期」に入力された文字の右から2文字目の文字
  Dim kilen As Integer  'ユーザフォームの「期」に入力された文字の文字数
  Dim trow As Long      '検索した先頭行
  Dim nrow As Long      '処理している行
  Dim obj As Range      '検索結果のオブジェクト
 
  '検索フォームの入力内容をセット
  ki = UserForm検索.TextBox期.Text
  txt = UserForm検索.TextBox氏名等.Text
  
  '「期」の入力が無い場合、シート全体から「氏名等」を検索
  If ki = "" Then
    With Sheets("名簿").Range("A1:IV65536")
      Set obj = .Find(txt, After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, _
                  SearchDirection:=xlNext, MatchCase:=False, MatchByte:=False)
    End With
 
    If obj Is Nothing Then    '検索されなければ、メッセージ表示
      MsgBox "該当の文字列は検索できませんでした！"
    Else                      '検索された場合は、そのセルを選択
      Range(obj.Address).Select
    End If
 
  '「期」の入力が有る場合、シート全体から「氏名等」を検索
  Else
    kil1 = Left(ki, 1)
    kir1 = Right(ki, 1)
    kir2 = Right(ki, 2)
    kilen = Len(ki)
    
    '「期」に卒業年を和暦を入力された場合、期に変換
    If kil1 = "S" Or kil1 = "s" Then        '昭和の場合は、年号に＋23で期を計算
      If kilen = 2 Then
        ki = kir1 + 23
      ElseIf kilen = 3 Then
        ki = kir2 + 23
      End If
    ElseIf kil1 = "H" Or kil1 = "h" Then    '平成の場合は、年号に+86で期を計算
      If kilen = 2 Then
        ki = kir1 + 86
      ElseIf kilen = 3 Then
        ki = kir2 + 86
      End If
    End If

    If Len(ki) = 2 Then
      ki = "0" & ki
    End If
   
    '「期」列を指定し検索
    With Sheets("名簿").Range("A1:A" & MEMBER_MAX)
      Set obj = .Find(ki, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                  SearchOrder:=xlByColumns, MatchCase:=False, MatchByte:=False)
    End With
 
    If obj Is Nothing Then       '該当の｢期｣が検索されない場合
      MsgBox "該当の｢期番号｣は検索できませんでした！"
    
    Else                         '該当の「期」が検索された場合
      If txt = "" Then           '「氏名等」が入力されていないとき、「期」の先頭セルを選択
        Range(obj.Address).Select
      Else
        trow = Range(obj.Address).Row      '「期」が検索された先頭行を取得
        nrow = trow                        '先頭行を処理する行の初期値にセット
     
        Do
          If Cells(nrow, 1) = Cells(nrow + 1, 1) Then   '１行下の行の「期」が同じ場合、次の行へ
            nrow = nrow + 1
          Else                                          '「期」が違う場合、繰り返し処理を中止
            Exit Do
          End If
        Loop Until nrow = DOUKI_MAX + 1       'DOUKI_MAX + 1 になるまで繰り返し
     
        '範囲を指定し、｢氏名等｣を検索する。その検索結果を変数｢G｣とする。
        With Sheets("名簿").Range(Range(Cells(trow, COL_NAME), Cells(nrow, COL_NAME)), Range(Cells(trow, COL_COMMENT), Cells(nrow, COL_COMMENT)))
     
          '2回目の検索で、検索条件の期が変更になっているなど、現在のセル位置が検索範囲外の場合は1回目に修正
          If ActiveCell.Row < trow Or ActiveCell.Row > nrow Or ActiveCell.Column < COL_NAME Or ActiveCell.Column > COL_COMMENT Then
            UserForm検索.Label5.Caption = 1
          End If
     
          'UserForm検索.Label5（隠しラベル）の数が｢1｣の場合、上記範囲の最上行から検索
          If UserForm検索.Label5.Caption = 1 Then
            Set obj = .Find(txt, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, _
                       SearchDirection:=xlNext, MatchCase:=False, MatchByte:=False)
          
          'そうでなければ、上記範囲の選択された行以降で検索
          Else
            
            Set obj = .Find(txt, After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, _
                     SearchDirection:=xlNext, MatchCase:=False, MatchByte:=False)
          End If
        End With
          
        If Not obj Is Nothing Then      '「氏名等」が検索された場合、当該セルを選択
          Range(obj.Address).Select
        Else                            '「氏名等」が検索されない場合、メッセージ表示
          MsgBox "該当の｢期・氏名｣の組み合わせでは検索できませんでした！"
        End If
      End If
    End If
  End If
   
  '検索ボタンをクリックする毎にカウンターを１増やす
  UserForm検索.Label5.Caption = UserForm検索.Label5.Caption + 1
 
End Sub


Private Sub CommandButton2_Click()    '｢閉じる｣ボタンをクリックする。
  With UserForm検索
    .TextBox期.Text = ""       'ユーザーフォーム検索の「期」欄の内容を空白にする。
    .TextBox氏名等.Text = ""   'ユーザーフォーム検索の「氏名等」欄の内容を空白にする。
    .Label5.Caption = 1        'ユーザーフォーム検索の「カウンター」欄の内容を「1」にする。
    .Hide                      '検索画面を閉じる。
  End With

End Sub

'検索後に検索結果を「変更登録」または「退会登録」する処理
Private Sub CommandButton3_Click()
   
  Dim nrow As Long     '処理する行
  Dim msg As String
  Dim lrow As Long     '入退会者一覧ファイルの最終行
  Dim i As Integer
   
  Select Case CommandButton3.Caption  'コマンドボタン3の名称で場合分け
    Case "変更登録"  '｢変更登録｣の場合
       
      Unload UserForm検索
      With UserForm登録
        .Caption = "変更登録"       '名称を｢変更登録｣に
        .TextBox1.Enabled = False   'テキストボックス1を入力不可に
        .Show
      End With

   
    Case "退会登録"   '｢退会登録｣の場合
      nrow = ActiveCell.Row     '選択されている行数を取得
      Range(Cells(nrow, COL_KI), Cells(nrow, COL_KI)).EntireRow.Select
    
      msg = "この行を退会登録にしますか？"
      If MsgBox(msg, 4 + 32, "退会") = 6 Then    '｢はい｣をクリックした場合
        Range(Cells(nrow, COL_KI), Cells(nrow, COL_REMARK)).Copy   '「期」列から「備考」列までをコピー
     
        '｢東京東筑会名簿【入退会者一覧】.xls｣を開き、次行に貼り付け
        With Workbooks("東京東筑会名簿【入退会者一覧】.xls").Sheets("退会者")
          lrow = .Range("A" & MEMBER_MAX).End(xlUp).Row + 1
          .Cells(lrow, 1) = Date
          .Cells(lrow, 2).PasteSpecial Paste:=xlValues, _
              Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End With
  
        '「名前」列から「メールアドレス」列までに｢－｣を入力し中央揃え
        '「年会費情報」と「コメント欄」を残してクリア←←変更2011
        For i = COL_NAME To COL_EMAIL
          Cells(nrow, i) = "－"
        Next i
        Range(Cells(nrow, COL_NAME), Cells(nrow, COL_EMAIL)).HorizontalAlignment = xlCenter
        Range(Cells(nrow, COL_BUKATSU), Cells(nrow, COL_COUPLE)).ClearContents
        Range(Cells(nrow, COL_KANJI), Cells(nrow, COL_KIPAY)).ClearContents
        
        Cells(nrow, COL_REMARK) = "退会のため欠番"  '備考列に｢退会のため欠番｣を入力
      End If
      Unload UserForm検索
      
  End Select
 
End Sub



