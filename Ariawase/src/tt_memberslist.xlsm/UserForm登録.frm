VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm登録 
   Caption         =   "登録"
   ClientHeight    =   3660
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10020
   OleObjectBlob   =   "UserForm登録.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm登録"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 登録画面の｢〒⇒住所｣ボタンが押された場合の動作を記述
Private Sub CommandButton1_Click()
   
  Dim zip As String        '郵便番号
  Dim abook As String      '郵便番号データファイル名
  Dim obj As Range         '検索結果
  Dim nrow As Long         '現在行
  Dim addrs(3) As String   '住所情報文字配列
   
  '郵便番号の文字列をセット
  zip = UserForm登録.TextBox4.Text & UserForm登録.TextBox5.Text
  
  '== Mar.15,2014 Yuji Ogihara ==
  'ファイル"郵便番号ﾃﾞｰﾀ【全国版】"についての変更
  '  1.  Excel 2007 形式(xlsx)への保存形式変更
  '  2. 元データからのファイル作成手順簡素化に伴うデータフィールドの変更
       
  
  '== Mar.15,2014 Yuji Ogihara ==
  ' xls -> xlsx
  'abook = "郵便番号ﾃﾞｰﾀ【全国版】.xls"
  abook = "郵便番号ﾃﾞｰﾀ【全国版】.xlsx"
  
       
  '郵便番号データファイルが開いているか確認
   If IsBookOpen(abook) = False Then
     MsgBox "郵便番号データファイル『" & abook & "』が開かれていません！" _
               & vbNewLine & "開いてからやり直して下さい。"
     End
   End If
  
  '郵便番号を検索
  '== Mar.15,2014 Yuji Ogihara ==
  '郵便番号の列を"C"に、範囲を「C列全体」に
  'With Workbooks(abook).Sheets("郵便番号1").Range("B1:B65001")
  With Workbooks(abook).Sheets("郵便番号1").Range("C:C")
    Set obj = .Find(zip, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                   MatchCase:=False, MatchByte:=False)
  End With
   
  If Not obj Is Nothing Then     'シート｢郵便番号1｣で郵便番号が検索さた場合
    With Workbooks(abook).Sheets("郵便番号1")
      nrow = .Range(obj.Address).Row
      
     '== Mar.15,2014 Yuji Ogihara ==
     ' 住所の列番号を7-9 に変更
     'addrs(0) = .Cells(nrow, 3)
     'addrs(1) = .Cells(nrow, 4)
     'addrs(2) = .Cells(nrow, 5)
      addrs(0) = .Cells(nrow, 7)
      addrs(1) = .Cells(nrow, 8)
      addrs(2) = .Cells(nrow, 9)
    End With
      
  Else                           'シート｢郵便番号1｣で郵便番号が検索されない場合
     '== Mar.15,2014 Yuji Ogihara ==
     ' シート「郵便番号2」廃止
    'With Workbooks(abook).Sheets("郵便番号2").Range("B1:B65001")
    '  Set obj = .Find(zip, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
    '                  MatchCase:=False, MatchByte:=False)
    'End With
      
    'If obj Is Nothing Then     'シート｢クエリ2｣で郵便番号が検索さない場合
      MsgBox "該当の郵便番号は検出されませんでした！"
      End
             
    'Else                         'シート｢クエリ2｣で郵便番号が検索された場合
    '  With Workbooks("郵便番号ﾃﾞｰﾀ【全国版】.xls").Sheets("郵便番号2")
    '    nrow = .Range(obj.Address).Row
    '    addrs(0) = .Cells(nrow, 3)
    '    addrs(1) = .Cells(nrow, 4)
    '    addrs(2) = .Cells(nrow, 5)
    '  End With
    'End If
  End If
     
  'テキストボックス6～8に検索結果を入力
  UserForm登録.TextBox6.Text = addrs(0)
  UserForm登録.TextBox7.Text = addrs(1)
  UserForm登録.TextBox8.Text = addrs(2)
   
End Sub

' 登録画面の｢登録｣ボタンが押された場合の動作を記述
Private Sub CommandButton2_Click()

  Dim obj As Range           '検索結果のオブジェクト
  Dim nrow As Long           '処理している行
  Dim msg As String
 
  Select Case UserForm登録.Caption
    Case "入会登録"    'ユーザーフォーム登録の名称が｢入会登録｣の場合｡

      Call GetKiTopRow(nrow)   '登録する期の先頭行を取得
      
      Do
        If Cells(nrow, COL_KI) = Cells(nrow + 1, COL_KI) Then   '１行下の行の期が同じ
          nrow = nrow + 1
        
        Else    '期番号が違えば、その期の最終行を選択し、繰り返し処理を中止する。
          Range(Cells(nrow, COL_KI), Cells(nrow, COL_KI)).EntireRow.Select
          msg = "この行の下に１行追加して、入会登録しますか？"
          
          If MsgBox(msg, 4 + 32, "入会") = 6 Then   '｢はい｣をクリックされた場合
            Application.Calculation = xlManual      '再計算を停止
            nrow = nrow + 1
            Call InsertNewRow(nrow)      '当該期の最終行に1行追加
            Call RegNewData(nrow)        '追加した行に新規データをセット
            Call RegNewDataToList(nrow)  '「入退会一覧」に新規データを登録
       
            Application.Calculation = xlAutomatic  '再計算を再開
            Cells(nrow, COL_NAME).Select
          End If
          Exit Do
        End If
      
      Loop Until nrow = DOUKI_MAX  'DOUKI_MAXになるまで繰り返す
      Unload UserForm登録
   
    Case "変更登録"    'ユーザーフォーム登録の名称が｢変更登録｣の場合
 
      msg = "この行に変更分を登録しますか？"
      If MsgBox(msg, 4 + 32, "変更") = 6 Then  '｢はい｣の場合
        Application.Calculation = xlManual     '再計算を停止
        nrow = ActiveCell.Row
   
        Call RegEditData(nrow)       '修正データを登録
     
        Application.Calculation = xlAutomatic  '再計算を再開
   
      End If
      Unload UserForm登録
      
  End Select

End Sub

'登録フォームに入力された「期」の先頭行を取得するプロシージャ
Sub GetKiTopRow(trow As Long)

  Dim ki As String       'ユーザフォームの「期」に入力された文字
  Dim obj As Range       '検索結果のセル
  Dim i As Integer
  
  ki = UserForm登録.TextBox1.Text
    
  '範囲を指定し｢期｣を検索
  With Sheets("名簿").Range("A1:A" & MEMBER_MAX)
    Set obj = .Find(ki, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                   SearchOrder:=xlByColumns, MatchCase:=True, MatchByte:=False)
  End With
 
  If Not obj Is Nothing Then  '「期」が検索された場合
    trow = Range(obj.Address).Row
    
  Else       '「期」が検索されない場合、「期」より小さい最大の「期」を検索
    
    i = 1
    Do
      
      '範囲を指定し、「期」を検索
      With Sheets("名簿").Range("A1:A" & MEMBER_MAX)
        Set obj = .Find(ki - i, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                       SearchOrder:=xlByColumns, MatchCase:=True, MatchByte:=False)
      End With
        
      If Not obj Is Nothing Then    '「期」が検索された場合
        trow = Range(obj.Address).Row
        Exit Do
      Else                          '「期」が検索されない場合
        i = i + 1    '1つ小さい「期」を検索
      End If
    Loop Until i = 30    '30小さい期まで検索
  End If
End Sub


'当該の期の最終行に新規データを登録するために1行追加するプロシージャ
Sub InsertNewRow(nrow As Long)

  Selection.EntireRow.Insert    '１行追加
                                               
  '該当期最終行の１行前に行追加されるので、該当期最終行を１行前にコピ
  '～入金集計を｢当番期｣、｢その他｣に分けて集計するためには、｢当番期｣または｢当番期の前の期｣で
  '　入会登録をする場合、最終行の１行前に行追加をしないと関数｢COUNIF｣の範囲(行数)が
  '　正しく更新されない。
           
  Range(Cells(nrow, COL_KI), Cells(nrow, COL_COMMENT)).Copy
  Cells(nrow - 1, COL_KI).PasteSpecial Paste:=xlValues, _
              Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False  'コピーモードをキャンセル
  Range(Cells(nrow, COL_KI), Cells(nrow, COL_COMMENT)).ClearContents     '最終行をクリア
         
  Application.Calculation = xlManual '再計算を停止
            
End Sub


'追加した行に新規データを登録するプロシージャ
Sub RegNewData(nrow As Long)

  Dim data As String       'ユーザフォームに入力された登録データ

  With UserForm登録
             
    '「期」が2桁の場合は先頭に0を付加し3桁に
    data = .TextBox1.Text
    If Len(data) = 2 Then
       data = "0" & data
    End If
    Cells(nrow, COL_KI) = "'" & data
            
    '期がJで始まる場合は、「種類」を1、そうでなければ2
    If Left(data, 1) = "J" Then
      Cells(nrow, COL_CLASS) = 1
    Else
      Cells(nrow, COL_CLASS) = 2
    End If
                    
    '追加する行の「期」と最終行の「期」が同じ場合、単純に「ID」を+1
    If Cells(nrow, COL_KI) = Cells(nrow - 1, COL_KI) Then
      data = FormatNumber(Right(Cells(nrow - 1, COL_ID), 3)) + 1
    Else   '違う場合、新しい「期」の追加と判断し「ID」を初期化
      data = "001"
    End If
    
    '「ID」を6桁数字に整形←2011年変更―新しい期のIDが正しくできるように
    If Len(data) = 1 Then
      Cells(nrow, COL_ID) = Cells(nrow, COL_KI).Text & "00" & data
    ElseIf Len(data) = 2 Then
      Cells(nrow, COL_ID) = Cells(nrow, COL_KI).Text & "0" & data
    Else
      Cells(nrow, COL_ID) = Cells(nrow, COL_KI).Text & data
    End If
            
    Cells(nrow, COL_NAME) = .TextBox2.Text   'ﾃｷｽﾄﾎﾞｯｸｽ2を｢氏名｣に
    
    Cells(nrow, COL_KANA) = StrConv(.TextBox15.Text, vbNarrow)    'ﾃｷｽﾄﾎﾞｯｸｽ15を｢カナ氏名｣に
          
    If OptionButton1.Value = True Then       '｢性別｣が｢男｣の場合
      Cells(nrow, COL_SEX) = OptionButton1.Caption
    ElseIf OptionButton2.Value = True Then   '｢性別｣が｢女｣の場合
      Cells(nrow, COL_SEX) = OptionButton2.Caption
    End If
    
    Cells(nrow, COL_ZIP) = .TextBox4.Text & "-" & .TextBox5.Text   'ﾃｷｽﾄﾎﾞｯｸｽ4～5を「〒」に
    Cells(nrow, COL_ADDR1) = .TextBox6.Text  'ﾃｷｽﾄﾎﾞｯｸｽ6を｢住所1｣に
    Cells(nrow, COL_ADDR2) = .TextBox7.Text  'ﾃｷｽﾄﾎﾞｯｸｽ7を｢住所2｣に
    Cells(nrow, COL_ADDR3) = .TextBox8.Text  'ﾃｷｽﾄﾎﾞｯｸｽ8を｢住所3｣に
    Cells(nrow, COL_ADDR4) = StrConv(.TextBox9.Text, vbNarrow)  'ﾃｷｽﾄﾎﾞｯｸｽ9を｢住所4｣に
    
    'ﾃｷｽﾄﾎﾞｯｸｽ10～12を｢電話番号｣に
    If .TextBox10.Text <> "" Then
      Cells(nrow, COL_TELNO) = StrConv(.TextBox10.Text & "-" & .TextBox11.Text & "-" & .TextBox12.Text, vbNarrow)
    End If
    
    'ﾃｷｽﾄﾎﾞｯｸｽ13～14を｢メールアドレス｣に
    If .TextBox13.Text <> "" Then
      Cells(nrow, COL_EMAIL) = StrConv(.TextBox13.Text & "@" & .TextBox14.Text, vbNarrow)
      Cells(nrow, COL_EMAIL).Font.Size = 9
      Cells(nrow, COL_EMAIL).Font.Underline = xlUnderlineStyleNone
    End If

    'ﾃｷｽﾄﾎﾞｯｸｽ16を整形して｢部活｣に
    Cells(nrow, COL_BUKATSU) = Replace(.TextBox16.Text, "部", "")
              
    'ﾃｷｽﾄﾎﾞｯｸｽ17を整形して｢出身中学｣に
    Call SetJHScool(nrow)
  
  End With
End Sub


'新規登録したデータを【入退会者一覧】ファイルにも登録するプロシージャ
Sub RegNewDataToList(nrow As Long)
            
  Dim lrow As Long           '最終行
            
  Range(Cells(nrow, COL_KI), Cells(nrow, COL_REMARK)).Copy
  With Workbooks("東京東筑会名簿【入退会者一覧】.xls").Sheets("入会者")
    lrow = .Range("A" & MEMBER_MAX).End(xlUp).Row + 1
    .Cells(lrow, 1) = Date
    .Cells(lrow, 2).PasteSpecial Paste:=xlValues, _
                       Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  End With
  
  Application.CutCopyMode = False        'コピーモードをキャンセル

End Sub


'修正登録したデータを登録するプロシージャ
Sub RegEditData(nrow As Long)
        
  With UserForm登録
    If .TextBox2.Text <> "" Then
      Cells(nrow, COL_NAME) = .TextBox2.Text  'ﾃｷｽﾄﾎﾞｯｸｽ2を｢氏名｣に
    End If
   
    If .TextBox15.Text <> "" Then
      Cells(nrow, COL_KANA) = StrConv(.TextBox15.Text, vbNarrow)  'ﾃｷｽﾄﾎﾞｯｸｽ15を｢カナ氏名｣に
    End If
   
    If OptionButton1.Value = True Then      '｢性別｣の｢男｣の場合
      Cells(nrow, COL_SEX) = OptionButton1.Caption
    ElseIf OptionButton2.Value = True Then  '｢性別｣の｢女｣の場合
      Cells(nrow, COL_SEX) = OptionButton2.Caption
    End If
   
    If .TextBox4.Text <> "" And .TextBox5.Text <> "" Then
      Cells(nrow, COL_ZIP) = .TextBox4.Text & "-" & .TextBox5.Text   'ﾃｷｽﾄﾎﾞｯｸｽ4～5を｢〒｣に
    End If
   
    If .TextBox6.Text <> "" Then
      Cells(nrow, COL_ADDR1) = .TextBox6.Text  'ﾃｷｽﾄﾎﾞｯｸｽ6を｢住所1｣に
    End If
     
    If .TextBox7.Text <> "" Then
      Cells(nrow, COL_ADDR2) = .TextBox7.Text  'ﾃｷｽﾄﾎﾞｯｸｽ7を｢住所2｣に
    End If
   
    If .TextBox8.Text <> "" Then
      Cells(nrow, COL_ADDR3) = .TextBox8.Text  'ﾃｷｽﾄﾎﾞｯｸｽ8を｢住所3｣に
    End If
   
    If .TextBox9.Text <> "" Then
      Cells(nrow, COL_ADDR4) = StrConv(.TextBox9.Text, vbNarrow)  'ﾃｷｽﾄﾎﾞｯｸｽ9を｢住所4｣に
    End If
   
    If .TextBox10.Text <> "" And .TextBox11.Text <> "" And .TextBox12.Text <> "" Then
      Cells(nrow, COL_TELNO) = StrConv(.TextBox10.Text & "-" & .TextBox11.Text & "-" & .TextBox12.Text, vbNarrow)  'ﾃｷｽﾄﾎﾞｯｸｽ10～12を｢電話番号｣に
    End If
   
    If .TextBox13.Text <> "" And .TextBox14.Text <> "" Then
      Cells(nrow, COL_EMAIL) = StrConv(.TextBox13.Text & "@" & .TextBox14.Text, vbNarrow)  'ﾃｷｽﾄﾎﾞｯｸｽ13～14を｢メールアドレス｣に
    End If
      
    'ﾃｷｽﾄﾎﾞｯｸｽ16を整形して｢部活｣に
    Cells(nrow, COL_BUKATSU) = Replace(.TextBox16.Text, "部", "")
              
    'ﾃｷｽﾄﾎﾞｯｸｽ17を整形して｢出身中学｣に
    Call SetJHScool(nrow)
  End With
End Sub


'中学校のデータを整形して登録するプロシージャ
Sub SetJHScool(nrow As Long)

  Dim data As String       '中学校のデータ
    
  data = TextBox17.Text
  If data = "" Then
  ElseIf Right(data, 3) = "中学校" Then
    Cells(nrow, COL_JHSCHOOL) = Replace(data, "中学校", "中")
  ElseIf Right(data, 2) = "中学" Then
    Cells(nrow, COL_JHSCHOOL) = Replace(data, "中学", "中")
  ElseIf Right(data, 1) = "中" Then
    Cells(nrow, COL_JHSCHOOL) = data
  Else
    Cells(nrow, COL_JHSCHOOL) = data & "中"
  End If

End Sub


' 登録画面の｢閉じる｣ボタンが押された場合の動作を記述
Private Sub CommandButton3_Click()

  With UserForm登録
    .TextBox1.Text = ""
    .TextBox2.Text = ""
    .TextBox4.Text = ""
    .TextBox5.Text = ""
    .TextBox6.Text = ""
    .TextBox7.Text = ""
    .TextBox8.Text = ""
    .TextBox9.Text = ""
    .TextBox10.Text = ""
    .TextBox11.Text = ""
    .TextBox12.Text = ""
    .TextBox13.Text = ""
    .TextBox14.Text = ""
    .TextBox15.Text = ""
    .TextBox16.Text = ""
    .TextBox17.Text = ""
    .OptionButton1.Value = False
    .OptionButton2.Value = False
    .Hide
  End With
  
End Sub



