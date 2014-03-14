Attribute VB_Name = "Module1"
Option Explicit

'以下の定数は当番期交代時に見直しを行う定数
Const MEIBO_PASSWORD As String = "touchiku84"   '集客担当毎に分割した名簿の読み取りパスワード


'以下の定数はモジュール全体で使う定数
'名簿の列を追加した場合は、列の位置を示す定数が変わるため、検証が必要になるので要注意！

'--- 基本的な定数
Public Const MEMBER_MAX As Integer = 10000      '会員の最大数
Public Const DOUKI_MAX As Integer = 300         '同期会員の最大数

'--- 名簿シートに関する定数（シートの行や列の追加削除を行った場合は見直し要！）
Public Const ROW_TOPTITLE As Integer = 2 '名簿シートのタイトルを含む先頭行
Public Const ROW_TOPDATA As Integer = 5  '名簿シートのデータの先頭行
'
Public Const COL_KI As Integer = 1           '「期」の列
Public Const COL_CLASS As Integer = 2        '「種類」の列
Public Const COL_ID As Integer = 3           '「ID」の列
Public Const COL_NAME As Integer = 4         '「氏名」の列
Public Const COL_KANA As Integer = 5         '「カナ氏名」の列
Public Const COL_SEX As Integer = 6          '「性別」の列
Public Const COL_ZIP As Integer = 7          '「〒」の列
Public Const COL_ADDR1 As Integer = 8        '「住所1」の列
Public Const COL_ADDR2 As Integer = 9        '「住所2」の列
Public Const COL_ADDR3 As Integer = 10       '「住所3」の列
Public Const COL_ADDR4 As Integer = 11       '「住所4」の列
Public Const COL_TELNO As Integer = 12       '「電話番号」の列
Public Const COL_EMAIL As Integer = 13       '「メールアドレス」の列
Public Const COL_BUKATSU As Integer = 14     '「部活動活」の列
Public Const COL_JHSCHOOL As Integer = 15    '「出身中学」の列
Public Const COL_COUPLE As Integer = 16      '「夫婦」の列
Public Const COL_KAIHI3 As Integer = 17      '「年会費3年前」の列
Public Const COL_KAIHI2 As Integer = 18      '「年会費2年前」の列
Public Const COL_KAIHI1 As Integer = 19      '「年会費1年前」の列
Public Const COL_KAIHI0 As Integer = 20      '「年会費（今年度）」の列
Public Const COL_KANJI As Integer = 21       '「幹事等」の列
Public Const COL_REMARK As Integer = 22      '「備考」の列
Public Const COL_CHECK As Integer = 23       '「転記チェック欄」の列

Public Const COL_RSLT5 As Integer = 24       '「懇親会実績5年前」の列
Public Const COL_RSLT4 As Integer = 25       '「懇親会実績4年前」の列
Public Const COL_RSLT3 As Integer = 26       '「懇親会実績3年前」の列
Public Const COL_RSLT2 As Integer = 27       '「懇親会実績2年前」の列
Public Const COL_RSLT1 As Integer = 28       '「懇親会実績1年前」の列
Public Const COL_RSLT0 As Integer = 29       '「懇親会実績（今年度）」の列

Public Const COL_CARD As Integer = 30        '「（今年度懇親会の出欠）返信」の列
Public Const COL_TEL As Integer = 31         '「（今年度懇親会の出欠）電話」の列
Public Const COL_KICARD As Integer = 32      '「（今年度懇親会の出欠）期別返信」の列
Public Const COL_KITEL As Integer = 33       '「（今年度懇親会の出欠）期別電話」の列
Public Const COL_PLAN As Integer = 34        '「（今年度懇親会の出欠）出予定」の列
Public Const COL_RSLT As Integer = 35        '「（今年度懇親会の出欠）当日」の列
Public Const COL_KIRSLT As Integer = 36      '「（今年度懇親会の出欠）期別当日」の列
Public Const COL_ADVPAY As Integer = 37      '「（今年度懇親会の入金）事前」の列
Public Const COL_PAY As Integer = 38         '「（今年度懇親会の入金）当日」の列
Public Const COL_KIPAY As Integer = 39       '「（今年度懇親会の入金）期別入金」の列
Public Const COL_COMMENT As Integer = 40        '「コメント」の列

'--- 期別出欠シート・期別入金シート共通の定数（シートの行や列の追加削除を行った場合は見直し要！）
Const ROW_KITOPDATA As Integer = 6    'データの先頭行数
Const COL_TANTO_MAX As Integer = 100       '集客担当の最大値
Const KI_MAX As Integer = 200         '期の最大値


'
'--- 出欠シートに関する定数（シートの行や列の追加削除を行った場合は見直し要！）
Const ROW_RSLTSUM1 As Integer = 4       '合計（上）の行数
Const ROW_RSLTSUM2 As Integer = 5       '合計（下）の行数

'
Const COL_CARD_OK As Integer = 2      '「返信ハガキの出」の列
Const COL_CARD_NG As Integer = 3      '「返信ハガキの欠」の列
Const COL_CARD_ERROR As Integer = 4   '「返信ハガキの不着」の列
Const COL_CARD_NOFIX As Integer = 5   '「返信ハガキの未定」の列
Const COL_CARD_SUM As Integer = 6     '「返信ハガキの計」の列
Const COL_EMAIL_OK As Integer = 7     '「メール・HPの出」の列
Const COL_EMAIL_NG As Integer = 8     '「メール・HPの欠」の列
Const COL_EMAIL_NOFIX As Integer = 9  '「メール・HPの未定」の列
Const COL_EMAIL_SUM As Integer = 10   '「メール・HPの計」の列
Const COL_TEL_OK As Integer = 11      '「電話の出」の列
Const COL_TEL_NG As Integer = 12      '「電話の欠」の列
Const COL_TEL_ERROR As Integer = 13   '「電話の不通」の列
Const COL_TEL_NOFIX As Integer = 14   '「電話の未定」の列
Const COL_TEL_SUM As Integer = 15     '「電話の計」の列
Const COL_SUM_OK As Integer = 16      '「出席連絡」の列

Const COL_OKPAY As Integer = 18       '「出席かつ事前入金」の列
Const COL_OTHERPAY As Integer = 19    '「出席以外かつ事前入金」の列
Const COL_SUMPAY As Integer = 20      '「事前入金の計」の列
  
Const COL_SUM_ALL As Integer = 22     '「出席予定」の列

Const COL_RSLT_OK As Integer = 24     '「当日出席」の列
Const COL_RSLT_PAYNG As Integer = 25  '「当日入金欠席」の列
Const COL_RSLT_SUM As Integer = 26    '「当日出席」の列

Const COL_MEMBER As Integer = 28      '「会員数」の列
Const COL_TANTO As Integer = 30       '「集客担当」の列

'== Mar.9,2014 Yuji Ogihara ==
'2007 マクロ有効形式への変換およびファイル名称の定数化
Const SAGYOHYO_FILENAME As String = "宛名シール・名札・出席者一覧表作成.xlsm" '作業表ファイル名称

'
'--- モジュール全体で使用する変数
'    （集客担当等が修正したデータを転記するために使用する）
Public OrgBook As String            '原本のファイル名
Public OrgBasicData As String       '原本の基本データ
Public OrgRow As Long               '修正候補の原本シートでの行数
Public NewBasicData As String       '修正候補の基本データ
Public PostFlag As Integer          '基本データ修正を行うか否かのフラグ



'


'ファイルが開かれているかを調べる関数
' 引数　　bName ：開かれているか調べるBOOK名
' 戻り値　True  ：開かれている、　False：開かれていない
Public Function IsBookOpen(bname As String) As Boolean

  Dim wbook As Workbook

  IsBookOpen = False

  For Each wbook In Workbooks
    If wbook.Name = bname Then
      IsBookOpen = True
      Exit For
    End If
  Next

End Function



'「名簿」シートの「検索」ボタンが押されたときの処理
Sub 検索画面表示()
    
   With UserForm検索
     .CommandButton3.Enabled = False
     .Show
   End With
   
End Sub

'「名簿」シートの「登録」ボタンが押されたときの処理
'（入会・変更・退会・〒→住所照合のサブメニューを表示）
Sub 登録種類選択画面表示()

    UserForm登録種類選択.Show
End Sub

'「名簿」シートの「凡例表示」ボタンが押されたときの処理
Sub 凡例画面表示()

    UserForm凡例.Show
End Sub

'「名簿」シートの「転記」ボタンが押されたときの処理
'（年会費、懇親会への申込み状況等の情報を転記するための処理・ファイルの指定）
Sub 転記画面表示()

    UserForm転記.Show
End Sub

Sub 表記確認()

 Dim nrow As Long
 Dim lrow As Long
 Dim i As Integer
 
 lrow = Cells(MEMBER_MAX, COL_KI).End(xlUp).Row

 For nrow = ROW_TOPDATA To lrow

   If Cells(nrow, COL_NAME) = "－" Or Cells(nrow, COL_NAME) = "-" Or Cells(nrow, COL_NAME) = "" Then
    
     For i = COL_NAME To COL_EMAIL
       Cells(nrow, i) = "－"
     Next i
     Range(Cells(nrow, COL_NAME), Cells(nrow, COL_EMAIL)).HorizontalAlignment = xlCenter
     Range(Cells(nrow, COL_BUKATSU), Cells(nrow, COL_COMMENT)).ClearContents
   End If
 
   If Cells(nrow, COL_COUPLE) <> "" Then
     Range(Cells(nrow, COL_KI), Cells(nrow, COL_COMMENT)).Interior.ColorIndex = 43
   End If
   
 
 Next nrow
    
End Sub

'期別の懇親会出欠状況を集計するための情報を、隠し列にセットするマクロ
' 出予：ハガキ、メール、電話で出席と回答があった場合
' 見込：ハガキ、メール、電話では欠席または回答が無いが、参加費の入金があった場合
Sub 出欠確認()
     
  Dim nrow As Long     '処理を行う行
  Dim lrow As Long     '最終行
  Dim ki As String     '処理する期の文字列
  Dim card As String   '「返信」列の文字列
  Dim tel As String    '「電話」列の文字列
  Dim rslt As String   '「当日」列の文字列
  Dim apay As String   '「事前入金」列の文字列
  Dim per As Integer   '処理割合
     
  Application.ScreenUpdating = False      '描画を中止
  ActiveSheet.EnableCalculation = False   '再計算を中止

  With Sheets("名簿")
 
    nrow = ROW_TOPDATA   '処理開始行をセット
   
    '最終行を「期」列のMEMBER_MAX行から上に調査
    lrow = .Cells(MEMBER_MAX, COL_KI).End(xlUp).Row
   
    '隠し列の期別返信、期別電話、出予定、期別当日出席をクリア
    .Range(.Cells(nrow, COL_KICARD), .Cells(lrow, COL_PLAN)).ClearContents
    .Range(.Cells(nrow, COL_KIRSLT), .Cells(lrow, COL_KIRSLT)).ClearContents

    Do  '以下の処理を繰り返す
     
      per = (nrow - ROW_TOPDATA) / (lrow - ROW_TOPDATA) * 100
      Application.StatusBar = "◆" & nrow - ROW_TOPDATA + 1 & "行目 (" & per & " %) の処理中．．．"
 
      '現在の行の期、返信、電話、当日、事前入金の文字を取得
      ki = .Cells(nrow, COL_KI).Text
      card = .Cells(nrow, COL_CARD).Text
      tel = .Cells(nrow, COL_TEL).Text
      rslt = .Cells(nrow, COL_RSLT).Text
      apay = .Cells(nrow, COL_ADVPAY).Text
    
     '｢期｣の文字列が空白の場合は処理を中止
      If ki = "" Then
        Exit Do
     
      Else
      
        '｢返信｣の文字列が空白でなければ、「期別返信」列に「期」+「返信」の文字をセット
        If card <> "" Then
          .Cells(nrow, COL_KICARD) = ki & card
          '｢返信｣の1文字目が｢出｣の場合は、「出予定」列に「期」+"出予"をセット
          If Left(card, 1) = "出" Then
            .Cells(nrow, COL_PLAN) = ki & "出予"
          End If
        
        End If
   
        '「電話」の文字列が空白でなければ、「期別電話」列に｢期｣+｢電話｣の文字をセット
        If tel <> "" Then
          .Cells(nrow, COL_KITEL) = ki & tel
          '｢電話｣が｢出｣の場合は、「出予定」列に「期」+"出予"をセット
          If Left(card, 1) = "出" Or Left(tel, 1) = "出" Then
            .Cells(nrow, COL_PLAN) = ki & "出予"
          End If
        End If

        '「出予定」が空白で「事前入金」が空白でなければ、「出予定」列に｢期｣+"見込"の文字をセット
        If .Cells(nrow, COL_PLAN) = "" And apay <> "" Then
          .Cells(nrow, COL_PLAN) = ki & "見込"
        End If

        '｢当日｣の文字列が空白でなければ、｢期｣+｢当日｣の文字をセット
        If rslt <> "" Then
          .Cells(nrow, COL_KIRSLT) = ki & rslt
        ElseIf apay <> "" Then
          .Cells(nrow, COL_KIRSLT) = ki & "入欠"
        End If
      End If
 
      nrow = nrow + 1  '次の行へ
     
    Loop Until nrow = lrow + 1  '最終行+1になるまで繰り返し
    
    Application.StatusBar = False   'ステータスバーの消去
   
  End With
   
  ActiveSheet.EnableCalculation = True   '再計算を再開
  Application.ScreenUpdating = True      '描画を再開
   
End Sub


'期別の懇親会参加費入金状況を集計するための情報を、隠し列にセットするマクロ
' XX事前をセット
Sub 期別入金集計()
     
  Dim nrow As Long      '処理を行う行
  Dim lrow As Long      '最終行
  Dim ki As String      '処理する期の文字列
  Dim apay As String    '「事前入金」列の文字列
  Dim pay As String     '「当日入金」列の文字列
  Dim per As Integer    '処理割合
     
  Application.ScreenUpdating = False      '描画を中止
  ActiveSheet.EnableCalculation = False   '再計算を中止
  
  With Sheets("名簿")
    nrow = ROW_TOPDATA
    
    '最終行を「期」列のMEMBER_MAX行から上に調査
    lrow = .Cells(MEMBER_MAX, COL_KI).End(xlUp).Row
    
    '期別入金の列を削除
    .Range(.Cells(nrow, COL_KIPAY), .Cells(lrow, COL_KIPAY)).ClearContents
  
    Do  '以下の処理を繰り返す。
      
      per = (nrow - ROW_TOPDATA) / (lrow - ROW_TOPDATA) * 100
      Application.StatusBar = "◆" & nrow - ROW_TOPDATA + 1 & "行目 (" & per & " %) の処理中．．．"
      
      '現在の行の期、事前入金の文字、当日入金の文字を取得
      ki = .Cells(nrow, COL_KI).Text
      apay = .Cells(nrow, COL_ADVPAY).Text
      pay = .Cells(nrow, COL_PAY).Text
        
      '｢期｣の文字列が空白の場合は処理を中止
      If ki = "" Then
        Exit Do
      Else
        '｢事前入金｣の文字列が空白でなければ、｢期｣と｢事前入金｣の文字列を加えたものをセット
        If apay <> "" Then
          .Cells(nrow, COL_KIPAY) = ki & apay
        End If
        
        '｢当日当日｣の文字列が空白でなければ、｢期｣と｢当日入金｣の文字列を加えたものをセット
        If pay <> "" Then
          .Cells(nrow, COL_KIPAY) = ki & pay
        End If
      End If
 
      nrow = nrow + 1  '次の行へ
     
    Loop Until nrow = lrow + 1  '最終行+1になるまで繰り返し
   
    Application.StatusBar = False   'ステータスバーの消去
  End With
  
  ActiveSheet.EnableCalculation = True   '再計算を再開
  Application.ScreenUpdating = True      '描画を再開

End Sub
'
'「期別出欠集計」シートの各セルに集計のための関数を埋め込むマクロ
Sub 出欠確認関数貼付()

  Dim nrow As Long      '処理を行う行
  Dim mlrow As Long     '名簿シートの最終行
  Dim slrow As Long     '期別出欠集計シートの最終行
  Dim card As String    '「返信」列の文字列
  Dim tel As String     '「電話」列の文字列
  Dim plan As String    '「出予定」列の文字列
  Dim rslt As String    '「当日」列の文字列
  Dim apay As String    '「事前入金」列の文字列
  Dim kipay As String   '「期別入金」列の文字列
  Dim ki As String      '「期」列の文字列
  Dim mmbr1 As String   '「会員数」列の文字列
  Dim mmbr2 As String   '「会員数」列の文字列

  With Sheets("名簿")
    mlrow = .Cells(MEMBER_MAX, COL_ID).End(xlUp).Row
  End With
 
  Call SetKiNo("期別出欠集計")    '「期」の列を設定
  
  With Sheets("期別出欠集計")
    
    '期別出欠集計シートの最終行を取得
    slrow = .Cells(MEMBER_MAX, COL_KI).End(xlUp).Row
  
    '名簿シートの返信、電話、当日の登録状況から合計を集計するための計算式の文字列を設定
    card = "名簿!R" & ROW_TOPDATA & "C" & COL_CARD & ":R" & mlrow & "C" & COL_CARD
    tel = "名簿!R" & ROW_TOPDATA & "C" & COL_TEL & ":R" & mlrow & "C" & COL_TEL
    rslt = "名簿!R" & ROW_TOPDATA & "C" & COL_RSLT & ":R" & mlrow & "C" & COL_RSLT
 
    '合計を名簿シートの返信や電話の状況から集計する関数を合計の上の列に埋め込み
 
    '名簿の「返信」列からハガキ返信の合計を集計
    Range("B4").Formula = "=COUNTIF(" & card & ",""出ハ"")"
    Range("C4").Formula = "=COUNTIF(" & card & ",""欠ハ"")"
    Range("D4").Formula = "=COUNTIF(" & card & ",""不着"")"
    Range("E4").Formula = "=COUNTIF(" & card & ",""未ハ"")"
    Range("F4").Formula = "=SUM(RC[-4]:RC[-1])"
  
    '名簿の「返信」列から電子メール・HP登録の合計を集計
    Range("G4").Formula = "=COUNTIF(" & card & ",""出メ"")"
    Range("H4").Formula = "=COUNTIF(" & card & ",""欠メ"")"
    Range("I4").Formula = "=COUNTIF(" & card & ",""未メ"")"
    Range("J4").Formula = "=SUM(RC[-3]:RC[-1])"
 
    '名簿の「電話」列から電話の合計を集計
    Range("K4").Formula = "=COUNTIF(" & tel & ",""出"")"
    Range("L4").Formula = "=COUNTIF(" & tel & ",""欠"")"
    Range("M4").Formula = "=COUNTIF(" & tel & ",""不通"")"
    Range("N4").Formula = "=COUNTIF(" & tel & ",""未定"")"
    Range("O4").Formula = "=SUM(RC[-4]:RC[-1])"

    '合計を期別出欠集計シートの期別集計結果から集計する関数を合計の下の列に埋め込み
    '期別出欠集計からハガキ返信の合計を集計
    Range("B5").Formula = "=SUM(B" & ROW_KITOPDATA & ":B" & slrow & ")"
    Range("C5").Formula = "=SUM(C" & ROW_KITOPDATA & ":C" & slrow & ")"
    Range("D5").Formula = "=SUM(D" & ROW_KITOPDATA & ":D" & slrow & ")"
    Range("E5").Formula = "=SUM(E" & ROW_KITOPDATA & ":E" & slrow & ")"
    Range("F5").Formula = "=SUM(F" & ROW_KITOPDATA & ":F" & slrow & ")"

    '期別出欠集計からメール・HPの合計を集計
    Range("G5").Formula = "=SUM(G" & ROW_KITOPDATA & ":G" & slrow & ")"
    Range("H5").Formula = "=SUM(H" & ROW_KITOPDATA & ":H" & slrow & ")"
    Range("I5").Formula = "=SUM(I" & ROW_KITOPDATA & ":I" & slrow & ")"
    Range("J5").Formula = "=SUM(J" & ROW_KITOPDATA & ":J" & slrow & ")"
   
    '期別出欠集計から電話の合計を集計
    Range("K5").Formula = "=SUM(K" & ROW_KITOPDATA & ":K" & slrow & ")"
    Range("L5").Formula = "=SUM(L" & ROW_KITOPDATA & ":L" & slrow & ")"
    Range("M5").Formula = "=SUM(M" & ROW_KITOPDATA & ":M" & slrow & ")"
    Range("N5").Formula = "=SUM(N" & ROW_KITOPDATA & ":N" & slrow & ")"
    Range("O5").Formula = "=SUM(O" & ROW_KITOPDATA & ":O" & slrow & ")"

    '出席連絡の合計
    Range("P4").Formula = "=SUM(P" & ROW_KITOPDATA & ":P" & slrow & ")"

    '期別出欠集計から事前入金の合計を集計
    Range("R4").Formula = "=SUM(R" & ROW_KITOPDATA & ":R" & slrow & ")"
    Range("S4").Formula = "=SUM(S" & ROW_KITOPDATA & ":S" & slrow & ")"
    Range("T4").Formula = "=SUM(T" & ROW_KITOPDATA & ":T" & slrow & ")"

    '出席予定の合計
    Range("V4").Formula = "=SUM(V" & ROW_KITOPDATA & ":V" & slrow & ")"
    
    '当日出席の合計
    Range("X4").Formula = "=SUM(X" & ROW_KITOPDATA & ":X" & slrow & ")"
    Range("Y4").Formula = "=SUM(Y" & ROW_KITOPDATA & ":Y" & slrow & ")"
    Range("Z4").Formula = "=SUM(Z" & ROW_KITOPDATA & ":Z" & slrow & ")"
    
    '会員数の合計
    Range("AB4").Formula = "=名簿!G1"
    Range("AB5").Formula = "=SUM(AB" & ROW_KITOPDATA & ":AB" & slrow & ")"
 
    '期別返信、期別電話、出予定、期別当日出欠の集計をする範囲を表す文字列をセット
    card = "名簿!R" & ROW_TOPDATA & "C" & COL_KICARD & ":R" & mlrow & "C" & COL_KICARD
    tel = "名簿!R" & ROW_TOPDATA & "C" & COL_KITEL & ":R" & mlrow & "C" & COL_KITEL
    plan = "名簿!R" & ROW_TOPDATA & "C" & COL_PLAN & ":R" & mlrow & "C" & COL_PLAN
    rslt = "名簿!R" & ROW_TOPDATA & "C" & COL_KIRSLT & ":R" & mlrow & "C" & COL_KIRSLT
 
    apay = "名簿!R" & ROW_TOPDATA & "C" & COL_ADVPAY & ":R" & mlrow & "C" & COL_ADVPAY
    kipay = "名簿!R" & ROW_TOPDATA & "C" & COL_KIPAY & ":R" & mlrow & "C" & COL_KIPAY
 
    mmbr1 = "名簿!R" & ROW_TOPDATA & "C" & COL_KI & ":R" & mlrow & "C" & COL_KI
    mmbr2 = "名簿!R" & ROW_TOPDATA & "C" & COL_NAME & ":R" & mlrow & "C" & COL_NAME
 
    '期別の返信状況や電話状況等を集計する関数を期別のセルに埋め込み
    nrow = ROW_KITOPDATA
    Do
      ki = Cells(nrow, COL_KI).Text   '対象の「期」の文字列を取得
     
      If ki = "" Then
        End
        Exit Do
      Else
   
        '名簿シートの期別返信からハガキ返信状況の集計をする関数を返信ハガキの各列に埋め込み
        Cells(nrow, COL_CARD_OK).FormulaR1C1 = "=COUNTIF(" & card & ",""" & ki & "出ハ"")"
        Cells(nrow, COL_CARD_NG).FormulaR1C1 = "=COUNTIF(" & card & ",""" & ki & "欠ハ"")"
        Cells(nrow, COL_CARD_ERROR).FormulaR1C1 = "=COUNTIF(" & card & ",""" & ki & "不着"")"
        Cells(nrow, COL_CARD_NOFIX).FormulaR1C1 = "=COUNTIF(" & card & ",""" & ki & "未ハ"")"
        Cells(nrow, COL_CARD_SUM).FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"

        '名簿シートの期別返信からメール・Web返信状況の集計をする関数を返信メール・HPの各列に埋め込み
        Cells(nrow, COL_EMAIL_OK).FormulaR1C1 = "=COUNTIF(" & card & ",""" & ki & "出メ"")"
        Cells(nrow, COL_EMAIL_NG).FormulaR1C1 = "=COUNTIF(" & card & ",""" & ki & "欠メ"")"
        Cells(nrow, COL_EMAIL_NOFIX).FormulaR1C1 = "=COUNTIF(" & card & ",""" & ki & "未メ"")"
        Cells(nrow, COL_EMAIL_SUM).FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"

        '名簿シートの期別電話から電話の状況を集計をする関数を電話の各列に埋め込み
        Cells(nrow, COL_TEL_OK).FormulaR1C1 = "=COUNTIF(" & tel & ",""" & ki & "出"")"
        Cells(nrow, COL_TEL_NG).FormulaR1C1 = "=COUNTIF(" & tel & ",""" & ki & "欠"")"
        Cells(nrow, COL_TEL_ERROR).FormulaR1C1 = "=COUNTIF(" & tel & ",""" & ki & "不通"")"
        Cells(nrow, COL_TEL_NOFIX).FormulaR1C1 = "=COUNTIF(" & tel & ",""" & ki & "未定"")"
        Cells(nrow, COL_TEL_SUM).FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"

        Cells(nrow, COL_SUM_OK).FormulaR1C1 = "=COUNTIF(" & plan & ",""" & ki & "出予"")"

        '名簿シートの期別当日出席と期別入金から出席返信と事前入金の状況を集計をする関数を埋め込み
        Cells(nrow, COL_OKPAY).FormulaR1C1 = _
          "=SUMPRODUCT((" & plan & "=""" & ki & "出予"")*(" & apay & "<>""""))"
        Cells(nrow, COL_OTHERPAY).FormulaR1C1 = "=COUNTIF(" & plan & ",""" & ki & "見込"")"
        Cells(nrow, COL_SUMPAY).FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"

        '出予定＋出予定以外の事前入金済みを集計
        Cells(nrow, COL_SUM_ALL).FormulaR1C1 = "=SUM(RC[-6],RC[-3])"
     
        '名簿シートの期別当日出席から当日出席の状況を集計をする関数を当日出席列に埋め込み
        Cells(nrow, COL_RSLT_OK).FormulaR1C1 = "=COUNTIF(" & rslt & ",""" & ki & "出"")"
        Cells(nrow, COL_RSLT_PAYNG).FormulaR1C1 = "=COUNTIF(" & rslt & ",""" & ki & "入欠"")"
        Cells(nrow, COL_RSLT_SUM).FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"

        '名簿シートから会員数を集計をする関数を会員数列に埋め込み
        Cells(nrow, COL_MEMBER).FormulaR1C1 = _
           "=SUMPRODUCT((" & mmbr1 & "=""" & ki & """)*(" & mmbr2 & "<>""－""))"
      End If
  
      nrow = nrow + 1
    Loop Until nrow = slrow + 1  '最終行+1になるまで繰り返し
  
  End With
  ActiveSheet.EnableCalculation = True   '再計算

End Sub

'
'「期別入金集計」シートの各セルに集計のための関数を埋め込むマクロ
Sub 期別入金集計関数貼付()
  
  Dim nrow As Long      '処理対象の行
  Dim lrow As Long      '最終行
  Dim kipay As String  '合計計算のための計算式の文字列
  Dim ki As String   '「期」列の文字列
  
  Application.ScreenUpdating = False      '描画を中止
  ActiveSheet.EnableCalculation = False   '再計算を中止
  
  With Sheets("名簿")
    lrow = .Cells(MEMBER_MAX, COL_ID).End(xlUp).Row
  End With
 
  Call SetKiNo("期別入金集計")    '「期」の列を設定
 
  '名簿の事前入金の状況から合計を集計するための計算式の文字列を設定
  kipay = "名簿!R" & ROW_TOPDATA & "C" & COL_KIPAY & ":R" & lrow & "C" & COL_KIPAY
 
  With Sheets("期別入金集計")
    lrow = .Cells(MEMBER_MAX, COL_KI).End(xlUp).Row

    '事前入金の合計を集計する範囲をセット
    For nrow = ROW_KITOPDATA To lrow
      ki = Cells(nrow, COL_KI).Text   '「期」列の文字をセット

      '該当の「期」の入金種類別の集計関数をセット
      Cells(nrow, 2).FormulaR1C1 = "=SUM(RC[2],RC[4],RC[6],RC[8],RC[10],RC[12],RC[14],RC[16],RC[18],RC[20],RC[22],RC[24],RC[26],RC[28],RC[30],RC[32],RC[34],RC[36],RC[38],RC[40],RC[42],RC[44],RC[46],RC[48])"
      Cells(nrow, 3).FormulaR1C1 = "=SUM(RC[2],RC[4],RC[6],RC[8],RC[10],RC[12],RC[14],RC[16],RC[18],RC[20],RC[22],RC[24],RC[26],RC[28],RC[30],RC[32],RC[34],RC[36],RC[38],RC[40],RC[42],RC[44],RC[46],RC[48])"
      Cells(nrow, 4).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "MM"")"
      Cells(nrow, 5).FormulaR1C1 = "=RC[-1]*MM"
      Cells(nrow, 6).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "MF"")"
      Cells(nrow, 7).FormulaR1C1 = "=RC[-1]*MF"
      Cells(nrow, 8).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "MC"")"
      Cells(nrow, 9).FormulaR1C1 = "=RC[-1]*MC"
      Cells(nrow, 10).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "MY"")"
      Cells(nrow, 11).FormulaR1C1 = "=RC[-1]*MY"
      Cells(nrow, 12).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "MB"")"
      Cells(nrow, 13).FormulaR1C1 = "=RC[-1]*MB"
      Cells(nrow, 14).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "BM"")"
      Cells(nrow, 15).FormulaR1C1 = "=RC[-1]*BM"
      Cells(nrow, 16).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "BF"")"
      Cells(nrow, 17).FormulaR1C1 = "=RC[-1]*BF"
      Cells(nrow, 18).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "BC"")"
      Cells(nrow, 19).FormulaR1C1 = "=RC[-1]*BC"
      Cells(nrow, 20).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "BY"")"
      Cells(nrow, 21).FormulaR1C1 = "=RC[-1]*BY"
      Cells(nrow, 22).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "BB"")"
      Cells(nrow, 23).FormulaR1C1 = "=RC[-1]*BB"
      Cells(nrow, 24).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "CM"")"
      Cells(nrow, 25).FormulaR1C1 = "=RC[-1]*CM"
      Cells(nrow, 26).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "CF"")"
      Cells(nrow, 27).FormulaR1C1 = "=RC[-1]*CF"
      Cells(nrow, 28).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "CC"")"
      Cells(nrow, 29).FormulaR1C1 = "=RC[-1]*CC"
      Cells(nrow, 30).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "CY"")"
      Cells(nrow, 31).FormulaR1C1 = "=RC[-1]*CY"
      Cells(nrow, 32).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "CB"")"
      Cells(nrow, 33).FormulaR1C1 = "=RC[-1]*CB"
      Cells(nrow, 34).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "YM"")"
      Cells(nrow, 35).FormulaR1C1 = "=RC[-1]*YM"
      Cells(nrow, 36).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "YF"")"
      Cells(nrow, 37).FormulaR1C1 = "=RC[-1]*YF"
      Cells(nrow, 38).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "YY"")"
      Cells(nrow, 39).FormulaR1C1 = "=RC[-1]*YY"
      Cells(nrow, 40).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "GM"")"
      Cells(nrow, 41).FormulaR1C1 = "=RC[-1]*GM"
      Cells(nrow, 42).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "GF"")"
      Cells(nrow, 43).FormulaR1C1 = "=RC[-1]*GF"
      Cells(nrow, 44).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "GY"")"
      Cells(nrow, 45).FormulaR1C1 = "=RC[-1]*GY"
      Cells(nrow, 46).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "KM"")"
      Cells(nrow, 47).FormulaR1C1 = "=RC[-1]*KM"
      Cells(nrow, 48).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "KF"")"
      Cells(nrow, 49).FormulaR1C1 = "=RC[-1]*KF"
      Cells(nrow, 50).FormulaR1C1 = "=COUNTIF(" & kipay & ",""" & ki & "KY"")"
      Cells(nrow, 51).FormulaR1C1 = "=RC[-1]*KY"
    Next nrow
  End With

  ActiveSheet.EnableCalculation = True   '再計算を再開
  Application.ScreenUpdating = True      '描画を再開

End Sub


'
'「入金集計」シートの各セルに集計のための関数を埋め込むマクロ
Sub 入金集計関数貼付()
  
  Dim nrow As Long      '処理対象の行
  Dim lrow As Long      '名簿シートの最終行
  Dim trow As Long      '当番期の先頭行
  Dim erow As Long      '当番期の最終行
  Dim kipay As String   '合計計算のための計算式の文字列
  Dim ki As String      '「期」列の文字列
  Dim obj As Range      '検索結果のオブジェクト
  
  
  Application.ScreenUpdating = False      '描画を中止
  ActiveSheet.EnableCalculation = False   '再計算を中止
  
  With Sheets("名簿")
    lrow = .Cells(MEMBER_MAX, COL_ID).End(xlUp).Row
  End With
 
  With Sheets("名簿").Range("A1:A" & MEMBER_MAX)
    ki = Sheets("入金集計").Range("C1").Text   '入金集計シートから当番期の文字列を取得
    Set obj = .Find(ki, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                 SearchOrder:=xlByColumns, MatchCase:=False, MatchByte:=False)
    trow = Range(obj.Address).Row        '当番期の先頭行を取得
  
    ki = ki + 1                   '検索文字列を当番期＋1に
    Set obj = .Find(ki, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                 SearchOrder:=xlByColumns, MatchCase:=False, MatchByte:=False)
    erow = Range(obj.Address).Row - 1    '当番期の最終行を取得
  
  End With
 
  With Sheets("入金集計")

    '名簿の事前入金の状況から当番期分を集計するための計算式の文字列を設定
    kipay = "名簿!R" & trow & "C" & COL_ADVPAY & ":R" & erow & "C" & COL_ADVPAY
    Range("D4").Formula = "=COUNTIF(" & kipay & ",""MM"")"
    Range("D5").Formula = "=COUNTIF(" & kipay & ",""MF"")"
    Range("D6").Formula = "=COUNTIF(" & kipay & ",""MC"")"
    Range("D8").Formula = "=COUNTIF(" & kipay & ",""BM"")"
    Range("D9").Formula = "=COUNTIF(" & kipay & ",""BF"")"
    Range("D10").Formula = "=COUNTIF(" & kipay & ",""BC"")"
    Range("D12").Formula = "=COUNTIF(" & kipay & ",""CM"")"
    Range("D13").Formula = "=COUNTIF(" & kipay & ",""CF"")"
    Range("D14").Formula = "=COUNTIF(" & kipay & ",""CC"")"
    
    '名簿の当日入金の状況から当番期分を集計するための計算式の文字列を設定
    kipay = "名簿!R" & trow & "C" & COL_PAY & ":R" & erow & "C" & COL_PAY
    Range("F4").Formula = "=COUNTIF(" & kipay & ",""YM"")"
    Range("F5").Formula = "=COUNTIF(" & kipay & ",""YF"")"
    Range("F8").Formula = "=COUNTIF(" & kipay & ",""GM"")"
    Range("F9").Formula = "=COUNTIF(" & kipay & ",""GF"")"
    Range("F12").Formula = "=COUNTIF(" & kipay & ",""KM"")"
    Range("F13").Formula = "=COUNTIF(" & kipay & ",""KF"")"
    
    '名簿の事前入金の状況から全体の集計するための計算式の文字列を設定
    kipay = "名簿!R" & ROW_TOPDATA & "C" & COL_ADVPAY & ":R" & lrow & "C" & COL_ADVPAY
    Range("D36").Formula = "=COUNTIF(" & kipay & ",""MM"")"
    Range("D37").Formula = "=COUNTIF(" & kipay & ",""MF"")"
    Range("D38").Formula = "=COUNTIF(" & kipay & ",""MC"")"
    Range("D39").Formula = "=COUNTIF(" & kipay & ",""MY"")"
    Range("D40").Formula = "=COUNTIF(" & kipay & ",""MB"")"
    Range("D42").Formula = "=COUNTIF(" & kipay & ",""BM"")"
    Range("D43").Formula = "=COUNTIF(" & kipay & ",""BF"")"
    Range("D44").Formula = "=COUNTIF(" & kipay & ",""BC"")"
    Range("D45").Formula = "=COUNTIF(" & kipay & ",""BY"")"
    Range("D46").Formula = "=COUNTIF(" & kipay & ",""BB"")"
    Range("D48").Formula = "=COUNTIF(" & kipay & ",""CM"")"
    Range("D49").Formula = "=COUNTIF(" & kipay & ",""CF"")"
    Range("D50").Formula = "=COUNTIF(" & kipay & ",""CC"")"
    Range("D51").Formula = "=COUNTIF(" & kipay & ",""CY"")"
    Range("D52").Formula = "=COUNTIF(" & kipay & ",""CB"")"
    
    '名簿の当日入金の状況から全体の集計するための計算式の文字列を設定
    kipay = "名簿!R" & ROW_TOPDATA & "C" & COL_PAY & ":R" & lrow & "C" & COL_PAY
    Range("F36").Formula = "=COUNTIF(" & kipay & ",""YM"")"
    Range("F37").Formula = "=COUNTIF(" & kipay & ",""YF"")"
    Range("F39").Formula = "=COUNTIF(" & kipay & ",""YY"")"
    Range("F42").Formula = "=COUNTIF(" & kipay & ",""GM"")"
    Range("F43").Formula = "=COUNTIF(" & kipay & ",""GF"")"
    Range("F45").Formula = "=COUNTIF(" & kipay & ",""GY"")"
    Range("F48").Formula = "=COUNTIF(" & kipay & ",""KM"")"
    Range("F49").Formula = "=COUNTIF(" & kipay & ",""KF"")"
    Range("F51").Formula = "=COUNTIF(" & kipay & ",""KY"")"
    
  End With

  ActiveSheet.EnableCalculation = True   '再計算を再開
  Application.ScreenUpdating = True      '描画を再開

End Sub

'期別出欠集計シート、期別入金集計シートの「期」の列を設定する。
'  StName  ：設定するシート名
Private Sub SetKiNo(StName As String)

  Dim nrow As Long      '処理対象の行
  Dim lrow As Long      '最終行
  Dim kinum As Integer  '期の数
  Dim ki(KI_MAX) As String   '期の文字列を格納する配列
  Dim i As Integer
    
  '対象の期別シートの「期」の列をクリア
  With Sheets(StName)
    lrow = .Cells(KI_MAX, COL_KI).End(xlUp).Row
    Worksheets(StName).Range(Cells(ROW_KITOPDATA, COL_KI), Cells(lrow, COL_KI)).ClearContents
  End With
  
  '名簿シートから「期」を取得
  With Sheets("名簿")
    lrow = .Cells(MEMBER_MAX, COL_ID).End(xlUp).Row
  
    kinum = 0
    ki(kinum) = .Cells(ROW_TOPDATA, COL_KI)
    For nrow = ROW_TOPDATA To lrow
      If ki(kinum) <> .Cells(nrow, COL_KI) Then
        kinum = kinum + 1
        ki(kinum) = .Cells(nrow, COL_KI)
      End If
    Next nrow
  End With

  '対象の期別シートに「期」をセット
  For i = 0 To kinum
    Worksheets(StName).Cells(ROW_KITOPDATA + i, COL_KI) = "'" & ki(i)
  Next i

End Sub

'
'懇親会終了後に本年度の実績を懇親会履歴に転記するマクロ
Sub 懇親会実績転記()

  Dim nrow As Long    '処理行
  Dim lrow As Long    '名簿シートの最終行

  'マクロを実行するか確認
  If MsgBox("懇親会終了後に本年度の実績を転記する処理です。実行しますか？", _
      vbYesNo + vbQuestion, "処理実行") = vbYes Then

    With Sheets("名簿")
      lrow = .Cells(MEMBER_MAX, COL_ID).End(xlUp).Row
    End With

    For nrow = ROW_TOPDATA To lrow
   
      If Cells(nrow, COL_RSLT) = "出" And Cells(nrow, COL_ADVPAY) <> "" Then    '当日出席(事前支払)
        Cells(nrow, COL_RSLT0) = "◎"
      ElseIf Cells(nrow, COL_RSLT) = "出" And Cells(nrow, COL_PAY) <> "" Then   '当日出席(当日支払)
        Cells(nrow, COL_RSLT0) = "○"
      ElseIf Cells(nrow, COL_RSLT) = "出" And Cells(nrow, COL_ADVPAY) = "" _
          And Cells(nrow, COL_PAY) = "" Then                                    '当日出席(無料招待)
        Cells(nrow, COL_RSLT0) = "☆"
      ElseIf Cells(nrow, COL_RSLT) = "" And Cells(nrow, COL_ADVPAY) <> "" Then  '当日欠席(事前振込)
        Cells(nrow, COL_RSLT0) = "欠◎"
      ElseIf Cells(nrow, COL_RSLT) = "" And _
          ((Cells(nrow, COL_CARD) = "出" And Cells(nrow, COL_TEL) <> "欠") _
            Or Cells(nrow, COL_TEL) = "出") Then                                 '出席連絡あるも当日欠席
        Cells(nrow, COL_RSLT0) = "／"
      ElseIf Cells(nrow, COL_RSLT) = "" And _
          ((Cells(nrow, COL_CARD) = "欠" And Cells(nrow, COL_TEL) <> "出") _
            Or Cells(nrow, COL_TEL) = "欠") Then                                 '欠席連絡あり
        Cells(nrow, COL_RSLT0) = "×"
      ElseIf Cells(nrow, COL_RSLT) = "" And Cells(nrow, COL_CARD) = "不着" Then  '郵便不着
        Cells(nrow, COL_RSLT0) = "？"
      End If
    Next nrow
  End If
End Sub
'
'「名簿」シートの「作業表転記」ボタンが押されたときの処理
    '== Mar.9,2014 Yuji Ogihara ==
    '保存形式を2007 以降の書式に変更、拡張子を"xlsm"に
'「宛名シール・名札・出席者一覧表作成.xlsm」で宛名シールの作成等を行うために原本データを転記
Sub 作業表転記()

  Dim lrow As Long       '名簿シートの最終行
  Dim msg As String

  OrgBook = ActiveWorkbook.Name      '開いているファイルの名前を原本として設定

  '処理に必要なファイルが開かれているか確認
     '== Mar.9,2014 Yuji Ogihara ==
     '保存形式を2007 以降の書式に変更
  'If IsBookOpen("宛名シール・名札・出席者一覧表作成.xls") = False Then
  If IsBookOpen(SAGYOHYO_FILENAME) = False Then
    '== Mar.9,2014 Yuji Ogihara ==
    '保存形式を2007 以降の書式に変更
    'MsgBox "「"宛名シール・名札・出席者一覧表作成.xls"」が開かれていません！" _

    MsgBox "「" & SAGYOHYO_FILENAME & "」が開かれていません！" _
      & vbNewLine & "開いてからやり直して下さい。"
    End
  End If
  
  msg = "作業表に名簿を転記しますか？"
  If MsgBox(msg, 4 + 32, "作業表転記") = 6 Then

    '期別の懇親会出欠状況の情報を隠し列にセットするマクロを実行
    Call 出欠確認

    '｢宛名シール・名札・出席者一覧表作成.xls｣を開き、貼付部分をクリア
        '== Mar.9,2014 Yuji Ogihara ==
        '保存形式を2007 以降の書式に変更
    'Workbooks("宛名シール・名札・出席者一覧表作成.xls").Sheets("原本").Activate
    Workbooks(SAGYOHYO_FILENAME).Sheets("原本").Activate
    Rows("2:65536").Delete Shift:=xlUp
    Range("A1").Select
 
    Workbooks(OrgBook).Sheets("名簿").Activate
    lrow = Cells(MEMBER_MAX, COL_KI).End(xlUp).Row    '最終行の行数を取得
    
    '名簿の記載部分全てをコピー
    Range(Cells(ROW_TOPTITLE, COL_KI), Cells(lrow, COL_COMMENT)).Select
    Selection.Copy
 
   '｢宛名シール・名札・出席者一覧表作成.xls｣を開き、作業表に貼り付け
        '== Mar.9,2014 Yuji Ogihara ==
        '保存形式を2007 以降の書式に変更
   'Workbooks("宛名シール・名札・出席者一覧表作成.xls").Sheets("原本").Activate
    Workbooks(SAGYOHYO_FILENAME).Sheets("原本").Activate
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range("A4").Select
 
    '名簿を開き、コピーモードをキャンセル
    Workbooks(OrgBook).Sheets("名簿").Activate
    Application.CutCopyMode = False
    Range("A5").Select
 
    '== Feb.9,2014 Yuji Ogihara ==
    '保存形式を2007 以降の書式に変更
    'Workbooks("宛名シール・名札・出席者一覧表作成.xls").Sheets("原本").Activate
    Workbooks(SAGYOHYO_FILENAME).Sheets("原本").Activate

 End If
 
End Sub

'集客担当用に担当分の名簿ファイルの作成
'　「期別出欠集計」シートの集客担当列の担当者名をキーに名簿ファイルを担当者毎に分割する
'　作成されるファイルは、原本ファイルのあるフォルダーに「配布用」フォルダーを作り保存される
Sub 集客担当用名簿ファイル作成()

  Dim tantoki(COL_TANTO_MAX, KI_MAX) As String     '担当者名と担当する期を保存するの２次元配列
  Dim kinum(KI_MAX) As Integer                   '担当者毎の担当する期の数
  Dim serow(COL_TANTO_MAX, KI_MAX * 2) As Long   '担当者名とその担当の期の開始行、終了行を保存するの２次元配列
  Dim tmax As Integer      '担当者数
  Dim opath As String      '原本ファイルのパス
     
  OrgBook = ActiveWorkbook.Name      '開いているファイルの名前を原本として設定
  If Mid(OrgBook, 9, 2) = "原本" Then
  
    Application.ScreenUpdating = False  '描画を停止

    '期別出欠集計シートの「集客担当」列の名前の情報と担当する期を取得
    Call GetSyukyakuInfo(tantoki, kinum, tmax)

    '原本ファイルから、集客担当が担当する期の開始行と終了行を取得
    Call GetSyukyakuRow(opath, tantoki, kinum, tmax, serow)
   
    '集客担当への配布用に原本ファイルから担当する期の情報を切り出しファイル作成
    Call MakeSyukyakuFile(opath, tantoki, kinum, tmax, serow)

    Application.ScreenUpdating = True   '描画を再開
    
  Else
    MsgBox ("このファイルは原本でないのでこの処理はできません！")
  End If

End Sub

'
'期別出欠集計シートの設定されて集客担当名と担当する期の情報を取得するプロシージャ
' tantoki() ：担当者名と担当する期が格納された２次元配列
' kinum()   ：担当者毎の担当する期の数が格納された配列
' tmax      ：担当者数
Private Sub GetSyukyakuInfo(tantoki() As String, kinum() As Integer, tmax As Integer)

  Dim lrow As Integer      '行数
  Dim ki As String      '期
  Dim tanto As String      '担当者名
  Dim tno As Integer      '担当者番号
  Dim i As Integer

  '期別出欠集計シートの行数をカウント
  lrow = Worksheets("期別出欠集計").Range("A200").End(xlUp).Row

  '期別出欠集計シートの「集客担当」列の名前分のシートを作成
  tmax = 0
  For i = ROW_KITOPDATA To lrow
    ki = Worksheets("期別出欠集計").Cells(i, COL_KI).Value
    tanto = Worksheets("期別出欠集計").Cells(i, COL_TANTO).Value
  

    '二次元配列tNameに(X,0)に担当者名、(X,1..)に担当の期をセット
    If tmax = 0 Then    '最初の期の集客担当の情報をセット
      tantoki(0, 0) = tanto
      kinum(0) = 1
      tantoki(0, kinum(0)) = ki
      tmax = 1
    Else             '以降の期の集客担当の情報をセット
      For tno = 0 To tmax - 1
        If tanto = tantoki(tno, 0) Then
          kinum(tno) = kinum(tno) + 1
          tantoki(tno, kinum(tno)) = ki
          Exit For
        ElseIf tno = tmax - 1 Then
          tantoki(tmax, 0) = tanto
          kinum(tmax) = 1
          tantoki(tmax, kinum(tmax)) = ki
          tmax = tmax + 1
        End If
      Next tno
    End If
  Next i

End Sub

'
'原本の名簿シートから担当毎の担当する期の開始行・終了行を取得するプロシージャ
' opath     ：原本ファイルがあるフォルダー名
' tantoki()   ：担当者名と担当する期が格納された２次元配列
' kinum()   ：担当者毎の担当する期の数が格納された配列
' tmax      ：担当者数
' serow()   ：担当者名と担当する期の開始行・終了行が格納された２次元配列
Private Sub GetSyukyakuRow(opath As String, tantoki() As String, kinum() As Integer, tmax As Integer, serow() As Long)

  Dim tnNum As Integer     '担当者番号
  Dim nrow As Integer      '処理している行番号
  Dim obj As Range         '検索結果のオブジェクト
  Dim i As Integer
  
  Workbooks(OrgBook).Activate      '原本ブックをアクティブに
  opath = ActiveWorkbook.Path      '原本ファイルのパスを取得
  
  For tnNum = 0 To tmax - 1
    For i = 1 To kinum(tnNum)
    
      '名簿のKINO列から期を検索
      With Sheets("名簿").Range("A1:A" & MEMBER_MAX)
        Set obj = .Find(tantoki(tnNum, i), LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                  SearchOrder:=xlByColumns, MatchCase:=False, MatchByte:=False)
      End With
     
      If Not obj Is Nothing Then
        nrow = Range(obj.Address).Row
        serow(tnNum, i * 2 - 1) = nrow     '開始行をセット
     
        '一行下の期が異なるまで繰り返し
        Do
          'ステータスバーに状況表示
          Application.StatusBar = "  ◆集客担当 " & tantoki(tnNum, 0) & " が担当する期の開始行・終了行を取得中．．．"
          
          If Cells(nrow, COL_KI) = Cells(nrow + 1, COL_KI) Then
            nrow = nrow + 1
          Else        '期番号が違えば、終了行をセットし処理を中止
            serow(tnNum, i * 2) = nrow
            Exit Do
          End If
        Loop Until nrow = MEMBER_MAX + 1
      Else
        serow(tnNum, i * 2 - 1) = 0
        serow(tnNum, i * 2) = 0
      End If
    Next i
  Next tnNum
  Application.StatusBar = False   'ステータスバーをクリア

End Sub

'
'原本の名簿シートから担当毎の担当する期の情報を分割して保存するプロシージャ
' opath     ：原本ファイルがあるフォルダー名
' tantoki() ：担当者名と担当する期が格納された２次元配列
' kinum()   ：担当者毎の担当する期の数が格納された配列
' tMmx      ：担当者数
' serow()：担当者名と担当する期の開始行・終了行が格納された２次元配列
Private Sub MakeSyukyakuFile(opath As String, tantoki() As String, kinum() As Integer, tmax As Integer, serow() As Long)
  
  Dim lrow As Integer   '行数
  Dim obj As Object     'フォルダーを示すオブジェクト
  Dim npath As String   '分割ファイルを保存するフォルダー名
  Dim tnum As Integer   '担当者数
  Dim nbook As String   '分割後のブック名
  Dim trow As Long      '不要部分の先頭行番号
  Dim brow As Long      '不要部分の最終行番号
  Dim dastr As String   'ファイル名使用する月日の文字列
  Dim i As Integer
   
  lrow = Worksheets("名簿").Range("A" & MEMBER_MAX).End(xlUp).Row  '名簿の行数を取得

  '原本ファイルと同じフォルダー内に配布用フォルダーがなければ作成
  Set obj = CreateObject("Scripting.FileSystemObject")
  npath = opath & "\配布用"
  If obj.FolderExists(folderspec:=npath) = False Then
    obj.createfolder npath
  End If

  '担当毎に分割した原本ファイルのコピーを作成
  For tnum = 0 To tmax - 1
     
    'ステータスバーに状況表示
    Application.StatusBar = "  ◆集客担当 " & tantoki(tnum, 0) & " 用のファイルを分割作成中．．．"
    
    Workbooks.Add                    '集客担当別に新規にブックを追加
    nbook = ActiveWorkbook.Name    '追加したブックの名前を取得
    Workbooks(OrgBook).Activate      '原本ブックをアクティブに
  
    '新規ブックのSheet1の前に「名簿」をコピー
    ThisWorkbook.Sheets("名簿").Copy After:=Workbooks(nbook).Sheets("Sheet1")
  
    '新規ブックのSheet1～3を確認メッセージなしに削除
    Application.DisplayAlerts = False
    Workbooks(nbook).Worksheets("Sheet1").Delete
    
    '== Feb.9,2014 Yuji Ogihara ==
    '新規ブック作成時のシート数はExcel 2007まで「3」, 2010以降は「1」(設定にて変更可能)
    '以下の2行はシート数「3」のときの処理なので、2010以降では不要
    'Workbooks(nbook).Worksheets("Sheet2").Delete
    'Workbooks(nbook).Worksheets("Sheet3").Delete
    Application.DisplayAlerts = True

    '== Feb.9,2014 Yuji Ogihara ==
    '以降の行削除処理 "Delete" の高速化のため
    'データ最終行からワークシート最終行までを削除
    Range(lrow + 100 & ":" & Rows.Count).Delete

    '担当以外の期の情報を最終行から削除
    brow = lrow

    For i = kinum(tnum) * 2 To 2 Step -2
      trow = serow(tnum, i) + 1
      If trow < brow And trow <> 1 Then
        Range(trow & ":" & brow).Delete
        brow = serow(tnum, i - 1) - 1
      ElseIf trow > brow Then
        brow = serow(tnum, i - 1) - 1
      End If
    Next i
    If brow > ROW_TOPDATA Then
      Range(ROW_TOPDATA & ":" & brow).Delete
    End If
    Cells(ROW_TOPDATA, COL_KI).Select
  
    '新規ブックを担当者の名前を付加したファイル名で保存
    '== Feb.9,2014 Yuji Ogihara ==
    '保存形式を2007 以降の書式に変更、拡張子を"xlsx"に
    dastr = Format(Month(Date), "00") & Format(Day(Date), "00")
    Workbooks(nbook).SaveAs _
       Filename:=opath & "\配布用\東京東筑会名簿【電話】to" & tantoki(tnum, 0) & dastr & ".xlsx", _
       Password:=MEIBO_PASSWORD
    ActiveWorkbook.Close
  Next tnum
  MsgBox "分割処理が終了しました。"
  Application.StatusBar = False   'ステータスバーをクリア
End Sub


'
'名簿引継ぎ時に、①年会費を直近3年分＋新年度分にし、②懇親会実績を直近5年分＋新年度分にし、
'③懇親会用データをクリアする
Sub 引継ぎ処理()

  Dim msg As String   'メッセージ
  Dim ystr As String  '年度文字列
  Dim afile As String '開いているファイル名
  Dim apath As String '開いているパス名
  Dim bfile As String 'バックアップファイル名
  Dim nrow As Long    '処理行
  Dim lrow As Long    '名簿シートの最終行

  'マクロを実行するか確認
  ystr = Cells(ROW_TOPDATA - 1, COL_KAIHI0)
  msg = "20" & ystr & "年度から翌年度への名簿引継ぎ時の処理を実行しますか？"
  If MsgBox(msg, vbYesNo + vbQuestion, "処理実行") = vbYes Then

    '処理前のファイルのバックアップを作成
'    afile = ActiveWorkbook.Name         '開いているファイルの名前を取得
'    apath = ActiveWorkbook.Path         '開いているファイルのパスを取得
'    bfile = Left(afile, Len(afile) - 4) & "_OLD.xls"
'    Application.DisplayAlerts = False   '保存確認のメッセージを出さない
'    ActiveWorkbook.SaveCopyAs bfile
'    Application.DisplayAlerts = True

    With Sheets("名簿")
      lrow = .Cells(MEMBER_MAX, COL_ID).End(xlUp).Row
    End With

    '年会費の「2年前～今年度」を「3年前～1年前」にコピーし、「今年度」をクリア
    Range(Cells(ROW_TOPDATA - 1, COL_KAIHI2), Cells(lrow, COL_KAIHI0)).Copy
    Range(Cells(ROW_TOPDATA - 1, COL_KAIHI3), Cells(lrow, COL_KAIHI1)).PasteSpecial Paste:=xlValues
    Range(Cells(ROW_TOPDATA, COL_KAIHI0), Cells(lrow, COL_KAIHI0)).ClearContents
    Cells(ROW_TOPDATA - 1, COL_KAIHI0) = ystr + 1

    '懇親会実績の「4年前～今年度」を「5年前～1年前」にコピーし、「今年度」をクリア
    Range(Cells(ROW_TOPDATA - 1, COL_RSLT4), Cells(lrow, COL_RSLT0)).Copy
    Range(Cells(ROW_TOPDATA - 1, COL_RSLT5), Cells(lrow, COL_RSLT1)).PasteSpecial Paste:=xlValues
    Range(Cells(ROW_TOPDATA, COL_RSLT0), Cells(lrow, COL_RSLT0)).ClearContents
    Cells(ROW_TOPDATA - 1, COL_RSLT0) = ystr + 1

    '懇親会用データをクリア
    Range(Cells(ROW_TOPDATA, COL_CARD), Cells(lrow, COL_COMMENT)).ClearContents
    Cells(ROW_TOPTITLE, COL_CARD) = "20" & (ystr + 1) & "懇親会"
    
    Range("T5").Select
    MsgBox "引継ぎ処理が終了しました。"
  End If
End Sub


