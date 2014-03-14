VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm転記 
   Caption         =   "転記"
   ClientHeight    =   3555
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4284
   OleObjectBlob   =   "UserForm転記.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm転記"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
'転記フォームで「転記」ボタンが押された場合のメイン処理
'　「返信」ファイル：返信列が異なる場合に返信列を転記
'　「電話」ファイル：電話列が異なる場合に電話列を転記
'　「事前」ファイル：事前列が異なる場合に事前列を転記
'　「本部」ファイル：年会費列が異なる場合に年会費列を転記
'　「出席」ファイル：出席列が異なる場合に出席列を転記
'　「当日」ファイル：当日列が異なる場合に当日列を転記
'   いずれのファイルの場合も、コメント欄に変更があった場合は追記。
'   また、基本情報（カナ氏名、性別、郵便番号、住所1～4、電話番号、メール、部活、出身中学、夫婦）が
'   変更されている場合は転記判断の後に基本情報についても転記。
'   なお、姓についてはデータ一致の確認に使用しているためマクロでは転記しないようにしている。
Private Sub CommandButton1_Click()

  Dim job As String       '転記処理の内容
  Dim pbook As String     '転記元ファイル名
  Dim ski As String       '転記する開始期
  Dim eki As String       '転記する終了期
  Dim srow As Long        '転記する開始行
  Dim erow As Long        '転記する終了行
   
  'マクロを起動したファイルが原本かを確認
  OrgBook = ActiveWorkbook.Name
  If Mid(OrgBook, 9, 2) <> "原本" Then
    MsgBox "原本ファイル名が異常です！（原本への転記になっていないと思われます）" _
      & vbNewLine & "原本ファイル『" & OrgBook & "』"
    End
  End If
 
  '転記処理の内容（処理内容、転記元ファイル、開始期、終了期）を取得
  Call GetInputData(job, pbook, ski, eki)
  If job = "ERROR" Then
    End
  End If

  '転記元ファイルを開いているか確認
   If IsBookOpen(pbook) = False Then
     MsgBox "転記元ファイル『" & pbook & "』が開かれていません！" & vbNewLine & "開いてからやり直して下さい。"
     End
   End If
  
  '転記処理する期が転記元ファイルに存在するかを確認し、その開始行・終了行を取得
  Call ChkKiData(pbook, ski, eki, srow, erow)
  If srow = -1 Or erow = -1 Then
    End
  End If
   
  '転記処理する開始行・終了行を表示し確認
  Call ShowStartEndRow(pbook, srow, erow)
  If srow = -1 Or erow = -1 Then
    End
  End If

  '転記元ファイルから原本に転記するための準備
  Call InitPostData(pbook)

  '転記元ファイルから原本に転記実施（基本情報も必要に応じて転記）
  Call PostAllData(job, pbook, srow, erow)

End Sub


'転記フォームに入力された情報を設定するプロシージャ
'   job     ：フォームで選択されて処理の文字列（エラーの場合は戻り値として"ERROR"を設定）
'   pbook ：転記元のファイル名
'   ski,eki ：転記開始の期と終了の期
Private Sub GetInputData(job As String, pbook As String, ski As String, eki As String)

  Dim jkey As String      'ファイル名に埋め込まれている処理キーワード

  '選択された処理のオプションを取得
  If OptionButton1.Value = True Then
    job = "返信"
  ElseIf OptionButton2.Value = True Then
    job = "電話"
  ElseIf OptionButton3.Value = True Then
    job = "事前"
  ElseIf OptionButton4.Value = True Then
    job = "本部"
  ElseIf OptionButton5.Value = True Then
    job = "出席"
  ElseIf OptionButton6.Value = True Then
    job = "当日"
  End If
  
  pbook = TextBox1.Text    'ターゲット（転記元）ファイル名
  jkey = Mid(pbook, 9, 2)
  
  If job = jkey Then   '選択された処理と指定されたファイルが一致
  
    ski = TextBox11.Text      '転記を開始する「期」
    eki = TextBox10.Text      '転記を終了する「期」

  Else
    MsgBox "選択された作業種類とファイル名が一致しません！"
    job = "ERROR"    'エラーとする
  End If

End Sub
'
'転記対象として設定された開始・終了の期が存在するか確認し、開始行・終了行を取得するプロシージャ
'   pbook  ：転記元のファイル名
'   ski,eki  ：転記開始の期と終了の期
'   sRow,eRow：転記開始行と終了行（エラーの場合は戻り値として"-1"を設定）
Private Sub ChkKiData(pbook As String, ski As String, eki As String, srow As Long, erow As Long)
   
  Dim msg As String
  Dim obj As Range
  Dim i As Integer
  
  msg = "指定されたファイル " & pbook & " から転記しますか？"
  
  If MsgBox(msg, 4 + 32, "転記") = 6 Then  '「はい」をクリック
     
    If ski <> "" And eki = "" Then    '終了の期だけが未設定の場合
      eki = ski
    End If
      
    '転記元ファイルのシート「名簿」の「期」列で「転記開始行」を検索
    With Workbooks(pbook).Sheets("名簿").Range("A1:A" & MEMBER_MAX)
      
      If ski = "" Then     '転記開始の期が未設定の場合
        srow = ROW_TOPDATA
      
      Else                 '転記開始の期が設定されている場合
        Set obj = .Find(ski, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                 SearchOrder:=xlByColumns, MatchCase:=False, MatchByte:=False)
                 
        If Not obj Is Nothing Then       '該当の「期」番号が検索された
          srow = .Range(obj.Address).Row
        Else  '期番号が検索されなければメッセージを表示し、エラーをセット
          MsgBox "「転記開始期」は検索できませんでした！"
          srow = -1
          Exit Sub
        End If
      End If
        
      '転記元ファイルのシート「名簿」の「期」列で「転記終了行」を検索
      If eki = "" Then     '転記終了の期が未設定の場合
        erow = .Range("A" & MEMBER_MAX).End(xlUp).Row
      
      Else                 '転記終了の期が設定されている場合
        Set obj = .Find(eki, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                 SearchOrder:=xlByColumns, MatchCase:=False, MatchByte:=False)
                                     
        If Not obj Is Nothing Then
          erow = .Range(obj.Address).Row     '「期」番号が検索された行数をセット
          i = 0
          Do
            If .Cells(erow, COL_KI) = .Cells(erow + 1, COL_KI) Then
              erow = erow + 1
              i = i + 1
            Else    '期番号が違えば、繰り返し処理を中止（その時点のLRowが最終行）
              Exit Do
            End If
          Loop Until i = DOUKI_MAX  'DOUKI_MAXになるまで繰り返す
    
        Else  '期番号が検索されなければ、メッセージを表示し、エラーをセット
          MsgBox "「転記終了期」は検索できませんでした！"
          erow = -1
          Exit Sub
        End If
      End If
    End With
  End If
End Sub
'
'転記対象として設定された開始・終了を表示し確認するプロシージャ
'   pbook  ：転記元のファイル名
'   srow,erow：転記開始行と終了行（エラーの場合は戻り値として"-1"を設定）
Private Sub ShowStartEndRow(pbook As String, srow As Long, erow As Long)

  Dim msg1 As String
  Dim msg2 As String

  '転記開始行のA列を選択し、メッセージを表示して確認
  Workbooks(pbook).Sheets("名簿").Activate
  Range(Cells(srow, COL_KI), Cells(srow, COL_KI)).EntireRow.Select

  msg1 = "この行から転記しますか？"
   
  '転記終了行のA列を選択し、メッセージを表示して確認
  If MsgBox(msg1, 4 + 32, "転記開始行確認") = 6 Then
    Range(Cells(erow, COL_KI), Cells(erow, COL_KI)).EntireRow.Select
    msg2 = "この行まで転記しますか？"
       
    If MsgBox(msg2, 4 + 32, "転記終了行確認") = 6 Then
      Cells(srow, COL_KI).Select    '転記開始行のA列を選択する。
    Else   '転記終了行でなければエラー
      erow = -1
    End If
      
  Else  '転記開始行でなければエラー
    srow = -1
  End If
End Sub
'
'転記元ファイルのデータを原本ファイルに転記する準備のためのプロシージャ
'  pBook   ：転記元ファイル名
Private Sub InitPostData(pbook As String)
 
  Dim lrow As Long        '最終行

 '原本ファイルと転記元ファイルのチェック列をクリア
  Workbooks(OrgBook).Sheets("名簿").Activate   '原本ファイルをアクティブ
  With Workbooks(OrgBook).Sheets("名簿")
    lrow = .Range("A" & MEMBER_MAX).End(xlUp).Row
    Worksheets("名簿").Range(Cells(ROW_TOPDATA, COL_CHECK), Cells(lrow, COL_CHECK)).ClearContents
  End With
  
  Workbooks(pbook).Sheets("名簿").Activate   '転記ファイルをアクティブ
  With Workbooks(pbook).Sheets("名簿")
    lrow = .Range("A" & MEMBER_MAX).End(xlUp).Row
    Worksheets("名簿").Range(Cells(ROW_TOPDATA, COL_CHECK), Cells(lrow, COL_CHECK)).ClearContents
  End With
  
  '基本情報を転記するか確認
  If MsgBox("原本と転記元の基本情報（住所等）が異なる場合、" & vbNewLine & "基本情報（住所等）を確認しながら転記しますか？", vbYesNo + vbQuestion) = vbYes Then
    PostFlag = 1
  Else
    PostFlag = 2
  End If
 
 End Sub
  
'
'転記元ファイルのデータを原本ファイルに転記するプロシージャ
'  job　     ：処理内容
'  pbook     ：転記元ファイル名
'  srow, erow：転記開始行、終了行
Private Sub PostAllData(job As String, pbook As String, srow As Long, erow As Long)
 
  Dim pdata(4) As String   '転記元ファイルの転記データ文字列配列
  Dim prow As Long         '転記処理中の行
  Dim obj As Range         '検索結果
  Dim orow As Long         '検索結果の原本行
  Dim per As Integer       '進捗パーセンテージ
  Dim ecode As Integer     'エラーコード
   
  Application.Calculation = xlManual  '再計算を停止
  Application.ScreenUpdating = False  '再描画を停止
  
  prow = srow
  Do
    Workbooks(OrgBook).Sheets("名簿").Activate
    
    '転記処理内容のデータ有無を調べ、基本情報データを取得
    Call GetPostData(job, pbook, prow, pdata)
    
    'ステータスバーに進行状況表示
    per = (prow - srow + 1) / (erow - srow + 1) * 100
    Application.StatusBar = "◆" & pbook & "からの転記　" & prow - srow + 1 & "行目の ID：" _
       & pdata(1) & "を処理中 (" & per & " %)．．．"

    '転記元データのID番号をキーに、原本ファイルの転記対象を検索
    Workbooks(OrgBook).Sheets("名簿").Activate
    With Sheets("名簿").Range("C1:C" & MEMBER_MAX)
      Set obj = .Find(pdata(1), LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                 SearchOrder:=xlByColumns, MatchCase:=False, MatchByte:=False)
    End With
 
    If Not obj Is Nothing Then         '該当の｢ID｣が検索された場合
      orow = Range(obj.Address).Row
     
      '｢期、氏名｣が一致している場合
      If Cells(orow, COL_KI) = pdata(0) And Cells(orow, COL_NAME) = pdata(2) Then
        
        '転記対象データを原本に転記
        If pdata(3) <> "NO_DATA" Then
          Call PostData(job, orow, pdata)
        End If
        
        'コメントが変更されている場合、原本に追記
        If pdata(4) <> "" Then
          Call AddCmnt(orow, pdata(4))
        End If
    
        '基本情報が変更されている場合、判断後に転記
        If PostFlag = 1 Then
          Call PostBasicData(orow, pbook, prow)
        End If
        
        '基本情報が変更されている場合、転記対象を塗りつぶし
        If PostFlag = 2 Then   '転記処理中に「以後キャンセル」が押された場合を想定し、別のIf文に
          Call PaintBasicData(orow, pbook, prow)
        End If
 
      Else       '「期、氏名」が一致しない場合、処理継続かを確認
        ecode = 1
        Call IsPostCont(ecode, pdata(1), pbook, prow)
        If ecode <> 0 Then   '中止
          Exit Do
        Else
          Workbooks(pbook).Sheets("名簿").Cells(prow, COL_CHECK) = "異常"
        End If
      End If
    
    Else      '｢ID｣が検索できない場合、処理継続かを確認
      ecode = 2
      Call IsPostCont(ecode, pdata(1), pbook, prow)
      If ecode <> 0 Then   '中止
        Exit Do
      Else
        Workbooks(pbook).Sheets("名簿").Cells(prow, COL_CHECK) = "異常"
      End If
    End If
   
    prow = prow + 1
  Loop Until prow = erow + 1
                
  MsgBox "転記処理が完了しました。"
  
  Application.StatusBar = False         'ステータスバーの消去
  Application.Calculation = xlAutomatic '再計算を再開
  Application.ScreenUpdating = True     '再描画を再開
  Workbooks(OrgBook).Sheets("名簿").Activate   '原本ファイルをアクティブ
  
  Unload UserForm転記

End Sub
'
'転記対象のデータを読み取るとともに基本情報を取得するプロシージャ
'  job　   ：処理内容
'  pbook   ：転記元ファイル名
'  prow    ：転記対象行
'  pdata() ：転記データ
Private Sub GetPostData(job As String, pbook As String, prow As Long, pdata() As String)
   
  Dim pcol As Long
   
  '転記元の転記対象列にデータがある場合、処理に合わせて転記データを取得
  With Workbooks(pbook).Sheets("名簿")
    pdata(0) = .Cells(prow, COL_KI)
    pdata(1) = .Cells(prow, COL_ID)
    pdata(2) = .Cells(prow, COL_NAME)
    
    If job = "返信" And .Cells(prow, COL_CARD) <> "" Then      '返信データの取得
      pdata(3) = .Cells(prow, COL_CARD)
      .Cells(prow, COL_CHECK) = "済"
    ElseIf job = "電話" And .Cells(prow, COL_TEL) <> "" Then     '電話データの取得
      pdata(3) = .Cells(prow, COL_TEL)
      .Cells(prow, COL_CHECK) = "済"
    ElseIf job = "事前" And .Cells(prow, COL_ADVPAY) <> "" Then   '事前データの取得
      pdata(3) = .Cells(prow, COL_ADVPAY)
      .Cells(prow, COL_CHECK) = "済"
    ElseIf job = "本部" And .Cells(prow, COL_KAIHI0) <> "" Then     '本部データの取得
      pdata(3) = .Cells(prow, COL_KAIHI0)
      .Cells(prow, COL_CHECK) = "済"
    ElseIf job = "出席" And .Cells(prow, COL_RSLT) <> "" Then  '出席データの取得
      pdata(3) = .Cells(prow, COL_RSLT)
      .Cells(prow, COL_CHECK) = "済"
    ElseIf job = "当日" And .Cells(prow, COL_PAY) <> "" Then     '当日データの取得
      pdata(3) = .Cells(prow, COL_PAY)
      .Cells(prow, COL_CHECK) = "済"
    Else
      pdata(3) = "NO_DATA"
    End If

    If .Cells(prow, COL_COMMENT) <> "" Then         'コメントを取得
      pdata(4) = .Cells(prow, COL_COMMENT)
      .Cells(prow, COL_CHECK) = "済"
    Else
      pdata(4) = ""
    End If

    '基本情報を取得し、モジュール全体で使用する変数NewBasicDataにセット
    NewBasicData = ""
    For pcol = COL_KI To COL_JHSCHOOL
        NewBasicData = NewBasicData & .Cells(prow, pcol) & ","
    Next
    NewBasicData = NewBasicData & .Cells(prow, COL_COUPLE)
  
  End With
End Sub

'
'転記対象のデータを原本から検索し、原本に転記するプロシージャ
'  job　   ：処理内容
'  orow    ：原本ファイルの行数
'  pdata() ：転記対象データ（0:期、1:ID、2:氏名、3:転記データ、4:転記コメント）
Private Sub PostData(job As String, orow As Long, pdata() As String)

  With Workbooks(OrgBook).Sheets("名簿")
    
    Select Case job
      Case "返信"    '｢返信｣データからの転記
        If .Cells(orow, COL_CARD) <> pdata(3) Then
          .Cells(orow, COL_CARD) = pdata(3)
          .Cells(orow, COL_CHECK) = "済"
         End If
             
       Case "電話"    '｢電話｣データからの転記
         If .Cells(orow, COL_TEL) <> pdata(3) Then
           .Cells(orow, COL_TEL) = pdata(3)
           .Cells(orow, COL_CHECK) = "済"
         End If
                                  
       Case "事前"    '｢事前｣データからの転記
         If .Cells(orow, COL_ADVPAY) <> pdata(3) Then
           .Cells(orow, COL_ADVPAY) = pdata(3)
           .Cells(orow, COL_CHECK) = "済"
         End If
        
       Case "本部"    '｢本部｣データからの転記
         If .Cells(orow, COL_KAIHI0) <> pdata(3) Then
           .Cells(orow, COL_KAIHI0) = pdata(3)
           .Cells(orow, COL_CHECK) = "済"
         End If
            
       Case "出席"    '｢出席｣データからの転記
         If .Cells(orow, COL_RSLT) <> pdata(3) Then
           .Cells(orow, COL_RSLT) = pdata(3)
           .Cells(orow, COL_CHECK) = "済"
         End If
            
       Case "当日"    '｢当日｣データからの転記
         If .Cells(orow, COL_PAY) <> pdata(3) Then
           .Cells(orow, COL_PAY) = pdata(3)
           .Cells(orow, COL_CHECK) = "済"
         End If
    End Select
  End With
End Sub

'
'転記元データにコメントがある場合に、改行し追記するプロシージャ
'  orow    ：転記処理行
'  cmnt    ：転記するコメント
Private Sub AddCmnt(orow As Long, cmnt As String)

  With Workbooks(OrgBook).Sheets("名簿")
    If StrComp(cmnt, Cells(orow, COL_COMMENT)) = 0 Then       '既存コメントと同じ場合
      '何もしない
    ElseIf Cells(orow, COL_COMMENT) = "" Then                 '既存コメントが無い場合
      .Cells(orow, COL_COMMENT) = cmnt
      .Cells(orow, COL_CHECK) = "済"
    ElseIf InStr(cmnt, Cells(orow, COL_COMMENT)) > 0 Then     '既存コメントが含まれている場合
      .Cells(orow, COL_COMMENT) = cmnt
      .Cells(orow, COL_CHECK) = "済"
    Else                                                      '既存コメントが含まれていない場合
      .Cells(orow, COL_COMMENT) = Cells(orow, COL_COMMENT) & vbLf & cmnt
      .Cells(orow, COL_CHECK) = "済"
    End If
  
    .Cells(orow, COL_COMMENT).Font.Size = 8
  End With

End Sub
'
'転記対象の「期・氏名」が一致しない、「ID」が検索できない場合に
'エラーを表示し処理を継続するか確認するプロシージャ
'  ecode  ：エラーコード
'  id     ：エラーになったID
'  pbook  ：転記元ファイル名
'  prow   ：エラーになった行
Private Sub IsPostCont(ecode As Integer, id As String, pbook As String, prow As Long)

  Dim msg As String

  Application.ScreenUpdating = True     '再描画を再開
  Workbooks(pbook).Sheets("名簿").Activate
  Range(Cells(prow, COL_ID), Cells(prow, COL_ID)).EntireRow.Select
  Application.ScreenUpdating = False    '再描画を停止

  If ecode = 1 Then
    msg = "｢期・氏名｣が一致しません！ <ID: " & id & " >" & vbNewLine & _
               "以後のデータの処理を継続しますか？"
  ElseIf ecode = 2 Then
    msg = "<ID: " & id & " >を原本で検索できません。" & vbNewLine & _
                "以後のデータの処理を継続しますか？"
  End If
                
  'エラーメッセージを表示し、処理継続か確認
  If MsgBox(msg, 4 + 32, "処理継続") = 6 Then    '継続
    ecode = 0
  End If
End Sub

'
'転記元データと原本データの基本情報が異なる場合に、基本情報も修正するプロシージャ
'  orow  ：原本の行数
'  pbook    ：転記元ファイル名
'  prow     ：転記元ファイルの処理行
Private Sub PostBasicData(orow As Long, pbook As String, prow As Long)

  Dim ocol As Long

  '原本の基本情報をモジュール全体で使用する変数OrgBasicDataにセット
  With Workbooks(OrgBook).Sheets("名簿")

    OrgBasicData = ""
    For ocol = COL_KI To COL_JHSCHOOL
        OrgBasicData = OrgBasicData & .Cells(orow, ocol) & ","
    Next
    OrgBasicData = OrgBasicData & .Cells(orow, COL_COUPLE)
  
  End With

  '基本情報が異なる場合に基本情報転記フォームの呼び出し
  If OrgBasicData <> NewBasicData Then
     Workbooks(pbook).Sheets("名簿").Cells(prow, COL_CHECK) = "済"
     OrgRow = orow     '基本データ修正用に原本の現在行をモジュール全体で使用する変数OrgRowにセット
     UserForm基本データ転記.Show
  End If
  
  Unload UserForm基本データ転記

End Sub


'転記元データと原本データの基本情報が異なる場合に、転記元の基本情報をベタ塗りするプロシージャ
'  orow  ：原本の行数
'  pbook    ：転記元ファイル名
'  prow     ：転記元ファイルの処理行
Private Sub PaintBasicData(orow As Long, pbook As String, prow As Long)
  
  Dim odata As Variant  '原本の基本データ
  Dim pdata As Variant  '転記元の基本データ
  Dim ocol As Long
  Dim i As Integer

  With Workbooks(OrgBook).Sheets("名簿")

    OrgBasicData = ""
    For ocol = COL_KI To COL_JHSCHOOL
        OrgBasicData = OrgBasicData & .Cells(orow, ocol) & ","
    Next
    OrgBasicData = OrgBasicData & .Cells(orow, COL_COUPLE)
  
  End With

  '原本データと修正データを配列変数にセットする
  'データは[0],[1],･･･とセットされるので、データの扱いは注意
  odata = Split(OrgBasicData, ",")
  pdata = Split(NewBasicData, ",")
  
  '原本データが退会処理されている（氏名が「－」）場合は転記しない
  If odata(COL_NAME - 1) <> "－" Then
    With Workbooks(pbook).Sheets("名簿")
      
      'データを比較（「カナ氏名」～「夫婦」）
      For i = COL_KANA To COL_COUPLE
      
        '原本データと修正データが異なる場合、転記元のセルをローズに（転記はしない）
        If odata(i - 1) <> pdata(i - 1) Then
          Workbooks(pbook).Sheets("名簿").Activate
          .Range(Cells(prow, i), Cells(prow, i)).Interior.ColorIndex = 38
          Workbooks(OrgBook).Sheets("名簿").Activate
        End If
      Next i
    End With
  End If

End Sub


'
'「閉じる」ボタンが押されたときの処理
Private Sub CommandButton2_Click()
        
  Unload UserForm転記
    
End Sub




