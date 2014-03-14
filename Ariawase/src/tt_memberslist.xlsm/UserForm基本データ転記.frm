VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm基本データ転記 
   Caption         =   "基本データ転記"
   ClientHeight    =   8625
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   9372
   OleObjectBlob   =   "UserForm基本データ転記.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm基本データ転記"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' モジュール内変数、定数の定義
Dim OrgData, NewData As Variant
Const TEXT_BOX_NUM As Integer = 26
'




' 転記対象のデータを表示する
Private Sub UserForm_Activate()
  
  Dim i, j, k As Integer
 
  Call ClearTextBox        '全てのTextBoxを初期化する

  '原本データと修正データを配列変数にセットする
  'データは[0],[1],･･･とセットされるので、データの扱いは注意
  OrgData = Split(OrgBasicData, ",")
  NewData = Split(NewBasicData, ",")
  
  '原本データが退会処理されている（氏名が「－」）場合は転記しない
  If OrgData(COL_NAME - 1) = "－" Then
    Unload Me
    DoEvents        '画面描写させるために処理をOSに渡す
  
  Else
  
    '表示対象の「ID」と「氏名」を表示する
    TextBox25.Value = OrgData(COL_ID - 1)
    TextBox26.Value = OrgData(COL_NAME - 1)
  
    'データを表示する（「カナ氏名」～「夫婦」）
    i = 1
    For j = COL_KANA To COL_COUPLE
      k = i + 1
      Me("TextBox" & i).Value = OrgData(j - 1)
      Me("TextBox" & k).Value = NewData(j - 1)
       
      '原本データと修正データが異なる場合、背景を赤にする
      If OrgData(j - 1) <> NewData(j - 1) Then
        Controls("TextBox" & k).BackColor = vbRed
      End If
    
      i = i + 2
    Next j
  End If

End Sub

' 全てのTextBoxをクリアする
Private Sub ClearTextBox()
  
  Dim i As Integer

  For i = 1 To TEXT_BOX_NUM
    Me("TextBox" & i).Value = ""
    Controls("TextBox" & i).BackColor = vbWhite
  Next
  DoEvents        '画面描写させるために処理をOSに渡す
End Sub

' 「転記」ボタンが押されたときの処理
Private Sub CommandButton1_Click()

  Dim i As Integer

  '　データを表示する（「カナ氏名」～「夫婦」）
  For i = COL_KANA To COL_COUPLE
    If OrgData(i - 1) <> NewData(i - 1) Then
      Workbooks(OrgBook).Sheets("名簿").Cells(OrgRow, i) = NewData(i - 1)
    End If
  Next
  Workbooks(OrgBook).Sheets("名簿").Cells(OrgRow, COL_CHECK) = "済"
  
  Call ClearTextBox        '全てのTextBoxを初期化する
  Unload Me
  DoEvents        '画面描写させるために処理をOSに渡す
End Sub




' 「キャンセル」ボタンが押されたとき
Private Sub CommandButton2_Click()

  Call ClearTextBox        '全てのTextBoxを初期化する
  Unload Me
  DoEvents        '画面描写させるために処理をOSに渡す
End Sub


' 「以後全てキャンセル」ボタンが押されたとき
Private Sub CommandButton3_Click()
  
  PostFlag = 2             '塗りつぶしに
  Call ClearTextBox        '全てのTextBoxを初期化する
  Unload Me
  DoEvents                 '画面描写させるために処理をOSに渡す
End Sub

