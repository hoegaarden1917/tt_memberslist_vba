VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm��{�f�[�^�]�L 
   Caption         =   "��{�f�[�^�]�L"
   ClientHeight    =   8625
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   9372
   OleObjectBlob   =   "UserForm��{�f�[�^�]�L.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm��{�f�[�^�]�L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ���W���[�����ϐ��A�萔�̒�`
Dim OrgData, NewData As Variant
Const TEXT_BOX_NUM As Integer = 26
'




' �]�L�Ώۂ̃f�[�^��\������
Private Sub UserForm_Activate()
  
  Dim i, j, k As Integer
 
  Call ClearTextBox        '�S�Ă�TextBox������������

  '���{�f�[�^�ƏC���f�[�^��z��ϐ��ɃZ�b�g����
  '�f�[�^��[0],[1],����ƃZ�b�g�����̂ŁA�f�[�^�̈����͒���
  OrgData = Split(OrgBasicData, ",")
  NewData = Split(NewBasicData, ",")
  
  '���{�f�[�^���މ������Ă���i�������u�|�v�j�ꍇ�͓]�L���Ȃ�
  If OrgData(COL_NAME - 1) = "�|" Then
    Unload Me
    DoEvents        '��ʕ`�ʂ����邽�߂ɏ�����OS�ɓn��
  
  Else
  
    '�\���Ώۂ́uID�v�Ɓu�����v��\������
    TextBox25.Value = OrgData(COL_ID - 1)
    TextBox26.Value = OrgData(COL_NAME - 1)
  
    '�f�[�^��\������i�u�J�i�����v�`�u�v�w�v�j
    i = 1
    For j = COL_KANA To COL_COUPLE
      k = i + 1
      Me("TextBox" & i).Value = OrgData(j - 1)
      Me("TextBox" & k).Value = NewData(j - 1)
       
      '���{�f�[�^�ƏC���f�[�^���قȂ�ꍇ�A�w�i��Ԃɂ���
      If OrgData(j - 1) <> NewData(j - 1) Then
        Controls("TextBox" & k).BackColor = vbRed
      End If
    
      i = i + 2
    Next j
  End If

End Sub

' �S�Ă�TextBox���N���A����
Private Sub ClearTextBox()
  
  Dim i As Integer

  For i = 1 To TEXT_BOX_NUM
    Me("TextBox" & i).Value = ""
    Controls("TextBox" & i).BackColor = vbWhite
  Next
  DoEvents        '��ʕ`�ʂ����邽�߂ɏ�����OS�ɓn��
End Sub

' �u�]�L�v�{�^���������ꂽ�Ƃ��̏���
Private Sub CommandButton1_Click()

  Dim i As Integer

  '�@�f�[�^��\������i�u�J�i�����v�`�u�v�w�v�j
  For i = COL_KANA To COL_COUPLE
    If OrgData(i - 1) <> NewData(i - 1) Then
      Workbooks(OrgBook).Sheets("����").Cells(OrgRow, i) = NewData(i - 1)
    End If
  Next
  Workbooks(OrgBook).Sheets("����").Cells(OrgRow, COL_CHECK) = "��"
  
  Call ClearTextBox        '�S�Ă�TextBox������������
  Unload Me
  DoEvents        '��ʕ`�ʂ����邽�߂ɏ�����OS�ɓn��
End Sub




' �u�L�����Z���v�{�^���������ꂽ�Ƃ�
Private Sub CommandButton2_Click()

  Call ClearTextBox        '�S�Ă�TextBox������������
  Unload Me
  DoEvents        '��ʕ`�ʂ����邽�߂ɏ�����OS�ɓn��
End Sub


' �u�Ȍ�S�ăL�����Z���v�{�^���������ꂽ�Ƃ�
Private Sub CommandButton3_Click()
  
  PostFlag = 2             '�h��Ԃ���
  Call ClearTextBox        '�S�Ă�TextBox������������
  Unload Me
  DoEvents                 '��ʕ`�ʂ����邽�߂ɏ�����OS�ɓn��
End Sub

