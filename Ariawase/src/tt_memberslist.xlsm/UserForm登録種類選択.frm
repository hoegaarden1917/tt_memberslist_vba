VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm�o�^��ޑI�� 
   Caption         =   "�o�^"
   ClientHeight    =   1005
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4620
   OleObjectBlob   =   "UserForm�o�^��ޑI��.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm�o�^��ޑI��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' 2010.7.3�@���{�t�@�C���̖���V�[�g�̗�̕��т̌������ɂ��C��
Private Sub CommandButton1_Click()

  '�I�v�V�����{�^��1�i����j���N���b�N�����ꍇ
  If OptionButton1.Value = True Then
    
    '�����ɕK�v�ȃt�@�C�����J����Ă��邩�m�F����
    If IsBookOpen("�X�֔ԍ��ް��y�S���Łz.xls") = False Then
      MsgBox "�u�X�֔ԍ��ް��y�S���Łz.xls�v���J����Ă��܂���I" _
          & vbNewLine & "�J���Ă����蒼���ĉ������B"
      End
    End If
    If IsBookOpen("�������}���y���މ�҈ꗗ�z.xls") = False Then
      MsgBox "�u�������}���y���މ�҈ꗗ�z.xls�v���J����Ă��܂���I" _
          & vbNewLine & "�J���Ă����蒼���ĉ������B"
      End
    End If
   
    With UserForm�o�^            '�o�^��ʂ�\��
      .Caption = "����o�^"      '�L���v�V�����������o�^��ɕύX
      .TextBox1.Enabled = True   '�e�L�X�g�{�b�N�X1����͉\��
      .Show
    End With
 
  '�I�v�V�����{�^��2�i�ύX�j���N���b�N�����ꍇ
  ElseIf OptionButton2.Value = True Then
    With UserForm����       '������ʂ�\��
      .Caption = "����"     '���̂𢌟�����
      .CommandButton3.Caption = "�ύX�o�^"   '�R�}���h�{�^��3�̖��̂�ύX�o�^���
      .CommandButton3.Enabled = True         '�R�}���h�{�^��3���g�p�\��
      .Show
    End With
  
  '�I�v�V�����{�^��3�i�މ�j���N���b�N�����ꍇ
  ElseIf OptionButton3.Value = True Then
   
    '�����ɕK�v�ȃt�@�C�����J����Ă��邩�m�F����
    If IsBookOpen("�������}���y���މ�҈ꗗ�z.xls") = False Then
       MsgBox "�u�������}���y���މ�҈ꗗ�z.xls�v���J����Ă��܂���I" _
          & vbNewLine & "�J���Ă����蒼���ĉ������B"
       End
    End If
   
    With UserForm����       '������ʂ�\��
      .Caption = "����"     '���̂𢌟�����
      .CommandButton3.Caption = "�މ�o�^"  '�R�}���h�{�^��3�̖��̂�މ�o�^���
      .CommandButton3.Enabled = True        '�R�}���h�{�^��3���g�p�\��
      .Show
    End With
  
  '�I�v�V�����{�^��4�i�Z���ꊇ�Ɖ�j���N���b�N�����ꍇ
  ElseIf OptionButton4.Value = True Then
    Call ChkAllAddr
  End If

End Sub

'�I�v�V�����{�^��4�i�Z���ꊇ�Ɖ�j�������ꂽ�Ƃ��̃v���V�[�W��
Private Sub ChkAllAddr()

  Dim msg As String
  Dim nrow As Long   '�s1
  Dim zip1 As String '��
  Dim zip2 As String '��2
  Dim obj As Range   '����1
  Dim zrow As Long   '�s2
  Dim addrs(3) As String
   
  '�����ɕK�v�ȃt�@�C�����J����Ă��邩�m�F����
  If IsBookOpen("�X�֔ԍ��ް��y�S���Łz.xls") = False Then
    MsgBox "�u�X�֔ԍ��ް��y�S���Łz.xls�v���J����Ă��܂���I" _
       & vbNewLine & "�J���Ă����蒼���ĉ������B"
    End
  End If
 
  msg = "�I�����ꂽ�s����A�X�֔ԍ��Ɋ�A�Z���ꊇ�ƍ����s���܂����H"  '���̃��b�Z�[�W��\�����A
  If MsgBox(msg, 4 + 32, "���ˏZ���ꊇ�ƍ�") = 6 Then  '��͂�����N���b�N����΁A
  
    nrow = ActiveCell.Row    '�A�N�e�B�u�Z���̍s����\���ϐ���s1��Ƃ���B
   
    Do '�ȉ��̎菇���J��Ԃ��B
      zip1 = Left(Cells(nrow, COL_ZIP), 3) & Right(Cells(nrow, COL_ZIP), 4)   '�I������Ă���s�ɋL�ڂ���Ă���X�֔ԍ��̢-����폜�����������ϐ��Zip1��Ƃ���B
      zip2 = Cells(nrow, COL_ZIP)     '�I������Ă���s�ɋL�ڂ���Ă��镶����\���ϐ��𢁧2��Ƃ���B
    
      If zip2 = "�|" Or zip2 = "" Then         '�ϐ��Zip2�����|��A�܂��͋󔒂ł���΁A�������Ȃ��B
      Else    '�ϐ��Zip2�����|��A�܂��͋󔒂łȂ���΁A�ȉ��̏������s���B
    
        '�X�֔ԍ�����������B�ϐ��Obj��͗X�֔ԍ��������ł������ǂ�����\���B
        With Workbooks("�X�֔ԍ��ް��y�S���Łz.xls").Sheets("�X�֔ԍ�1").Range("B1:B65001")
          Set obj = .Find(zip1, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                    MatchCase:=False, MatchByte:=False)
        End With
   
        If Not obj Is Nothing Then      '�V�[�g��X�֔ԍ�1��ŗX�֔ԍ������������΁A
          With Workbooks("�X�֔ԍ��ް��y�S���Łz.xls").Sheets("�X�֔ԍ�1")
            zrow = .Range(obj.Address).Row
            addrs(0) = .Cells(zrow, 3)
            addrs(1) = .Cells(zrow, 4)
            addrs(2) = .Cells(zrow, 5)
          End With
      
        Else                             '�V�[�g��X�֔ԍ�1��ŗX�֔ԍ�����������Ȃ���΁A
          With Workbooks("�X�֔ԍ��ް��y�S���Łz.xls").Sheets("�X�֔ԍ�2").Range("B1:B65001")
            Set obj = .Find(zip1, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                      MatchCase:=False, MatchByte:=False)
          End With
      
          If obj Is Nothing Then     '�V�[�g��X�֔ԍ�2��ŗX�֔ԍ����������Ȃ���΁A�Y���s��I�����A�ȉ��̃��b�Z�[�W��\�����A�ƍ��𒆒f����B
            Cells(nrow, 7).Select
            MsgBox "�Y���̗X�֔ԍ��͌�������܂���ł����I"
            Exit Do
      
          Else                         '�V�[�g��X�֔ԍ�2��ŗX�֔ԍ������������΁A
            With Workbooks("�X�֔ԍ��ް��y�S���Łz.xls").Sheets("�X�֔ԍ�2")
              zrow = .Range(obj.Address).Row
              addrs(0) = .Cells(zrow, 3)
              addrs(1) = .Cells(zrow, 4)
              addrs(2) = .Cells(zrow, 5)
            End With
          End If
        End If
        
        'H��ƕϐ���Z��1��AI��ƕϐ���Z��2��AJ��ƕϐ���Z��3�����������΁A�������Ȃ��B
        If Cells(nrow, COL_ADDR1) = addrs(0) And Cells(nrow, COL_ADDR2) = addrs(1) _
            And Cells(nrow, COL_ADDR3) = addrs(2) Then
        
        Else  '�������Ȃ���΁A�Y���s��I�����A�ȉ��̃��b�Z�[�W��\�����A�ƍ��𒆒f����B
          Cells(nrow, COL_ZIP).Select
          MsgBox "�Z������v���܂���I�w" & addrs(0) & addrs(1) & addrs(2) & "�x"
          Exit Do
        End If
      End If
     
      nrow = nrow + 1
            
      If Cells(nrow, 1) = "" Then  '�����AA�񂪋󔒂ł���΁A�ȉ��̃��b�Z�[�W��\�����āA��Ƃ𒆎~����B
        Cells(nrow, COL_ZIP).Select
        MsgBox "�ƍ����I�����܂����I"
        Exit Do
      Else
      End If
     
    Loop Until nrow = 5000   '���[�v��5000��J��Ԃ��B
  End If
End Sub

Private Sub CommandButton2_Click()
  UserForm�o�^��ޑI��.Hide
End Sub


