Attribute VB_Name = "Module1"
Option Explicit

'�ȉ��̒萔�͓��Ԋ���㎞�Ɍ��������s���萔
Const MEIBO_PASSWORD As String = "touchiku84"   '�W�q�S�����ɕ�����������̓ǂݎ��p�X���[�h


'�ȉ��̒萔�̓��W���[���S�̂Ŏg���萔
'����̗��ǉ������ꍇ�́A��̈ʒu�������萔���ς�邽�߁A���؂��K�v�ɂȂ�̂ŗv���ӁI

'--- ��{�I�Ȓ萔
Public Const MEMBER_MAX As Integer = 10000      '����̍ő吔
Public Const DOUKI_MAX As Integer = 300         '��������̍ő吔

'--- ����V�[�g�Ɋւ���萔�i�V�[�g�̍s���̒ǉ��폜���s�����ꍇ�͌������v�I�j
Public Const ROW_TOPTITLE As Integer = 2 '����V�[�g�̃^�C�g�����܂ސ擪�s
Public Const ROW_TOPDATA As Integer = 5  '����V�[�g�̃f�[�^�̐擪�s
'
Public Const COL_KI As Integer = 1           '�u���v�̗�
Public Const COL_CLASS As Integer = 2        '�u��ށv�̗�
Public Const COL_ID As Integer = 3           '�uID�v�̗�
Public Const COL_NAME As Integer = 4         '�u�����v�̗�
Public Const COL_KANA As Integer = 5         '�u�J�i�����v�̗�
Public Const COL_SEX As Integer = 6          '�u���ʁv�̗�
Public Const COL_ZIP As Integer = 7          '�u���v�̗�
Public Const COL_ADDR1 As Integer = 8        '�u�Z��1�v�̗�
Public Const COL_ADDR2 As Integer = 9        '�u�Z��2�v�̗�
Public Const COL_ADDR3 As Integer = 10       '�u�Z��3�v�̗�
Public Const COL_ADDR4 As Integer = 11       '�u�Z��4�v�̗�
Public Const COL_TELNO As Integer = 12       '�u�d�b�ԍ��v�̗�
Public Const COL_EMAIL As Integer = 13       '�u���[���A�h���X�v�̗�
Public Const COL_BUKATSU As Integer = 14     '�u���������v�̗�
Public Const COL_JHSCHOOL As Integer = 15    '�u�o�g���w�v�̗�
Public Const COL_COUPLE As Integer = 16      '�u�v�w�v�̗�
Public Const COL_KAIHI3 As Integer = 17      '�u�N���3�N�O�v�̗�
Public Const COL_KAIHI2 As Integer = 18      '�u�N���2�N�O�v�̗�
Public Const COL_KAIHI1 As Integer = 19      '�u�N���1�N�O�v�̗�
Public Const COL_KAIHI0 As Integer = 20      '�u�N���i���N�x�j�v�̗�
Public Const COL_KANJI As Integer = 21       '�u�������v�̗�
Public Const COL_REMARK As Integer = 22      '�u���l�v�̗�
Public Const COL_CHECK As Integer = 23       '�u�]�L�`�F�b�N���v�̗�

Public Const COL_RSLT5 As Integer = 24       '�u���e�����5�N�O�v�̗�
Public Const COL_RSLT4 As Integer = 25       '�u���e�����4�N�O�v�̗�
Public Const COL_RSLT3 As Integer = 26       '�u���e�����3�N�O�v�̗�
Public Const COL_RSLT2 As Integer = 27       '�u���e�����2�N�O�v�̗�
Public Const COL_RSLT1 As Integer = 28       '�u���e�����1�N�O�v�̗�
Public Const COL_RSLT0 As Integer = 29       '�u���e����сi���N�x�j�v�̗�

Public Const COL_CARD As Integer = 30        '�u�i���N�x���e��̏o���j�ԐM�v�̗�
Public Const COL_TEL As Integer = 31         '�u�i���N�x���e��̏o���j�d�b�v�̗�
Public Const COL_KICARD As Integer = 32      '�u�i���N�x���e��̏o���j���ʕԐM�v�̗�
Public Const COL_KITEL As Integer = 33       '�u�i���N�x���e��̏o���j���ʓd�b�v�̗�
Public Const COL_PLAN As Integer = 34        '�u�i���N�x���e��̏o���j�o�\��v�̗�
Public Const COL_RSLT As Integer = 35        '�u�i���N�x���e��̏o���j�����v�̗�
Public Const COL_KIRSLT As Integer = 36      '�u�i���N�x���e��̏o���j���ʓ����v�̗�
Public Const COL_ADVPAY As Integer = 37      '�u�i���N�x���e��̓����j���O�v�̗�
Public Const COL_PAY As Integer = 38         '�u�i���N�x���e��̓����j�����v�̗�
Public Const COL_KIPAY As Integer = 39       '�u�i���N�x���e��̓����j���ʓ����v�̗�
Public Const COL_COMMENT As Integer = 40        '�u�R�����g�v�̗�

'--- ���ʏo���V�[�g�E���ʓ����V�[�g���ʂ̒萔�i�V�[�g�̍s���̒ǉ��폜���s�����ꍇ�͌������v�I�j
Const ROW_KITOPDATA As Integer = 6    '�f�[�^�̐擪�s��
Const COL_TANTO_MAX As Integer = 100       '�W�q�S���̍ő�l
Const KI_MAX As Integer = 200         '���̍ő�l


'
'--- �o���V�[�g�Ɋւ���萔�i�V�[�g�̍s���̒ǉ��폜���s�����ꍇ�͌������v�I�j
Const ROW_RSLTSUM1 As Integer = 4       '���v�i��j�̍s��
Const ROW_RSLTSUM2 As Integer = 5       '���v�i���j�̍s��

'
Const COL_CARD_OK As Integer = 2      '�u�ԐM�n�K�L�̏o�v�̗�
Const COL_CARD_NG As Integer = 3      '�u�ԐM�n�K�L�̌��v�̗�
Const COL_CARD_ERROR As Integer = 4   '�u�ԐM�n�K�L�̕s���v�̗�
Const COL_CARD_NOFIX As Integer = 5   '�u�ԐM�n�K�L�̖���v�̗�
Const COL_CARD_SUM As Integer = 6     '�u�ԐM�n�K�L�̌v�v�̗�
Const COL_EMAIL_OK As Integer = 7     '�u���[���EHP�̏o�v�̗�
Const COL_EMAIL_NG As Integer = 8     '�u���[���EHP�̌��v�̗�
Const COL_EMAIL_NOFIX As Integer = 9  '�u���[���EHP�̖���v�̗�
Const COL_EMAIL_SUM As Integer = 10   '�u���[���EHP�̌v�v�̗�
Const COL_TEL_OK As Integer = 11      '�u�d�b�̏o�v�̗�
Const COL_TEL_NG As Integer = 12      '�u�d�b�̌��v�̗�
Const COL_TEL_ERROR As Integer = 13   '�u�d�b�̕s�ʁv�̗�
Const COL_TEL_NOFIX As Integer = 14   '�u�d�b�̖���v�̗�
Const COL_TEL_SUM As Integer = 15     '�u�d�b�̌v�v�̗�
Const COL_SUM_OK As Integer = 16      '�u�o�ȘA���v�̗�

Const COL_OKPAY As Integer = 18       '�u�o�Ȃ����O�����v�̗�
Const COL_OTHERPAY As Integer = 19    '�u�o�ȈȊO�����O�����v�̗�
Const COL_SUMPAY As Integer = 20      '�u���O�����̌v�v�̗�
  
Const COL_SUM_ALL As Integer = 22     '�u�o�ȗ\��v�̗�

Const COL_RSLT_OK As Integer = 24     '�u�����o�ȁv�̗�
Const COL_RSLT_PAYNG As Integer = 25  '�u�����������ȁv�̗�
Const COL_RSLT_SUM As Integer = 26    '�u�����o�ȁv�̗�

Const COL_MEMBER As Integer = 28      '�u������v�̗�
Const COL_TANTO As Integer = 30       '�u�W�q�S���v�̗�

'== Mar.9,2014 Yuji Ogihara ==
'2007 �}�N���L���`���ւ̕ϊ�����уt�@�C�����̂̒萔��
Const SAGYOHYO_FILENAME As String = "�����V�[���E���D�E�o�Ȏ҈ꗗ�\�쐬.xlsm" '��ƕ\�t�@�C������

'
'--- ���W���[���S�̂Ŏg�p����ϐ�
'    �i�W�q�S�������C�������f�[�^��]�L���邽�߂Ɏg�p����j
Public OrgBook As String            '���{�̃t�@�C����
Public OrgBasicData As String       '���{�̊�{�f�[�^
Public OrgRow As Long               '�C�����̌��{�V�[�g�ł̍s��
Public NewBasicData As String       '�C�����̊�{�f�[�^
Public PostFlag As Integer          '��{�f�[�^�C�����s�����ۂ��̃t���O



'


'�t�@�C�����J����Ă��邩�𒲂ׂ�֐�
' �����@�@bName �F�J����Ă��邩���ׂ�BOOK��
' �߂�l�@True  �F�J����Ă���A�@False�F�J����Ă��Ȃ�
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



'�u����v�V�[�g�́u�����v�{�^���������ꂽ�Ƃ��̏���
Sub ������ʕ\��()
    
   With UserForm����
     .CommandButton3.Enabled = False
     .Show
   End With
   
End Sub

'�u����v�V�[�g�́u�o�^�v�{�^���������ꂽ�Ƃ��̏���
'�i����E�ύX�E�މ�E�����Z���ƍ��̃T�u���j���[��\���j
Sub �o�^��ޑI����ʕ\��()

    UserForm�o�^��ޑI��.Show
End Sub

'�u����v�V�[�g�́u�}��\���v�{�^���������ꂽ�Ƃ��̏���
Sub �}���ʕ\��()

    UserForm�}��.Show
End Sub

'�u����v�V�[�g�́u�]�L�v�{�^���������ꂽ�Ƃ��̏���
'�i�N���A���e��ւ̐\���ݏ󋵓��̏���]�L���邽�߂̏����E�t�@�C���̎w��j
Sub �]�L��ʕ\��()

    UserForm�]�L.Show
End Sub

Sub �\�L�m�F()

 Dim nrow As Long
 Dim lrow As Long
 Dim i As Integer
 
 lrow = Cells(MEMBER_MAX, COL_KI).End(xlUp).Row

 For nrow = ROW_TOPDATA To lrow

   If Cells(nrow, COL_NAME) = "�|" Or Cells(nrow, COL_NAME) = "-" Or Cells(nrow, COL_NAME) = "" Then
    
     For i = COL_NAME To COL_EMAIL
       Cells(nrow, i) = "�|"
     Next i
     Range(Cells(nrow, COL_NAME), Cells(nrow, COL_EMAIL)).HorizontalAlignment = xlCenter
     Range(Cells(nrow, COL_BUKATSU), Cells(nrow, COL_COMMENT)).ClearContents
   End If
 
   If Cells(nrow, COL_COUPLE) <> "" Then
     Range(Cells(nrow, COL_KI), Cells(nrow, COL_COMMENT)).Interior.ColorIndex = 43
   End If
   
 
 Next nrow
    
End Sub

'���ʂ̍��e��o���󋵂��W�v���邽�߂̏����A�B����ɃZ�b�g����}�N��
' �o�\�F�n�K�L�A���[���A�d�b�ŏo�ȂƉ񓚂��������ꍇ
' �����F�n�K�L�A���[���A�d�b�ł͌��Ȃ܂��͉񓚂��������A�Q����̓������������ꍇ
Sub �o���m�F()
     
  Dim nrow As Long     '�������s���s
  Dim lrow As Long     '�ŏI�s
  Dim ki As String     '����������̕�����
  Dim card As String   '�u�ԐM�v��̕�����
  Dim tel As String    '�u�d�b�v��̕�����
  Dim rslt As String   '�u�����v��̕�����
  Dim apay As String   '�u���O�����v��̕�����
  Dim per As Integer   '��������
     
  Application.ScreenUpdating = False      '�`��𒆎~
  ActiveSheet.EnableCalculation = False   '�Čv�Z�𒆎~

  With Sheets("����")
 
    nrow = ROW_TOPDATA   '�����J�n�s���Z�b�g
   
    '�ŏI�s���u���v���MEMBER_MAX�s�����ɒ���
    lrow = .Cells(MEMBER_MAX, COL_KI).End(xlUp).Row
   
    '�B����̊��ʕԐM�A���ʓd�b�A�o�\��A���ʓ����o�Ȃ��N���A
    .Range(.Cells(nrow, COL_KICARD), .Cells(lrow, COL_PLAN)).ClearContents
    .Range(.Cells(nrow, COL_KIRSLT), .Cells(lrow, COL_KIRSLT)).ClearContents

    Do  '�ȉ��̏������J��Ԃ�
     
      per = (nrow - ROW_TOPDATA) / (lrow - ROW_TOPDATA) * 100
      Application.StatusBar = "��" & nrow - ROW_TOPDATA + 1 & "�s�� (" & per & " %) �̏������D�D�D"
 
      '���݂̍s�̊��A�ԐM�A�d�b�A�����A���O�����̕������擾
      ki = .Cells(nrow, COL_KI).Text
      card = .Cells(nrow, COL_CARD).Text
      tel = .Cells(nrow, COL_TEL).Text
      rslt = .Cells(nrow, COL_RSLT).Text
      apay = .Cells(nrow, COL_ADVPAY).Text
    
     '�����̕����񂪋󔒂̏ꍇ�͏����𒆎~
      If ki = "" Then
        Exit Do
     
      Else
      
        '��ԐM��̕����񂪋󔒂łȂ���΁A�u���ʕԐM�v��Ɂu���v+�u�ԐM�v�̕������Z�b�g
        If card <> "" Then
          .Cells(nrow, COL_KICARD) = ki & card
          '��ԐM���1�����ڂ���o��̏ꍇ�́A�u�o�\��v��Ɂu���v+"�o�\"���Z�b�g
          If Left(card, 1) = "�o" Then
            .Cells(nrow, COL_PLAN) = ki & "�o�\"
          End If
        
        End If
   
        '�u�d�b�v�̕����񂪋󔒂łȂ���΁A�u���ʓd�b�v��ɢ���+��d�b��̕������Z�b�g
        If tel <> "" Then
          .Cells(nrow, COL_KITEL) = ki & tel
          '��d�b�����o��̏ꍇ�́A�u�o�\��v��Ɂu���v+"�o�\"���Z�b�g
          If Left(card, 1) = "�o" Or Left(tel, 1) = "�o" Then
            .Cells(nrow, COL_PLAN) = ki & "�o�\"
          End If
        End If

        '�u�o�\��v���󔒂Łu���O�����v���󔒂łȂ���΁A�u�o�\��v��ɢ���+"����"�̕������Z�b�g
        If .Cells(nrow, COL_PLAN) = "" And apay <> "" Then
          .Cells(nrow, COL_PLAN) = ki & "����"
        End If

        '�������̕����񂪋󔒂łȂ���΁A����+�������̕������Z�b�g
        If rslt <> "" Then
          .Cells(nrow, COL_KIRSLT) = ki & rslt
        ElseIf apay <> "" Then
          .Cells(nrow, COL_KIRSLT) = ki & "����"
        End If
      End If
 
      nrow = nrow + 1  '���̍s��
     
    Loop Until nrow = lrow + 1  '�ŏI�s+1�ɂȂ�܂ŌJ��Ԃ�
    
    Application.StatusBar = False   '�X�e�[�^�X�o�[�̏���
   
  End With
   
  ActiveSheet.EnableCalculation = True   '�Čv�Z���ĊJ
  Application.ScreenUpdating = True      '�`����ĊJ
   
End Sub


'���ʂ̍��e��Q��������󋵂��W�v���邽�߂̏����A�B����ɃZ�b�g����}�N��
' XX���O���Z�b�g
Sub ���ʓ����W�v()
     
  Dim nrow As Long      '�������s���s
  Dim lrow As Long      '�ŏI�s
  Dim ki As String      '����������̕�����
  Dim apay As String    '�u���O�����v��̕�����
  Dim pay As String     '�u���������v��̕�����
  Dim per As Integer    '��������
     
  Application.ScreenUpdating = False      '�`��𒆎~
  ActiveSheet.EnableCalculation = False   '�Čv�Z�𒆎~
  
  With Sheets("����")
    nrow = ROW_TOPDATA
    
    '�ŏI�s���u���v���MEMBER_MAX�s�����ɒ���
    lrow = .Cells(MEMBER_MAX, COL_KI).End(xlUp).Row
    
    '���ʓ����̗���폜
    .Range(.Cells(nrow, COL_KIPAY), .Cells(lrow, COL_KIPAY)).ClearContents
  
    Do  '�ȉ��̏������J��Ԃ��B
      
      per = (nrow - ROW_TOPDATA) / (lrow - ROW_TOPDATA) * 100
      Application.StatusBar = "��" & nrow - ROW_TOPDATA + 1 & "�s�� (" & per & " %) �̏������D�D�D"
      
      '���݂̍s�̊��A���O�����̕����A���������̕������擾
      ki = .Cells(nrow, COL_KI).Text
      apay = .Cells(nrow, COL_ADVPAY).Text
      pay = .Cells(nrow, COL_PAY).Text
        
      '�����̕����񂪋󔒂̏ꍇ�͏����𒆎~
      If ki = "" Then
        Exit Do
      Else
        '����O���ࣂ̕����񂪋󔒂łȂ���΁A�����Ƣ���O���ࣂ̕���������������̂��Z�b�g
        If apay <> "" Then
          .Cells(nrow, COL_KIPAY) = ki & apay
        End If
        
        '�����������̕����񂪋󔒂łȂ���΁A�����Ƣ�������ࣂ̕���������������̂��Z�b�g
        If pay <> "" Then
          .Cells(nrow, COL_KIPAY) = ki & pay
        End If
      End If
 
      nrow = nrow + 1  '���̍s��
     
    Loop Until nrow = lrow + 1  '�ŏI�s+1�ɂȂ�܂ŌJ��Ԃ�
   
    Application.StatusBar = False   '�X�e�[�^�X�o�[�̏���
  End With
  
  ActiveSheet.EnableCalculation = True   '�Čv�Z���ĊJ
  Application.ScreenUpdating = True      '�`����ĊJ

End Sub
'
'�u���ʏo���W�v�v�V�[�g�̊e�Z���ɏW�v�̂��߂̊֐��𖄂ߍ��ރ}�N��
Sub �o���m�F�֐��\�t()

  Dim nrow As Long      '�������s���s
  Dim mlrow As Long     '����V�[�g�̍ŏI�s
  Dim slrow As Long     '���ʏo���W�v�V�[�g�̍ŏI�s
  Dim card As String    '�u�ԐM�v��̕�����
  Dim tel As String     '�u�d�b�v��̕�����
  Dim plan As String    '�u�o�\��v��̕�����
  Dim rslt As String    '�u�����v��̕�����
  Dim apay As String    '�u���O�����v��̕�����
  Dim kipay As String   '�u���ʓ����v��̕�����
  Dim ki As String      '�u���v��̕�����
  Dim mmbr1 As String   '�u������v��̕�����
  Dim mmbr2 As String   '�u������v��̕�����

  With Sheets("����")
    mlrow = .Cells(MEMBER_MAX, COL_ID).End(xlUp).Row
  End With
 
  Call SetKiNo("���ʏo���W�v")    '�u���v�̗��ݒ�
  
  With Sheets("���ʏo���W�v")
    
    '���ʏo���W�v�V�[�g�̍ŏI�s���擾
    slrow = .Cells(MEMBER_MAX, COL_KI).End(xlUp).Row
  
    '����V�[�g�̕ԐM�A�d�b�A�����̓o�^�󋵂��獇�v���W�v���邽�߂̌v�Z���̕������ݒ�
    card = "����!R" & ROW_TOPDATA & "C" & COL_CARD & ":R" & mlrow & "C" & COL_CARD
    tel = "����!R" & ROW_TOPDATA & "C" & COL_TEL & ":R" & mlrow & "C" & COL_TEL
    rslt = "����!R" & ROW_TOPDATA & "C" & COL_RSLT & ":R" & mlrow & "C" & COL_RSLT
 
    '���v�𖼕�V�[�g�̕ԐM��d�b�̏󋵂���W�v����֐������v�̏�̗�ɖ��ߍ���
 
    '����́u�ԐM�v�񂩂�n�K�L�ԐM�̍��v���W�v
    Range("B4").Formula = "=COUNTIF(" & card & ",""�o�n"")"
    Range("C4").Formula = "=COUNTIF(" & card & ",""���n"")"
    Range("D4").Formula = "=COUNTIF(" & card & ",""�s��"")"
    Range("E4").Formula = "=COUNTIF(" & card & ",""���n"")"
    Range("F4").Formula = "=SUM(RC[-4]:RC[-1])"
  
    '����́u�ԐM�v�񂩂�d�q���[���EHP�o�^�̍��v���W�v
    Range("G4").Formula = "=COUNTIF(" & card & ",""�o��"")"
    Range("H4").Formula = "=COUNTIF(" & card & ",""����"")"
    Range("I4").Formula = "=COUNTIF(" & card & ",""����"")"
    Range("J4").Formula = "=SUM(RC[-3]:RC[-1])"
 
    '����́u�d�b�v�񂩂�d�b�̍��v���W�v
    Range("K4").Formula = "=COUNTIF(" & tel & ",""�o"")"
    Range("L4").Formula = "=COUNTIF(" & tel & ",""��"")"
    Range("M4").Formula = "=COUNTIF(" & tel & ",""�s��"")"
    Range("N4").Formula = "=COUNTIF(" & tel & ",""����"")"
    Range("O4").Formula = "=SUM(RC[-4]:RC[-1])"

    '���v�����ʏo���W�v�V�[�g�̊��ʏW�v���ʂ���W�v����֐������v�̉��̗�ɖ��ߍ���
    '���ʏo���W�v����n�K�L�ԐM�̍��v���W�v
    Range("B5").Formula = "=SUM(B" & ROW_KITOPDATA & ":B" & slrow & ")"
    Range("C5").Formula = "=SUM(C" & ROW_KITOPDATA & ":C" & slrow & ")"
    Range("D5").Formula = "=SUM(D" & ROW_KITOPDATA & ":D" & slrow & ")"
    Range("E5").Formula = "=SUM(E" & ROW_KITOPDATA & ":E" & slrow & ")"
    Range("F5").Formula = "=SUM(F" & ROW_KITOPDATA & ":F" & slrow & ")"

    '���ʏo���W�v���烁�[���EHP�̍��v���W�v
    Range("G5").Formula = "=SUM(G" & ROW_KITOPDATA & ":G" & slrow & ")"
    Range("H5").Formula = "=SUM(H" & ROW_KITOPDATA & ":H" & slrow & ")"
    Range("I5").Formula = "=SUM(I" & ROW_KITOPDATA & ":I" & slrow & ")"
    Range("J5").Formula = "=SUM(J" & ROW_KITOPDATA & ":J" & slrow & ")"
   
    '���ʏo���W�v����d�b�̍��v���W�v
    Range("K5").Formula = "=SUM(K" & ROW_KITOPDATA & ":K" & slrow & ")"
    Range("L5").Formula = "=SUM(L" & ROW_KITOPDATA & ":L" & slrow & ")"
    Range("M5").Formula = "=SUM(M" & ROW_KITOPDATA & ":M" & slrow & ")"
    Range("N5").Formula = "=SUM(N" & ROW_KITOPDATA & ":N" & slrow & ")"
    Range("O5").Formula = "=SUM(O" & ROW_KITOPDATA & ":O" & slrow & ")"

    '�o�ȘA���̍��v
    Range("P4").Formula = "=SUM(P" & ROW_KITOPDATA & ":P" & slrow & ")"

    '���ʏo���W�v���玖�O�����̍��v���W�v
    Range("R4").Formula = "=SUM(R" & ROW_KITOPDATA & ":R" & slrow & ")"
    Range("S4").Formula = "=SUM(S" & ROW_KITOPDATA & ":S" & slrow & ")"
    Range("T4").Formula = "=SUM(T" & ROW_KITOPDATA & ":T" & slrow & ")"

    '�o�ȗ\��̍��v
    Range("V4").Formula = "=SUM(V" & ROW_KITOPDATA & ":V" & slrow & ")"
    
    '�����o�Ȃ̍��v
    Range("X4").Formula = "=SUM(X" & ROW_KITOPDATA & ":X" & slrow & ")"
    Range("Y4").Formula = "=SUM(Y" & ROW_KITOPDATA & ":Y" & slrow & ")"
    Range("Z4").Formula = "=SUM(Z" & ROW_KITOPDATA & ":Z" & slrow & ")"
    
    '������̍��v
    Range("AB4").Formula = "=����!G1"
    Range("AB5").Formula = "=SUM(AB" & ROW_KITOPDATA & ":AB" & slrow & ")"
 
    '���ʕԐM�A���ʓd�b�A�o�\��A���ʓ����o���̏W�v������͈͂�\����������Z�b�g
    card = "����!R" & ROW_TOPDATA & "C" & COL_KICARD & ":R" & mlrow & "C" & COL_KICARD
    tel = "����!R" & ROW_TOPDATA & "C" & COL_KITEL & ":R" & mlrow & "C" & COL_KITEL
    plan = "����!R" & ROW_TOPDATA & "C" & COL_PLAN & ":R" & mlrow & "C" & COL_PLAN
    rslt = "����!R" & ROW_TOPDATA & "C" & COL_KIRSLT & ":R" & mlrow & "C" & COL_KIRSLT
 
    apay = "����!R" & ROW_TOPDATA & "C" & COL_ADVPAY & ":R" & mlrow & "C" & COL_ADVPAY
    kipay = "����!R" & ROW_TOPDATA & "C" & COL_KIPAY & ":R" & mlrow & "C" & COL_KIPAY
 
    mmbr1 = "����!R" & ROW_TOPDATA & "C" & COL_KI & ":R" & mlrow & "C" & COL_KI
    mmbr2 = "����!R" & ROW_TOPDATA & "C" & COL_NAME & ":R" & mlrow & "C" & COL_NAME
 
    '���ʂ̕ԐM�󋵂�d�b�󋵓����W�v����֐������ʂ̃Z���ɖ��ߍ���
    nrow = ROW_KITOPDATA
    Do
      ki = Cells(nrow, COL_KI).Text   '�Ώۂ́u���v�̕�������擾
     
      If ki = "" Then
        End
        Exit Do
      Else
   
        '����V�[�g�̊��ʕԐM����n�K�L�ԐM�󋵂̏W�v������֐���ԐM�n�K�L�̊e��ɖ��ߍ���
        Cells(nrow, COL_CARD_OK).FormulaR1C1 = "=COUNTIF(" & card & ",""" & ki & "�o�n"")"
        Cells(nrow, COL_CARD_NG).FormulaR1C1 = "=COUNTIF(" & card & ",""" & ki & "���n"")"
        Cells(nrow, COL_CARD_ERROR).FormulaR1C1 = "=COUNTIF(" & card & ",""" & ki & "�s��"")"
        Cells(nrow, COL_CARD_NOFIX).FormulaR1C1 = "=COUNTIF(" & card & ",""" & ki & "���n"")"
        Cells(nrow, COL_CARD_SUM).FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"

        '����V�[�g�̊��ʕԐM���烁�[���EWeb�ԐM�󋵂̏W�v������֐���ԐM���[���EHP�̊e��ɖ��ߍ���
        Cells(nrow, COL_EMAIL_OK).FormulaR1C1 = "=COUNTIF(" & card & ",""" & ki & "�o��"")"
        Cells(nrow, COL_EMAIL_NG).FormulaR1C1 = "=COUNTIF(" & card & ",""" & ki & "����"")"
        Cells(nrow, COL_EMAIL_NOFIX).FormulaR1C1 = "=COUNTIF(" & card & ",""" & ki & "����"")"
        Cells(nrow, COL_EMAIL_SUM).FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"

        '����V�[�g�̊��ʓd�b����d�b�̏󋵂��W�v������֐���d�b�̊e��ɖ��ߍ���
        Cells(nrow, COL_TEL_OK).FormulaR1C1 = "=COUNTIF(" & tel & ",""" & ki & "�o"")"
        Cells(nrow, COL_TEL_NG).FormulaR1C1 = "=COUNTIF(" & tel & ",""" & ki & "��"")"
        Cells(nrow, COL_TEL_ERROR).FormulaR1C1 = "=COUNTIF(" & tel & ",""" & ki & "�s��"")"
        Cells(nrow, COL_TEL_NOFIX).FormulaR1C1 = "=COUNTIF(" & tel & ",""" & ki & "����"")"
        Cells(nrow, COL_TEL_SUM).FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"

        Cells(nrow, COL_SUM_OK).FormulaR1C1 = "=COUNTIF(" & plan & ",""" & ki & "�o�\"")"

        '����V�[�g�̊��ʓ����o�ȂƊ��ʓ�������o�ȕԐM�Ǝ��O�����̏󋵂��W�v������֐��𖄂ߍ���
        Cells(nrow, COL_OKPAY).FormulaR1C1 = _
          "=SUMPRODUCT((" & plan & "=""" & ki & "�o�\"")*(" & apay & "<>""""))"
        Cells(nrow, COL_OTHERPAY).FormulaR1C1 = "=COUNTIF(" & plan & ",""" & ki & "����"")"
        Cells(nrow, COL_SUMPAY).FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"

        '�o�\��{�o�\��ȊO�̎��O�����ς݂��W�v
        Cells(nrow, COL_SUM_ALL).FormulaR1C1 = "=SUM(RC[-6],RC[-3])"
     
        '����V�[�g�̊��ʓ����o�Ȃ��瓖���o�Ȃ̏󋵂��W�v������֐��𓖓��o�ȗ�ɖ��ߍ���
        Cells(nrow, COL_RSLT_OK).FormulaR1C1 = "=COUNTIF(" & rslt & ",""" & ki & "�o"")"
        Cells(nrow, COL_RSLT_PAYNG).FormulaR1C1 = "=COUNTIF(" & rslt & ",""" & ki & "����"")"
        Cells(nrow, COL_RSLT_SUM).FormulaR1C1 = "=SUM(RC[-2]:RC[-1])"

        '����V�[�g�����������W�v������֐����������ɖ��ߍ���
        Cells(nrow, COL_MEMBER).FormulaR1C1 = _
           "=SUMPRODUCT((" & mmbr1 & "=""" & ki & """)*(" & mmbr2 & "<>""�|""))"
      End If
  
      nrow = nrow + 1
    Loop Until nrow = slrow + 1  '�ŏI�s+1�ɂȂ�܂ŌJ��Ԃ�
  
  End With
  ActiveSheet.EnableCalculation = True   '�Čv�Z

End Sub

'
'�u���ʓ����W�v�v�V�[�g�̊e�Z���ɏW�v�̂��߂̊֐��𖄂ߍ��ރ}�N��
Sub ���ʓ����W�v�֐��\�t()
  
  Dim nrow As Long      '�����Ώۂ̍s
  Dim lrow As Long      '�ŏI�s
  Dim kipay As String  '���v�v�Z�̂��߂̌v�Z���̕�����
  Dim ki As String   '�u���v��̕�����
  
  Application.ScreenUpdating = False      '�`��𒆎~
  ActiveSheet.EnableCalculation = False   '�Čv�Z�𒆎~
  
  With Sheets("����")
    lrow = .Cells(MEMBER_MAX, COL_ID).End(xlUp).Row
  End With
 
  Call SetKiNo("���ʓ����W�v")    '�u���v�̗��ݒ�
 
  '����̎��O�����̏󋵂��獇�v���W�v���邽�߂̌v�Z���̕������ݒ�
  kipay = "����!R" & ROW_TOPDATA & "C" & COL_KIPAY & ":R" & lrow & "C" & COL_KIPAY
 
  With Sheets("���ʓ����W�v")
    lrow = .Cells(MEMBER_MAX, COL_KI).End(xlUp).Row

    '���O�����̍��v���W�v����͈͂��Z�b�g
    For nrow = ROW_KITOPDATA To lrow
      ki = Cells(nrow, COL_KI).Text   '�u���v��̕������Z�b�g

      '�Y���́u���v�̓�����ޕʂ̏W�v�֐����Z�b�g
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

  ActiveSheet.EnableCalculation = True   '�Čv�Z���ĊJ
  Application.ScreenUpdating = True      '�`����ĊJ

End Sub


'
'�u�����W�v�v�V�[�g�̊e�Z���ɏW�v�̂��߂̊֐��𖄂ߍ��ރ}�N��
Sub �����W�v�֐��\�t()
  
  Dim nrow As Long      '�����Ώۂ̍s
  Dim lrow As Long      '����V�[�g�̍ŏI�s
  Dim trow As Long      '���Ԋ��̐擪�s
  Dim erow As Long      '���Ԋ��̍ŏI�s
  Dim kipay As String   '���v�v�Z�̂��߂̌v�Z���̕�����
  Dim ki As String      '�u���v��̕�����
  Dim obj As Range      '�������ʂ̃I�u�W�F�N�g
  
  
  Application.ScreenUpdating = False      '�`��𒆎~
  ActiveSheet.EnableCalculation = False   '�Čv�Z�𒆎~
  
  With Sheets("����")
    lrow = .Cells(MEMBER_MAX, COL_ID).End(xlUp).Row
  End With
 
  With Sheets("����").Range("A1:A" & MEMBER_MAX)
    ki = Sheets("�����W�v").Range("C1").Text   '�����W�v�V�[�g���瓖�Ԋ��̕�������擾
    Set obj = .Find(ki, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                 SearchOrder:=xlByColumns, MatchCase:=False, MatchByte:=False)
    trow = Range(obj.Address).Row        '���Ԋ��̐擪�s���擾
  
    ki = ki + 1                   '����������𓖔Ԋ��{1��
    Set obj = .Find(ki, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                 SearchOrder:=xlByColumns, MatchCase:=False, MatchByte:=False)
    erow = Range(obj.Address).Row - 1    '���Ԋ��̍ŏI�s���擾
  
  End With
 
  With Sheets("�����W�v")

    '����̎��O�����̏󋵂��瓖�Ԋ������W�v���邽�߂̌v�Z���̕������ݒ�
    kipay = "����!R" & trow & "C" & COL_ADVPAY & ":R" & erow & "C" & COL_ADVPAY
    Range("D4").Formula = "=COUNTIF(" & kipay & ",""MM"")"
    Range("D5").Formula = "=COUNTIF(" & kipay & ",""MF"")"
    Range("D6").Formula = "=COUNTIF(" & kipay & ",""MC"")"
    Range("D8").Formula = "=COUNTIF(" & kipay & ",""BM"")"
    Range("D9").Formula = "=COUNTIF(" & kipay & ",""BF"")"
    Range("D10").Formula = "=COUNTIF(" & kipay & ",""BC"")"
    Range("D12").Formula = "=COUNTIF(" & kipay & ",""CM"")"
    Range("D13").Formula = "=COUNTIF(" & kipay & ",""CF"")"
    Range("D14").Formula = "=COUNTIF(" & kipay & ",""CC"")"
    
    '����̓��������̏󋵂��瓖�Ԋ������W�v���邽�߂̌v�Z���̕������ݒ�
    kipay = "����!R" & trow & "C" & COL_PAY & ":R" & erow & "C" & COL_PAY
    Range("F4").Formula = "=COUNTIF(" & kipay & ",""YM"")"
    Range("F5").Formula = "=COUNTIF(" & kipay & ",""YF"")"
    Range("F8").Formula = "=COUNTIF(" & kipay & ",""GM"")"
    Range("F9").Formula = "=COUNTIF(" & kipay & ",""GF"")"
    Range("F12").Formula = "=COUNTIF(" & kipay & ",""KM"")"
    Range("F13").Formula = "=COUNTIF(" & kipay & ",""KF"")"
    
    '����̎��O�����̏󋵂���S�̂̏W�v���邽�߂̌v�Z���̕������ݒ�
    kipay = "����!R" & ROW_TOPDATA & "C" & COL_ADVPAY & ":R" & lrow & "C" & COL_ADVPAY
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
    
    '����̓��������̏󋵂���S�̂̏W�v���邽�߂̌v�Z���̕������ݒ�
    kipay = "����!R" & ROW_TOPDATA & "C" & COL_PAY & ":R" & lrow & "C" & COL_PAY
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

  ActiveSheet.EnableCalculation = True   '�Čv�Z���ĊJ
  Application.ScreenUpdating = True      '�`����ĊJ

End Sub

'���ʏo���W�v�V�[�g�A���ʓ����W�v�V�[�g�́u���v�̗��ݒ肷��B
'  StName  �F�ݒ肷��V�[�g��
Private Sub SetKiNo(StName As String)

  Dim nrow As Long      '�����Ώۂ̍s
  Dim lrow As Long      '�ŏI�s
  Dim kinum As Integer  '���̐�
  Dim ki(KI_MAX) As String   '���̕�������i�[����z��
  Dim i As Integer
    
  '�Ώۂ̊��ʃV�[�g�́u���v�̗���N���A
  With Sheets(StName)
    lrow = .Cells(KI_MAX, COL_KI).End(xlUp).Row
    Worksheets(StName).Range(Cells(ROW_KITOPDATA, COL_KI), Cells(lrow, COL_KI)).ClearContents
  End With
  
  '����V�[�g����u���v���擾
  With Sheets("����")
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

  '�Ώۂ̊��ʃV�[�g�Ɂu���v���Z�b�g
  For i = 0 To kinum
    Worksheets(StName).Cells(ROW_KITOPDATA + i, COL_KI) = "'" & ki(i)
  Next i

End Sub

'
'���e��I����ɖ{�N�x�̎��т����e����ɓ]�L����}�N��
Sub ���e����ѓ]�L()

  Dim nrow As Long    '�����s
  Dim lrow As Long    '����V�[�g�̍ŏI�s

  '�}�N�������s���邩�m�F
  If MsgBox("���e��I����ɖ{�N�x�̎��т�]�L���鏈���ł��B���s���܂����H", _
      vbYesNo + vbQuestion, "�������s") = vbYes Then

    With Sheets("����")
      lrow = .Cells(MEMBER_MAX, COL_ID).End(xlUp).Row
    End With

    For nrow = ROW_TOPDATA To lrow
   
      If Cells(nrow, COL_RSLT) = "�o" And Cells(nrow, COL_ADVPAY) <> "" Then    '�����o��(���O�x��)
        Cells(nrow, COL_RSLT0) = "��"
      ElseIf Cells(nrow, COL_RSLT) = "�o" And Cells(nrow, COL_PAY) <> "" Then   '�����o��(�����x��)
        Cells(nrow, COL_RSLT0) = "��"
      ElseIf Cells(nrow, COL_RSLT) = "�o" And Cells(nrow, COL_ADVPAY) = "" _
          And Cells(nrow, COL_PAY) = "" Then                                    '�����o��(��������)
        Cells(nrow, COL_RSLT0) = "��"
      ElseIf Cells(nrow, COL_RSLT) = "" And Cells(nrow, COL_ADVPAY) <> "" Then  '��������(���O�U��)
        Cells(nrow, COL_RSLT0) = "����"
      ElseIf Cells(nrow, COL_RSLT) = "" And _
          ((Cells(nrow, COL_CARD) = "�o" And Cells(nrow, COL_TEL) <> "��") _
            Or Cells(nrow, COL_TEL) = "�o") Then                                 '�o�ȘA���������������
        Cells(nrow, COL_RSLT0) = "�^"
      ElseIf Cells(nrow, COL_RSLT) = "" And _
          ((Cells(nrow, COL_CARD) = "��" And Cells(nrow, COL_TEL) <> "�o") _
            Or Cells(nrow, COL_TEL) = "��") Then                                 '���ȘA������
        Cells(nrow, COL_RSLT0) = "�~"
      ElseIf Cells(nrow, COL_RSLT) = "" And Cells(nrow, COL_CARD) = "�s��" Then  '�X�֕s��
        Cells(nrow, COL_RSLT0) = "�H"
      End If
    Next nrow
  End If
End Sub
'
'�u����v�V�[�g�́u��ƕ\�]�L�v�{�^���������ꂽ�Ƃ��̏���
    '== Mar.9,2014 Yuji Ogihara ==
    '�ۑ��`����2007 �ȍ~�̏����ɕύX�A�g���q��"xlsm"��
'�u�����V�[���E���D�E�o�Ȏ҈ꗗ�\�쐬.xlsm�v�ň����V�[���̍쐬�����s�����߂Ɍ��{�f�[�^��]�L
Sub ��ƕ\�]�L()

  Dim lrow As Long       '����V�[�g�̍ŏI�s
  Dim msg As String

  OrgBook = ActiveWorkbook.Name      '�J���Ă���t�@�C���̖��O�����{�Ƃ��Đݒ�

  '�����ɕK�v�ȃt�@�C�����J����Ă��邩�m�F
     '== Mar.9,2014 Yuji Ogihara ==
     '�ۑ��`����2007 �ȍ~�̏����ɕύX
  'If IsBookOpen("�����V�[���E���D�E�o�Ȏ҈ꗗ�\�쐬.xls") = False Then
  If IsBookOpen(SAGYOHYO_FILENAME) = False Then
    '== Mar.9,2014 Yuji Ogihara ==
    '�ۑ��`����2007 �ȍ~�̏����ɕύX
    'MsgBox "�u"�����V�[���E���D�E�o�Ȏ҈ꗗ�\�쐬.xls"�v���J����Ă��܂���I" _

    MsgBox "�u" & SAGYOHYO_FILENAME & "�v���J����Ă��܂���I" _
      & vbNewLine & "�J���Ă����蒼���ĉ������B"
    End
  End If
  
  msg = "��ƕ\�ɖ����]�L���܂����H"
  If MsgBox(msg, 4 + 32, "��ƕ\�]�L") = 6 Then

    '���ʂ̍��e��o���󋵂̏����B����ɃZ�b�g����}�N�������s
    Call �o���m�F

    '������V�[���E���D�E�o�Ȏ҈ꗗ�\�쐬.xls����J���A�\�t�������N���A
        '== Mar.9,2014 Yuji Ogihara ==
        '�ۑ��`����2007 �ȍ~�̏����ɕύX
    'Workbooks("�����V�[���E���D�E�o�Ȏ҈ꗗ�\�쐬.xls").Sheets("���{").Activate
    Workbooks(SAGYOHYO_FILENAME).Sheets("���{").Activate
    Rows("2:65536").Delete Shift:=xlUp
    Range("A1").Select
 
    Workbooks(OrgBook).Sheets("����").Activate
    lrow = Cells(MEMBER_MAX, COL_KI).End(xlUp).Row    '�ŏI�s�̍s�����擾
    
    '����̋L�ڕ����S�Ă��R�s�[
    Range(Cells(ROW_TOPTITLE, COL_KI), Cells(lrow, COL_COMMENT)).Select
    Selection.Copy
 
   '������V�[���E���D�E�o�Ȏ҈ꗗ�\�쐬.xls����J���A��ƕ\�ɓ\��t��
        '== Mar.9,2014 Yuji Ogihara ==
        '�ۑ��`����2007 �ȍ~�̏����ɕύX
   'Workbooks("�����V�[���E���D�E�o�Ȏ҈ꗗ�\�쐬.xls").Sheets("���{").Activate
    Workbooks(SAGYOHYO_FILENAME).Sheets("���{").Activate
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range("A4").Select
 
    '������J���A�R�s�[���[�h���L�����Z��
    Workbooks(OrgBook).Sheets("����").Activate
    Application.CutCopyMode = False
    Range("A5").Select
 
    '== Feb.9,2014 Yuji Ogihara ==
    '�ۑ��`����2007 �ȍ~�̏����ɕύX
    'Workbooks("�����V�[���E���D�E�o�Ȏ҈ꗗ�\�쐬.xls").Sheets("���{").Activate
    Workbooks(SAGYOHYO_FILENAME).Sheets("���{").Activate

 End If
 
End Sub

'�W�q�S���p�ɒS�����̖���t�@�C���̍쐬
'�@�u���ʏo���W�v�v�V�[�g�̏W�q�S����̒S���Җ����L�[�ɖ���t�@�C����S���Җ��ɕ�������
'�@�쐬�����t�@�C���́A���{�t�@�C���̂���t�H���_�[�Ɂu�z�z�p�v�t�H���_�[�����ۑ������
Sub �W�q�S���p����t�@�C���쐬()

  Dim tantoki(COL_TANTO_MAX, KI_MAX) As String     '�S���Җ��ƒS���������ۑ�����̂Q�����z��
  Dim kinum(KI_MAX) As Integer                   '�S���Җ��̒S��������̐�
  Dim serow(COL_TANTO_MAX, KI_MAX * 2) As Long   '�S���Җ��Ƃ��̒S���̊��̊J�n�s�A�I���s��ۑ�����̂Q�����z��
  Dim tmax As Integer      '�S���Ґ�
  Dim opath As String      '���{�t�@�C���̃p�X
     
  OrgBook = ActiveWorkbook.Name      '�J���Ă���t�@�C���̖��O�����{�Ƃ��Đݒ�
  If Mid(OrgBook, 9, 2) = "���{" Then
  
    Application.ScreenUpdating = False  '�`����~

    '���ʏo���W�v�V�[�g�́u�W�q�S���v��̖��O�̏��ƒS����������擾
    Call GetSyukyakuInfo(tantoki, kinum, tmax)

    '���{�t�@�C������A�W�q�S�����S��������̊J�n�s�ƏI���s���擾
    Call GetSyukyakuRow(opath, tantoki, kinum, tmax, serow)
   
    '�W�q�S���ւ̔z�z�p�Ɍ��{�t�@�C������S��������̏���؂�o���t�@�C���쐬
    Call MakeSyukyakuFile(opath, tantoki, kinum, tmax, serow)

    Application.ScreenUpdating = True   '�`����ĊJ
    
  Else
    MsgBox ("���̃t�@�C���͌��{�łȂ��̂ł��̏����͂ł��܂���I")
  End If

End Sub

'
'���ʏo���W�v�V�[�g�̐ݒ肳��ďW�q�S�����ƒS��������̏����擾����v���V�[�W��
' tantoki() �F�S���Җ��ƒS����������i�[���ꂽ�Q�����z��
' kinum()   �F�S���Җ��̒S��������̐����i�[���ꂽ�z��
' tmax      �F�S���Ґ�
Private Sub GetSyukyakuInfo(tantoki() As String, kinum() As Integer, tmax As Integer)

  Dim lrow As Integer      '�s��
  Dim ki As String      '��
  Dim tanto As String      '�S���Җ�
  Dim tno As Integer      '�S���Ҕԍ�
  Dim i As Integer

  '���ʏo���W�v�V�[�g�̍s�����J�E���g
  lrow = Worksheets("���ʏo���W�v").Range("A200").End(xlUp).Row

  '���ʏo���W�v�V�[�g�́u�W�q�S���v��̖��O���̃V�[�g���쐬
  tmax = 0
  For i = ROW_KITOPDATA To lrow
    ki = Worksheets("���ʏo���W�v").Cells(i, COL_KI).Value
    tanto = Worksheets("���ʏo���W�v").Cells(i, COL_TANTO).Value
  

    '�񎟌��z��tName��(X,0)�ɒS���Җ��A(X,1..)�ɒS���̊����Z�b�g
    If tmax = 0 Then    '�ŏ��̊��̏W�q�S���̏����Z�b�g
      tantoki(0, 0) = tanto
      kinum(0) = 1
      tantoki(0, kinum(0)) = ki
      tmax = 1
    Else             '�ȍ~�̊��̏W�q�S���̏����Z�b�g
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
'���{�̖���V�[�g����S�����̒S��������̊J�n�s�E�I���s���擾����v���V�[�W��
' opath     �F���{�t�@�C��������t�H���_�[��
' tantoki()   �F�S���Җ��ƒS����������i�[���ꂽ�Q�����z��
' kinum()   �F�S���Җ��̒S��������̐����i�[���ꂽ�z��
' tmax      �F�S���Ґ�
' serow()   �F�S���Җ��ƒS��������̊J�n�s�E�I���s���i�[���ꂽ�Q�����z��
Private Sub GetSyukyakuRow(opath As String, tantoki() As String, kinum() As Integer, tmax As Integer, serow() As Long)

  Dim tnNum As Integer     '�S���Ҕԍ�
  Dim nrow As Integer      '�������Ă���s�ԍ�
  Dim obj As Range         '�������ʂ̃I�u�W�F�N�g
  Dim i As Integer
  
  Workbooks(OrgBook).Activate      '���{�u�b�N���A�N�e�B�u��
  opath = ActiveWorkbook.Path      '���{�t�@�C���̃p�X���擾
  
  For tnNum = 0 To tmax - 1
    For i = 1 To kinum(tnNum)
    
      '�����KINO�񂩂��������
      With Sheets("����").Range("A1:A" & MEMBER_MAX)
        Set obj = .Find(tantoki(tnNum, i), LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                  SearchOrder:=xlByColumns, MatchCase:=False, MatchByte:=False)
      End With
     
      If Not obj Is Nothing Then
        nrow = Range(obj.Address).Row
        serow(tnNum, i * 2 - 1) = nrow     '�J�n�s���Z�b�g
     
        '��s���̊����قȂ�܂ŌJ��Ԃ�
        Do
          '�X�e�[�^�X�o�[�ɏ󋵕\��
          Application.StatusBar = "  ���W�q�S�� " & tantoki(tnNum, 0) & " ���S��������̊J�n�s�E�I���s���擾���D�D�D"
          
          If Cells(nrow, COL_KI) = Cells(nrow + 1, COL_KI) Then
            nrow = nrow + 1
          Else        '���ԍ����Ⴆ�΁A�I���s���Z�b�g�������𒆎~
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
  Application.StatusBar = False   '�X�e�[�^�X�o�[���N���A

End Sub

'
'���{�̖���V�[�g����S�����̒S��������̏��𕪊����ĕۑ�����v���V�[�W��
' opath     �F���{�t�@�C��������t�H���_�[��
' tantoki() �F�S���Җ��ƒS����������i�[���ꂽ�Q�����z��
' kinum()   �F�S���Җ��̒S��������̐����i�[���ꂽ�z��
' tMmx      �F�S���Ґ�
' serow()�F�S���Җ��ƒS��������̊J�n�s�E�I���s���i�[���ꂽ�Q�����z��
Private Sub MakeSyukyakuFile(opath As String, tantoki() As String, kinum() As Integer, tmax As Integer, serow() As Long)
  
  Dim lrow As Integer   '�s��
  Dim obj As Object     '�t�H���_�[�������I�u�W�F�N�g
  Dim npath As String   '�����t�@�C����ۑ�����t�H���_�[��
  Dim tnum As Integer   '�S���Ґ�
  Dim nbook As String   '������̃u�b�N��
  Dim trow As Long      '�s�v�����̐擪�s�ԍ�
  Dim brow As Long      '�s�v�����̍ŏI�s�ԍ�
  Dim dastr As String   '�t�@�C�����g�p���錎���̕�����
  Dim i As Integer
   
  lrow = Worksheets("����").Range("A" & MEMBER_MAX).End(xlUp).Row  '����̍s�����擾

  '���{�t�@�C���Ɠ����t�H���_�[���ɔz�z�p�t�H���_�[���Ȃ���΍쐬
  Set obj = CreateObject("Scripting.FileSystemObject")
  npath = opath & "\�z�z�p"
  If obj.FolderExists(folderspec:=npath) = False Then
    obj.createfolder npath
  End If

  '�S�����ɕ����������{�t�@�C���̃R�s�[���쐬
  For tnum = 0 To tmax - 1
     
    '�X�e�[�^�X�o�[�ɏ󋵕\��
    Application.StatusBar = "  ���W�q�S�� " & tantoki(tnum, 0) & " �p�̃t�@�C���𕪊��쐬���D�D�D"
    
    Workbooks.Add                    '�W�q�S���ʂɐV�K�Ƀu�b�N��ǉ�
    nbook = ActiveWorkbook.Name    '�ǉ������u�b�N�̖��O���擾
    Workbooks(OrgBook).Activate      '���{�u�b�N���A�N�e�B�u��
  
    '�V�K�u�b�N��Sheet1�̑O�Ɂu����v���R�s�[
    ThisWorkbook.Sheets("����").Copy After:=Workbooks(nbook).Sheets("Sheet1")
  
    '�V�K�u�b�N��Sheet1�`3���m�F���b�Z�[�W�Ȃ��ɍ폜
    Application.DisplayAlerts = False
    Workbooks(nbook).Worksheets("Sheet1").Delete
    
    '== Feb.9,2014 Yuji Ogihara ==
    '�V�K�u�b�N�쐬���̃V�[�g����Excel 2007�܂Łu3�v, 2010�ȍ~�́u1�v(�ݒ�ɂĕύX�\)
    '�ȉ���2�s�̓V�[�g���u3�v�̂Ƃ��̏����Ȃ̂ŁA2010�ȍ~�ł͕s�v
    'Workbooks(nbook).Worksheets("Sheet2").Delete
    'Workbooks(nbook).Worksheets("Sheet3").Delete
    Application.DisplayAlerts = True

    '== Feb.9,2014 Yuji Ogihara ==
    '�ȍ~�̍s�폜���� "Delete" �̍������̂���
    '�f�[�^�ŏI�s���烏�[�N�V�[�g�ŏI�s�܂ł��폜
    Range(lrow + 100 & ":" & Rows.Count).Delete

    '�S���ȊO�̊��̏����ŏI�s����폜
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
  
    '�V�K�u�b�N��S���҂̖��O��t�������t�@�C�����ŕۑ�
    '== Feb.9,2014 Yuji Ogihara ==
    '�ۑ��`����2007 �ȍ~�̏����ɕύX�A�g���q��"xlsx"��
    dastr = Format(Month(Date), "00") & Format(Day(Date), "00")
    Workbooks(nbook).SaveAs _
       Filename:=opath & "\�z�z�p\�������}���y�d�b�zto" & tantoki(tnum, 0) & dastr & ".xlsx", _
       Password:=MEIBO_PASSWORD
    ActiveWorkbook.Close
  Next tnum
  MsgBox "�����������I�����܂����B"
  Application.StatusBar = False   '�X�e�[�^�X�o�[���N���A
End Sub


'
'������p�����ɁA�@�N���𒼋�3�N���{�V�N�x���ɂ��A�A���e����т𒼋�5�N���{�V�N�x���ɂ��A
'�B���e��p�f�[�^���N���A����
Sub ���p������()

  Dim msg As String   '���b�Z�[�W
  Dim ystr As String  '�N�x������
  Dim afile As String '�J���Ă���t�@�C����
  Dim apath As String '�J���Ă���p�X��
  Dim bfile As String '�o�b�N�A�b�v�t�@�C����
  Dim nrow As Long    '�����s
  Dim lrow As Long    '����V�[�g�̍ŏI�s

  '�}�N�������s���邩�m�F
  ystr = Cells(ROW_TOPDATA - 1, COL_KAIHI0)
  msg = "20" & ystr & "�N�x���痂�N�x�ւ̖�����p�����̏��������s���܂����H"
  If MsgBox(msg, vbYesNo + vbQuestion, "�������s") = vbYes Then

    '�����O�̃t�@�C���̃o�b�N�A�b�v���쐬
'    afile = ActiveWorkbook.Name         '�J���Ă���t�@�C���̖��O���擾
'    apath = ActiveWorkbook.Path         '�J���Ă���t�@�C���̃p�X���擾
'    bfile = Left(afile, Len(afile) - 4) & "_OLD.xls"
'    Application.DisplayAlerts = False   '�ۑ��m�F�̃��b�Z�[�W���o���Ȃ�
'    ActiveWorkbook.SaveCopyAs bfile
'    Application.DisplayAlerts = True

    With Sheets("����")
      lrow = .Cells(MEMBER_MAX, COL_ID).End(xlUp).Row
    End With

    '�N���́u2�N�O�`���N�x�v���u3�N�O�`1�N�O�v�ɃR�s�[���A�u���N�x�v���N���A
    Range(Cells(ROW_TOPDATA - 1, COL_KAIHI2), Cells(lrow, COL_KAIHI0)).Copy
    Range(Cells(ROW_TOPDATA - 1, COL_KAIHI3), Cells(lrow, COL_KAIHI1)).PasteSpecial Paste:=xlValues
    Range(Cells(ROW_TOPDATA, COL_KAIHI0), Cells(lrow, COL_KAIHI0)).ClearContents
    Cells(ROW_TOPDATA - 1, COL_KAIHI0) = ystr + 1

    '���e����т́u4�N�O�`���N�x�v���u5�N�O�`1�N�O�v�ɃR�s�[���A�u���N�x�v���N���A
    Range(Cells(ROW_TOPDATA - 1, COL_RSLT4), Cells(lrow, COL_RSLT0)).Copy
    Range(Cells(ROW_TOPDATA - 1, COL_RSLT5), Cells(lrow, COL_RSLT1)).PasteSpecial Paste:=xlValues
    Range(Cells(ROW_TOPDATA, COL_RSLT0), Cells(lrow, COL_RSLT0)).ClearContents
    Cells(ROW_TOPDATA - 1, COL_RSLT0) = ystr + 1

    '���e��p�f�[�^���N���A
    Range(Cells(ROW_TOPDATA, COL_CARD), Cells(lrow, COL_COMMENT)).ClearContents
    Cells(ROW_TOPTITLE, COL_CARD) = "20" & (ystr + 1) & "���e��"
    
    Range("T5").Select
    MsgBox "���p���������I�����܂����B"
  End If
End Sub


