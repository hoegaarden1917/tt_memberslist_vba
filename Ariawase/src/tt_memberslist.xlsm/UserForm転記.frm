VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm�]�L 
   Caption         =   "�]�L"
   ClientHeight    =   3555
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4284
   OleObjectBlob   =   "UserForm�]�L.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm�]�L"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
'�]�L�t�H�[���Łu�]�L�v�{�^���������ꂽ�ꍇ�̃��C������
'�@�u�ԐM�v�t�@�C���F�ԐM�񂪈قȂ�ꍇ�ɕԐM���]�L
'�@�u�d�b�v�t�@�C���F�d�b�񂪈قȂ�ꍇ�ɓd�b���]�L
'�@�u���O�v�t�@�C���F���O�񂪈قȂ�ꍇ�Ɏ��O���]�L
'�@�u�{���v�t�@�C���F�N���񂪈قȂ�ꍇ�ɔN�����]�L
'�@�u�o�ȁv�t�@�C���F�o�ȗ񂪈قȂ�ꍇ�ɏo�ȗ��]�L
'�@�u�����v�t�@�C���F�����񂪈قȂ�ꍇ�ɓ������]�L
'   ������̃t�@�C���̏ꍇ���A�R�����g���ɕύX���������ꍇ�͒ǋL�B
'   �܂��A��{���i�J�i�����A���ʁA�X�֔ԍ��A�Z��1�`4�A�d�b�ԍ��A���[���A�����A�o�g���w�A�v�w�j��
'   �ύX����Ă���ꍇ�͓]�L���f�̌�Ɋ�{���ɂ��Ă��]�L�B
'   �Ȃ��A���ɂ��Ă̓f�[�^��v�̊m�F�Ɏg�p���Ă��邽�߃}�N���ł͓]�L���Ȃ��悤�ɂ��Ă���B
Private Sub CommandButton1_Click()

  Dim job As String       '�]�L�����̓��e
  Dim pbook As String     '�]�L���t�@�C����
  Dim ski As String       '�]�L����J�n��
  Dim eki As String       '�]�L����I����
  Dim srow As Long        '�]�L����J�n�s
  Dim erow As Long        '�]�L����I���s
   
  '�}�N�����N�������t�@�C�������{�����m�F
  OrgBook = ActiveWorkbook.Name
  If Mid(OrgBook, 9, 2) <> "���{" Then
    MsgBox "���{�t�@�C�������ُ�ł��I�i���{�ւ̓]�L�ɂȂ��Ă��Ȃ��Ǝv���܂��j" _
      & vbNewLine & "���{�t�@�C���w" & OrgBook & "�x"
    End
  End If
 
  '�]�L�����̓��e�i�������e�A�]�L���t�@�C���A�J�n���A�I�����j���擾
  Call GetInputData(job, pbook, ski, eki)
  If job = "ERROR" Then
    End
  End If

  '�]�L���t�@�C�����J���Ă��邩�m�F
   If IsBookOpen(pbook) = False Then
     MsgBox "�]�L���t�@�C���w" & pbook & "�x���J����Ă��܂���I" & vbNewLine & "�J���Ă����蒼���ĉ������B"
     End
   End If
  
  '�]�L������������]�L���t�@�C���ɑ��݂��邩���m�F���A���̊J�n�s�E�I���s���擾
  Call ChkKiData(pbook, ski, eki, srow, erow)
  If srow = -1 Or erow = -1 Then
    End
  End If
   
  '�]�L��������J�n�s�E�I���s��\�����m�F
  Call ShowStartEndRow(pbook, srow, erow)
  If srow = -1 Or erow = -1 Then
    End
  End If

  '�]�L���t�@�C�����猴�{�ɓ]�L���邽�߂̏���
  Call InitPostData(pbook)

  '�]�L���t�@�C�����猴�{�ɓ]�L���{�i��{�����K�v�ɉ����ē]�L�j
  Call PostAllData(job, pbook, srow, erow)

End Sub


'�]�L�t�H�[���ɓ��͂��ꂽ����ݒ肷��v���V�[�W��
'   job     �F�t�H�[���őI������ď����̕�����i�G���[�̏ꍇ�͖߂�l�Ƃ���"ERROR"��ݒ�j
'   pbook �F�]�L���̃t�@�C����
'   ski,eki �F�]�L�J�n�̊��ƏI���̊�
Private Sub GetInputData(job As String, pbook As String, ski As String, eki As String)

  Dim jkey As String      '�t�@�C�����ɖ��ߍ��܂�Ă��鏈���L�[���[�h

  '�I�����ꂽ�����̃I�v�V�������擾
  If OptionButton1.Value = True Then
    job = "�ԐM"
  ElseIf OptionButton2.Value = True Then
    job = "�d�b"
  ElseIf OptionButton3.Value = True Then
    job = "���O"
  ElseIf OptionButton4.Value = True Then
    job = "�{��"
  ElseIf OptionButton5.Value = True Then
    job = "�o��"
  ElseIf OptionButton6.Value = True Then
    job = "����"
  End If
  
  pbook = TextBox1.Text    '�^�[�Q�b�g�i�]�L���j�t�@�C����
  jkey = Mid(pbook, 9, 2)
  
  If job = jkey Then   '�I�����ꂽ�����Ǝw�肳�ꂽ�t�@�C������v
  
    ski = TextBox11.Text      '�]�L���J�n����u���v
    eki = TextBox10.Text      '�]�L���I������u���v

  Else
    MsgBox "�I�����ꂽ��Ǝ�ނƃt�@�C��������v���܂���I"
    job = "ERROR"    '�G���[�Ƃ���
  End If

End Sub
'
'�]�L�ΏۂƂ��Đݒ肳�ꂽ�J�n�E�I���̊������݂��邩�m�F���A�J�n�s�E�I���s���擾����v���V�[�W��
'   pbook  �F�]�L���̃t�@�C����
'   ski,eki  �F�]�L�J�n�̊��ƏI���̊�
'   sRow,eRow�F�]�L�J�n�s�ƏI���s�i�G���[�̏ꍇ�͖߂�l�Ƃ���"-1"��ݒ�j
Private Sub ChkKiData(pbook As String, ski As String, eki As String, srow As Long, erow As Long)
   
  Dim msg As String
  Dim obj As Range
  Dim i As Integer
  
  msg = "�w�肳�ꂽ�t�@�C�� " & pbook & " ����]�L���܂����H"
  
  If MsgBox(msg, 4 + 32, "�]�L") = 6 Then  '�u�͂��v���N���b�N
     
    If ski <> "" And eki = "" Then    '�I���̊����������ݒ�̏ꍇ
      eki = ski
    End If
      
    '�]�L���t�@�C���̃V�[�g�u����v�́u���v��Łu�]�L�J�n�s�v������
    With Workbooks(pbook).Sheets("����").Range("A1:A" & MEMBER_MAX)
      
      If ski = "" Then     '�]�L�J�n�̊������ݒ�̏ꍇ
        srow = ROW_TOPDATA
      
      Else                 '�]�L�J�n�̊����ݒ肳��Ă���ꍇ
        Set obj = .Find(ski, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                 SearchOrder:=xlByColumns, MatchCase:=False, MatchByte:=False)
                 
        If Not obj Is Nothing Then       '�Y���́u���v�ԍ����������ꂽ
          srow = .Range(obj.Address).Row
        Else  '���ԍ�����������Ȃ���΃��b�Z�[�W��\�����A�G���[���Z�b�g
          MsgBox "�u�]�L�J�n���v�͌����ł��܂���ł����I"
          srow = -1
          Exit Sub
        End If
      End If
        
      '�]�L���t�@�C���̃V�[�g�u����v�́u���v��Łu�]�L�I���s�v������
      If eki = "" Then     '�]�L�I���̊������ݒ�̏ꍇ
        erow = .Range("A" & MEMBER_MAX).End(xlUp).Row
      
      Else                 '�]�L�I���̊����ݒ肳��Ă���ꍇ
        Set obj = .Find(eki, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                 SearchOrder:=xlByColumns, MatchCase:=False, MatchByte:=False)
                                     
        If Not obj Is Nothing Then
          erow = .Range(obj.Address).Row     '�u���v�ԍ����������ꂽ�s�����Z�b�g
          i = 0
          Do
            If .Cells(erow, COL_KI) = .Cells(erow + 1, COL_KI) Then
              erow = erow + 1
              i = i + 1
            Else    '���ԍ����Ⴆ�΁A�J��Ԃ������𒆎~�i���̎��_��LRow���ŏI�s�j
              Exit Do
            End If
          Loop Until i = DOUKI_MAX  'DOUKI_MAX�ɂȂ�܂ŌJ��Ԃ�
    
        Else  '���ԍ�����������Ȃ���΁A���b�Z�[�W��\�����A�G���[���Z�b�g
          MsgBox "�u�]�L�I�����v�͌����ł��܂���ł����I"
          erow = -1
          Exit Sub
        End If
      End If
    End With
  End If
End Sub
'
'�]�L�ΏۂƂ��Đݒ肳�ꂽ�J�n�E�I����\�����m�F����v���V�[�W��
'   pbook  �F�]�L���̃t�@�C����
'   srow,erow�F�]�L�J�n�s�ƏI���s�i�G���[�̏ꍇ�͖߂�l�Ƃ���"-1"��ݒ�j
Private Sub ShowStartEndRow(pbook As String, srow As Long, erow As Long)

  Dim msg1 As String
  Dim msg2 As String

  '�]�L�J�n�s��A���I�����A���b�Z�[�W��\�����Ċm�F
  Workbooks(pbook).Sheets("����").Activate
  Range(Cells(srow, COL_KI), Cells(srow, COL_KI)).EntireRow.Select

  msg1 = "���̍s����]�L���܂����H"
   
  '�]�L�I���s��A���I�����A���b�Z�[�W��\�����Ċm�F
  If MsgBox(msg1, 4 + 32, "�]�L�J�n�s�m�F") = 6 Then
    Range(Cells(erow, COL_KI), Cells(erow, COL_KI)).EntireRow.Select
    msg2 = "���̍s�܂œ]�L���܂����H"
       
    If MsgBox(msg2, 4 + 32, "�]�L�I���s�m�F") = 6 Then
      Cells(srow, COL_KI).Select    '�]�L�J�n�s��A���I������B
    Else   '�]�L�I���s�łȂ���΃G���[
      erow = -1
    End If
      
  Else  '�]�L�J�n�s�łȂ���΃G���[
    srow = -1
  End If
End Sub
'
'�]�L���t�@�C���̃f�[�^�����{�t�@�C���ɓ]�L���鏀���̂��߂̃v���V�[�W��
'  pBook   �F�]�L���t�@�C����
Private Sub InitPostData(pbook As String)
 
  Dim lrow As Long        '�ŏI�s

 '���{�t�@�C���Ɠ]�L���t�@�C���̃`�F�b�N����N���A
  Workbooks(OrgBook).Sheets("����").Activate   '���{�t�@�C�����A�N�e�B�u
  With Workbooks(OrgBook).Sheets("����")
    lrow = .Range("A" & MEMBER_MAX).End(xlUp).Row
    Worksheets("����").Range(Cells(ROW_TOPDATA, COL_CHECK), Cells(lrow, COL_CHECK)).ClearContents
  End With
  
  Workbooks(pbook).Sheets("����").Activate   '�]�L�t�@�C�����A�N�e�B�u
  With Workbooks(pbook).Sheets("����")
    lrow = .Range("A" & MEMBER_MAX).End(xlUp).Row
    Worksheets("����").Range(Cells(ROW_TOPDATA, COL_CHECK), Cells(lrow, COL_CHECK)).ClearContents
  End With
  
  '��{����]�L���邩�m�F
  If MsgBox("���{�Ɠ]�L���̊�{���i�Z�����j���قȂ�ꍇ�A" & vbNewLine & "��{���i�Z�����j���m�F���Ȃ���]�L���܂����H", vbYesNo + vbQuestion) = vbYes Then
    PostFlag = 1
  Else
    PostFlag = 2
  End If
 
 End Sub
  
'
'�]�L���t�@�C���̃f�[�^�����{�t�@�C���ɓ]�L����v���V�[�W��
'  job�@     �F�������e
'  pbook     �F�]�L���t�@�C����
'  srow, erow�F�]�L�J�n�s�A�I���s
Private Sub PostAllData(job As String, pbook As String, srow As Long, erow As Long)
 
  Dim pdata(4) As String   '�]�L���t�@�C���̓]�L�f�[�^������z��
  Dim prow As Long         '�]�L�������̍s
  Dim obj As Range         '��������
  Dim orow As Long         '�������ʂ̌��{�s
  Dim per As Integer       '�i���p�[�Z���e�[�W
  Dim ecode As Integer     '�G���[�R�[�h
   
  Application.Calculation = xlManual  '�Čv�Z���~
  Application.ScreenUpdating = False  '�ĕ`����~
  
  prow = srow
  Do
    Workbooks(OrgBook).Sheets("����").Activate
    
    '�]�L�������e�̃f�[�^�L���𒲂ׁA��{���f�[�^���擾
    Call GetPostData(job, pbook, prow, pdata)
    
    '�X�e�[�^�X�o�[�ɐi�s�󋵕\��
    per = (prow - srow + 1) / (erow - srow + 1) * 100
    Application.StatusBar = "��" & pbook & "����̓]�L�@" & prow - srow + 1 & "�s�ڂ� ID�F" _
       & pdata(1) & "�������� (" & per & " %)�D�D�D"

    '�]�L���f�[�^��ID�ԍ����L�[�ɁA���{�t�@�C���̓]�L�Ώۂ�����
    Workbooks(OrgBook).Sheets("����").Activate
    With Sheets("����").Range("C1:C" & MEMBER_MAX)
      Set obj = .Find(pdata(1), LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                 SearchOrder:=xlByColumns, MatchCase:=False, MatchByte:=False)
    End With
 
    If Not obj Is Nothing Then         '�Y���̢ID����������ꂽ�ꍇ
      orow = Range(obj.Address).Row
     
      '����A���������v���Ă���ꍇ
      If Cells(orow, COL_KI) = pdata(0) And Cells(orow, COL_NAME) = pdata(2) Then
        
        '�]�L�Ώۃf�[�^�����{�ɓ]�L
        If pdata(3) <> "NO_DATA" Then
          Call PostData(job, orow, pdata)
        End If
        
        '�R�����g���ύX����Ă���ꍇ�A���{�ɒǋL
        If pdata(4) <> "" Then
          Call AddCmnt(orow, pdata(4))
        End If
    
        '��{��񂪕ύX����Ă���ꍇ�A���f��ɓ]�L
        If PostFlag = 1 Then
          Call PostBasicData(orow, pbook, prow)
        End If
        
        '��{��񂪕ύX����Ă���ꍇ�A�]�L�Ώۂ�h��Ԃ�
        If PostFlag = 2 Then   '�]�L�������Ɂu�Ȍ�L�����Z���v�������ꂽ�ꍇ��z�肵�A�ʂ�If����
          Call PaintBasicData(orow, pbook, prow)
        End If
 
      Else       '�u���A�����v����v���Ȃ��ꍇ�A�����p�������m�F
        ecode = 1
        Call IsPostCont(ecode, pdata(1), pbook, prow)
        If ecode <> 0 Then   '���~
          Exit Do
        Else
          Workbooks(pbook).Sheets("����").Cells(prow, COL_CHECK) = "�ُ�"
        End If
      End If
    
    Else      '�ID��������ł��Ȃ��ꍇ�A�����p�������m�F
      ecode = 2
      Call IsPostCont(ecode, pdata(1), pbook, prow)
      If ecode <> 0 Then   '���~
        Exit Do
      Else
        Workbooks(pbook).Sheets("����").Cells(prow, COL_CHECK) = "�ُ�"
      End If
    End If
   
    prow = prow + 1
  Loop Until prow = erow + 1
                
  MsgBox "�]�L�������������܂����B"
  
  Application.StatusBar = False         '�X�e�[�^�X�o�[�̏���
  Application.Calculation = xlAutomatic '�Čv�Z���ĊJ
  Application.ScreenUpdating = True     '�ĕ`����ĊJ
  Workbooks(OrgBook).Sheets("����").Activate   '���{�t�@�C�����A�N�e�B�u
  
  Unload UserForm�]�L

End Sub
'
'�]�L�Ώۂ̃f�[�^��ǂݎ��ƂƂ��Ɋ�{�����擾����v���V�[�W��
'  job�@   �F�������e
'  pbook   �F�]�L���t�@�C����
'  prow    �F�]�L�Ώۍs
'  pdata() �F�]�L�f�[�^
Private Sub GetPostData(job As String, pbook As String, prow As Long, pdata() As String)
   
  Dim pcol As Long
   
  '�]�L���̓]�L�Ώۗ�Ƀf�[�^������ꍇ�A�����ɍ��킹�ē]�L�f�[�^���擾
  With Workbooks(pbook).Sheets("����")
    pdata(0) = .Cells(prow, COL_KI)
    pdata(1) = .Cells(prow, COL_ID)
    pdata(2) = .Cells(prow, COL_NAME)
    
    If job = "�ԐM" And .Cells(prow, COL_CARD) <> "" Then      '�ԐM�f�[�^�̎擾
      pdata(3) = .Cells(prow, COL_CARD)
      .Cells(prow, COL_CHECK) = "��"
    ElseIf job = "�d�b" And .Cells(prow, COL_TEL) <> "" Then     '�d�b�f�[�^�̎擾
      pdata(3) = .Cells(prow, COL_TEL)
      .Cells(prow, COL_CHECK) = "��"
    ElseIf job = "���O" And .Cells(prow, COL_ADVPAY) <> "" Then   '���O�f�[�^�̎擾
      pdata(3) = .Cells(prow, COL_ADVPAY)
      .Cells(prow, COL_CHECK) = "��"
    ElseIf job = "�{��" And .Cells(prow, COL_KAIHI0) <> "" Then     '�{���f�[�^�̎擾
      pdata(3) = .Cells(prow, COL_KAIHI0)
      .Cells(prow, COL_CHECK) = "��"
    ElseIf job = "�o��" And .Cells(prow, COL_RSLT) <> "" Then  '�o�ȃf�[�^�̎擾
      pdata(3) = .Cells(prow, COL_RSLT)
      .Cells(prow, COL_CHECK) = "��"
    ElseIf job = "����" And .Cells(prow, COL_PAY) <> "" Then     '�����f�[�^�̎擾
      pdata(3) = .Cells(prow, COL_PAY)
      .Cells(prow, COL_CHECK) = "��"
    Else
      pdata(3) = "NO_DATA"
    End If

    If .Cells(prow, COL_COMMENT) <> "" Then         '�R�����g���擾
      pdata(4) = .Cells(prow, COL_COMMENT)
      .Cells(prow, COL_CHECK) = "��"
    Else
      pdata(4) = ""
    End If

    '��{�����擾���A���W���[���S�̂Ŏg�p����ϐ�NewBasicData�ɃZ�b�g
    NewBasicData = ""
    For pcol = COL_KI To COL_JHSCHOOL
        NewBasicData = NewBasicData & .Cells(prow, pcol) & ","
    Next
    NewBasicData = NewBasicData & .Cells(prow, COL_COUPLE)
  
  End With
End Sub

'
'�]�L�Ώۂ̃f�[�^�����{���猟�����A���{�ɓ]�L����v���V�[�W��
'  job�@   �F�������e
'  orow    �F���{�t�@�C���̍s��
'  pdata() �F�]�L�Ώۃf�[�^�i0:���A1:ID�A2:�����A3:�]�L�f�[�^�A4:�]�L�R�����g�j
Private Sub PostData(job As String, orow As Long, pdata() As String)

  With Workbooks(OrgBook).Sheets("����")
    
    Select Case job
      Case "�ԐM"    '��ԐM��f�[�^����̓]�L
        If .Cells(orow, COL_CARD) <> pdata(3) Then
          .Cells(orow, COL_CARD) = pdata(3)
          .Cells(orow, COL_CHECK) = "��"
         End If
             
       Case "�d�b"    '��d�b��f�[�^����̓]�L
         If .Cells(orow, COL_TEL) <> pdata(3) Then
           .Cells(orow, COL_TEL) = pdata(3)
           .Cells(orow, COL_CHECK) = "��"
         End If
                                  
       Case "���O"    '����O��f�[�^����̓]�L
         If .Cells(orow, COL_ADVPAY) <> pdata(3) Then
           .Cells(orow, COL_ADVPAY) = pdata(3)
           .Cells(orow, COL_CHECK) = "��"
         End If
        
       Case "�{��"    '��{����f�[�^����̓]�L
         If .Cells(orow, COL_KAIHI0) <> pdata(3) Then
           .Cells(orow, COL_KAIHI0) = pdata(3)
           .Cells(orow, COL_CHECK) = "��"
         End If
            
       Case "�o��"    '��o�ȣ�f�[�^����̓]�L
         If .Cells(orow, COL_RSLT) <> pdata(3) Then
           .Cells(orow, COL_RSLT) = pdata(3)
           .Cells(orow, COL_CHECK) = "��"
         End If
            
       Case "����"    '�������f�[�^����̓]�L
         If .Cells(orow, COL_PAY) <> pdata(3) Then
           .Cells(orow, COL_PAY) = pdata(3)
           .Cells(orow, COL_CHECK) = "��"
         End If
    End Select
  End With
End Sub

'
'�]�L���f�[�^�ɃR�����g������ꍇ�ɁA���s���ǋL����v���V�[�W��
'  orow    �F�]�L�����s
'  cmnt    �F�]�L����R�����g
Private Sub AddCmnt(orow As Long, cmnt As String)

  With Workbooks(OrgBook).Sheets("����")
    If StrComp(cmnt, Cells(orow, COL_COMMENT)) = 0 Then       '�����R�����g�Ɠ����ꍇ
      '�������Ȃ�
    ElseIf Cells(orow, COL_COMMENT) = "" Then                 '�����R�����g�������ꍇ
      .Cells(orow, COL_COMMENT) = cmnt
      .Cells(orow, COL_CHECK) = "��"
    ElseIf InStr(cmnt, Cells(orow, COL_COMMENT)) > 0 Then     '�����R�����g���܂܂�Ă���ꍇ
      .Cells(orow, COL_COMMENT) = cmnt
      .Cells(orow, COL_CHECK) = "��"
    Else                                                      '�����R�����g���܂܂�Ă��Ȃ��ꍇ
      .Cells(orow, COL_COMMENT) = Cells(orow, COL_COMMENT) & vbLf & cmnt
      .Cells(orow, COL_CHECK) = "��"
    End If
  
    .Cells(orow, COL_COMMENT).Font.Size = 8
  End With

End Sub
'
'�]�L�Ώۂ́u���E�����v����v���Ȃ��A�uID�v�������ł��Ȃ��ꍇ��
'�G���[��\�����������p�����邩�m�F����v���V�[�W��
'  ecode  �F�G���[�R�[�h
'  id     �F�G���[�ɂȂ���ID
'  pbook  �F�]�L���t�@�C����
'  prow   �F�G���[�ɂȂ����s
Private Sub IsPostCont(ecode As Integer, id As String, pbook As String, prow As Long)

  Dim msg As String

  Application.ScreenUpdating = True     '�ĕ`����ĊJ
  Workbooks(pbook).Sheets("����").Activate
  Range(Cells(prow, COL_ID), Cells(prow, COL_ID)).EntireRow.Select
  Application.ScreenUpdating = False    '�ĕ`����~

  If ecode = 1 Then
    msg = "����E���������v���܂���I <ID: " & id & " >" & vbNewLine & _
               "�Ȍ�̃f�[�^�̏������p�����܂����H"
  ElseIf ecode = 2 Then
    msg = "<ID: " & id & " >�����{�Ō����ł��܂���B" & vbNewLine & _
                "�Ȍ�̃f�[�^�̏������p�����܂����H"
  End If
                
  '�G���[���b�Z�[�W��\�����A�����p�����m�F
  If MsgBox(msg, 4 + 32, "�����p��") = 6 Then    '�p��
    ecode = 0
  End If
End Sub

'
'�]�L���f�[�^�ƌ��{�f�[�^�̊�{��񂪈قȂ�ꍇ�ɁA��{�����C������v���V�[�W��
'  orow  �F���{�̍s��
'  pbook    �F�]�L���t�@�C����
'  prow     �F�]�L���t�@�C���̏����s
Private Sub PostBasicData(orow As Long, pbook As String, prow As Long)

  Dim ocol As Long

  '���{�̊�{�������W���[���S�̂Ŏg�p����ϐ�OrgBasicData�ɃZ�b�g
  With Workbooks(OrgBook).Sheets("����")

    OrgBasicData = ""
    For ocol = COL_KI To COL_JHSCHOOL
        OrgBasicData = OrgBasicData & .Cells(orow, ocol) & ","
    Next
    OrgBasicData = OrgBasicData & .Cells(orow, COL_COUPLE)
  
  End With

  '��{��񂪈قȂ�ꍇ�Ɋ�{���]�L�t�H�[���̌Ăяo��
  If OrgBasicData <> NewBasicData Then
     Workbooks(pbook).Sheets("����").Cells(prow, COL_CHECK) = "��"
     OrgRow = orow     '��{�f�[�^�C���p�Ɍ��{�̌��ݍs�����W���[���S�̂Ŏg�p����ϐ�OrgRow�ɃZ�b�g
     UserForm��{�f�[�^�]�L.Show
  End If
  
  Unload UserForm��{�f�[�^�]�L

End Sub


'�]�L���f�[�^�ƌ��{�f�[�^�̊�{��񂪈قȂ�ꍇ�ɁA�]�L���̊�{�����x�^�h�肷��v���V�[�W��
'  orow  �F���{�̍s��
'  pbook    �F�]�L���t�@�C����
'  prow     �F�]�L���t�@�C���̏����s
Private Sub PaintBasicData(orow As Long, pbook As String, prow As Long)
  
  Dim odata As Variant  '���{�̊�{�f�[�^
  Dim pdata As Variant  '�]�L���̊�{�f�[�^
  Dim ocol As Long
  Dim i As Integer

  With Workbooks(OrgBook).Sheets("����")

    OrgBasicData = ""
    For ocol = COL_KI To COL_JHSCHOOL
        OrgBasicData = OrgBasicData & .Cells(orow, ocol) & ","
    Next
    OrgBasicData = OrgBasicData & .Cells(orow, COL_COUPLE)
  
  End With

  '���{�f�[�^�ƏC���f�[�^��z��ϐ��ɃZ�b�g����
  '�f�[�^��[0],[1],����ƃZ�b�g�����̂ŁA�f�[�^�̈����͒���
  odata = Split(OrgBasicData, ",")
  pdata = Split(NewBasicData, ",")
  
  '���{�f�[�^���މ������Ă���i�������u�|�v�j�ꍇ�͓]�L���Ȃ�
  If odata(COL_NAME - 1) <> "�|" Then
    With Workbooks(pbook).Sheets("����")
      
      '�f�[�^���r�i�u�J�i�����v�`�u�v�w�v�j
      For i = COL_KANA To COL_COUPLE
      
        '���{�f�[�^�ƏC���f�[�^���قȂ�ꍇ�A�]�L���̃Z�������[�Y�Ɂi�]�L�͂��Ȃ��j
        If odata(i - 1) <> pdata(i - 1) Then
          Workbooks(pbook).Sheets("����").Activate
          .Range(Cells(prow, i), Cells(prow, i)).Interior.ColorIndex = 38
          Workbooks(OrgBook).Sheets("����").Activate
        End If
      Next i
    End With
  End If

End Sub


'
'�u����v�{�^���������ꂽ�Ƃ��̏���
Private Sub CommandButton2_Click()
        
  Unload UserForm�]�L
    
End Sub




