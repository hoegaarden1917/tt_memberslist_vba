VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm���� 
   Caption         =   "����"
   ClientHeight    =   1800
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5268
   OleObjectBlob   =   "UserForm����.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'�����{�^�����������Ƃ��̓���
Private Sub �����{�^��_Click()
 
  Dim ki As String      '���[�U�t�H�[���́u���v�ɓ��͂��ꂽ����
  Dim txt As String     '���[�U�t�H�[���́u�������v�ɓ��͂��ꂽ����
  Dim kil1 As String    '���[�U�t�H�[���́u���v�ɓ��͂��ꂽ�����̍�����1�����ڂ̕���
  Dim kir1 As String    '���[�U�t�H�[���́u���v�ɓ��͂��ꂽ�����̉E����1�����ڂ̕���
  Dim kir2 As String    '���[�U�t�H�[���́u���v�ɓ��͂��ꂽ�����̉E����2�����ڂ̕���
  Dim kilen As Integer  '���[�U�t�H�[���́u���v�ɓ��͂��ꂽ�����̕�����
  Dim trow As Long      '���������擪�s
  Dim nrow As Long      '�������Ă���s
  Dim obj As Range      '�������ʂ̃I�u�W�F�N�g
 
  '�����t�H�[���̓��͓��e���Z�b�g
  ki = UserForm����.TextBox��.Text
  txt = UserForm����.TextBox������.Text
  
  '�u���v�̓��͂������ꍇ�A�V�[�g�S�̂���u�������v������
  If ki = "" Then
    With Sheets("����").Range("A1:IV65536")
      Set obj = .Find(txt, After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, _
                  SearchDirection:=xlNext, MatchCase:=False, MatchByte:=False)
    End With
 
    If obj Is Nothing Then    '��������Ȃ���΁A���b�Z�[�W�\��
      MsgBox "�Y���̕�����͌����ł��܂���ł����I"
    Else                      '�������ꂽ�ꍇ�́A���̃Z����I��
      Range(obj.Address).Select
    End If
 
  '�u���v�̓��͂��L��ꍇ�A�V�[�g�S�̂���u�������v������
  Else
    kil1 = Left(ki, 1)
    kir1 = Right(ki, 1)
    kir2 = Right(ki, 2)
    kilen = Len(ki)
    
    '�u���v�ɑ��ƔN��a�����͂��ꂽ�ꍇ�A���ɕϊ�
    If kil1 = "S" Or kil1 = "s" Then        '���a�̏ꍇ�́A�N���Ɂ{23�Ŋ����v�Z
      If kilen = 2 Then
        ki = kir1 + 23
      ElseIf kilen = 3 Then
        ki = kir2 + 23
      End If
    ElseIf kil1 = "H" Or kil1 = "h" Then    '�����̏ꍇ�́A�N����+86�Ŋ����v�Z
      If kilen = 2 Then
        ki = kir1 + 86
      ElseIf kilen = 3 Then
        ki = kir2 + 86
      End If
    End If

    If Len(ki) = 2 Then
      ki = "0" & ki
    End If
   
    '�u���v����w�肵����
    With Sheets("����").Range("A1:A" & MEMBER_MAX)
      Set obj = .Find(ki, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                  SearchOrder:=xlByColumns, MatchCase:=False, MatchByte:=False)
    End With
 
    If obj Is Nothing Then       '�Y���̢�������������Ȃ��ꍇ
      MsgBox "�Y���̢���ԍ���͌����ł��܂���ł����I"
    
    Else                         '�Y���́u���v���������ꂽ�ꍇ
      If txt = "" Then           '�u�������v�����͂���Ă��Ȃ��Ƃ��A�u���v�̐擪�Z����I��
        Range(obj.Address).Select
      Else
        trow = Range(obj.Address).Row      '�u���v���������ꂽ�擪�s���擾
        nrow = trow                        '�擪�s����������s�̏����l�ɃZ�b�g
     
        Do
          If Cells(nrow, 1) = Cells(nrow + 1, 1) Then   '�P�s���̍s�́u���v�������ꍇ�A���̍s��
            nrow = nrow + 1
          Else                                          '�u���v���Ⴄ�ꍇ�A�J��Ԃ������𒆎~
            Exit Do
          End If
        Loop Until nrow = DOUKI_MAX + 1       'DOUKI_MAX + 1 �ɂȂ�܂ŌJ��Ԃ�
     
        '�͈͂��w�肵�A������������������B���̌������ʂ�ϐ��G��Ƃ���B
        With Sheets("����").Range(Range(Cells(trow, COL_NAME), Cells(nrow, COL_NAME)), Range(Cells(trow, COL_COMMENT), Cells(nrow, COL_COMMENT)))
     
          '2��ڂ̌����ŁA���������̊����ύX�ɂȂ��Ă���ȂǁA���݂̃Z���ʒu�������͈͊O�̏ꍇ��1��ڂɏC��
          If ActiveCell.Row < trow Or ActiveCell.Row > nrow Or ActiveCell.Column < COL_NAME Or ActiveCell.Column > COL_COMMENT Then
            UserForm����.Label5.Caption = 1
          End If
     
          'UserForm����.Label5�i�B�����x���j�̐����1��̏ꍇ�A��L�͈͂̍ŏ�s���猟��
          If UserForm����.Label5.Caption = 1 Then
            Set obj = .Find(txt, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, _
                       SearchDirection:=xlNext, MatchCase:=False, MatchByte:=False)
          
          '�����łȂ���΁A��L�͈͂̑I�����ꂽ�s�ȍ~�Ō���
          Else
            
            Set obj = .Find(txt, After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, _
                     SearchDirection:=xlNext, MatchCase:=False, MatchByte:=False)
          End If
        End With
          
        If Not obj Is Nothing Then      '�u�������v���������ꂽ�ꍇ�A���Y�Z����I��
          Range(obj.Address).Select
        Else                            '�u�������v����������Ȃ��ꍇ�A���b�Z�[�W�\��
          MsgBox "�Y���̢���E������̑g�ݍ��킹�ł͌����ł��܂���ł����I"
        End If
      End If
    End If
  End If
   
  '�����{�^�����N���b�N���閈�ɃJ�E���^�[���P���₷
  UserForm����.Label5.Caption = UserForm����.Label5.Caption + 1
 
End Sub


Private Sub CommandButton2_Click()    '����飃{�^�����N���b�N����B
  With UserForm����
    .TextBox��.Text = ""       '���[�U�[�t�H�[�������́u���v���̓��e���󔒂ɂ���B
    .TextBox������.Text = ""   '���[�U�[�t�H�[�������́u�������v���̓��e���󔒂ɂ���B
    .Label5.Caption = 1        '���[�U�[�t�H�[�������́u�J�E���^�[�v���̓��e���u1�v�ɂ���B
    .Hide                      '������ʂ����B
  End With

End Sub

'������Ɍ������ʂ��u�ύX�o�^�v�܂��́u�މ�o�^�v���鏈��
Private Sub CommandButton3_Click()
   
  Dim nrow As Long     '��������s
  Dim msg As String
  Dim lrow As Long     '���މ�҈ꗗ�t�@�C���̍ŏI�s
  Dim i As Integer
   
  Select Case CommandButton3.Caption  '�R�}���h�{�^��3�̖��̂ŏꍇ����
    Case "�ύX�o�^"  '��ύX�o�^��̏ꍇ
       
      Unload UserForm����
      With UserForm�o�^
        .Caption = "�ύX�o�^"       '���̂�ύX�o�^���
        .TextBox1.Enabled = False   '�e�L�X�g�{�b�N�X1����͕s��
        .Show
      End With

   
    Case "�މ�o�^"   '��މ�o�^��̏ꍇ
      nrow = ActiveCell.Row     '�I������Ă���s�����擾
      Range(Cells(nrow, COL_KI), Cells(nrow, COL_KI)).EntireRow.Select
    
      msg = "���̍s��މ�o�^�ɂ��܂����H"
      If MsgBox(msg, 4 + 32, "�މ�") = 6 Then    '��͂�����N���b�N�����ꍇ
        Range(Cells(nrow, COL_KI), Cells(nrow, COL_REMARK)).Copy   '�u���v�񂩂�u���l�v��܂ł��R�s�[
     
        '��������}���y���މ�҈ꗗ�z.xls����J���A���s�ɓ\��t��
        With Workbooks("�������}���y���މ�҈ꗗ�z.xls").Sheets("�މ��")
          lrow = .Range("A" & MEMBER_MAX).End(xlUp).Row + 1
          .Cells(lrow, 1) = Date
          .Cells(lrow, 2).PasteSpecial Paste:=xlValues, _
              Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End With
  
        '�u���O�v�񂩂�u���[���A�h���X�v��܂łɢ�|�����͂���������
        '�u�N�����v�Ɓu�R�����g���v���c���ăN���A�����ύX2011
        For i = COL_NAME To COL_EMAIL
          Cells(nrow, i) = "�|"
        Next i
        Range(Cells(nrow, COL_NAME), Cells(nrow, COL_EMAIL)).HorizontalAlignment = xlCenter
        Range(Cells(nrow, COL_BUKATSU), Cells(nrow, COL_COUPLE)).ClearContents
        Range(Cells(nrow, COL_KANJI), Cells(nrow, COL_KIPAY)).ClearContents
        
        Cells(nrow, COL_REMARK) = "�މ�̂��ߌ���"  '���l��ɢ�މ�̂��ߌ��ԣ�����
      End If
      Unload UserForm����
      
  End Select
 
End Sub



