VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm�o�^ 
   Caption         =   "�o�^"
   ClientHeight    =   3660
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10020
   OleObjectBlob   =   "UserForm�o�^.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm�o�^"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' �o�^��ʂ̢���ˏZ����{�^���������ꂽ�ꍇ�̓�����L�q
Private Sub CommandButton1_Click()
   
  Dim zip As String        '�X�֔ԍ�
  Dim abook As String      '�X�֔ԍ��f�[�^�t�@�C����
  Dim obj As Range         '��������
  Dim nrow As Long         '���ݍs
  Dim addrs(3) As String   '�Z����񕶎��z��
   
  '�X�֔ԍ��̕�������Z�b�g
  zip = UserForm�o�^.TextBox4.Text & UserForm�o�^.TextBox5.Text
  
  '== Mar.15,2014 Yuji Ogihara ==
  '�t�@�C��"�X�֔ԍ��ް��y�S���Łz"�ɂ��Ă̕ύX
  '  1.  Excel 2007 �`��(xlsx)�ւ̕ۑ��`���ύX
  '  2. ���f�[�^����̃t�@�C���쐬�菇�ȑf���ɔ����f�[�^�t�B�[���h�̕ύX
       
  
  '== Mar.15,2014 Yuji Ogihara ==
  ' xls -> xlsx
  'abook = "�X�֔ԍ��ް��y�S���Łz.xls"
  abook = "�X�֔ԍ��ް��y�S���Łz.xlsx"
  
       
  '�X�֔ԍ��f�[�^�t�@�C�����J���Ă��邩�m�F
   If IsBookOpen(abook) = False Then
     MsgBox "�X�֔ԍ��f�[�^�t�@�C���w" & abook & "�x���J����Ă��܂���I" _
               & vbNewLine & "�J���Ă����蒼���ĉ������B"
     End
   End If
  
  '�X�֔ԍ�������
  '== Mar.15,2014 Yuji Ogihara ==
  '�X�֔ԍ��̗��"C"�ɁA�͈͂��uC��S�́v��
  'With Workbooks(abook).Sheets("�X�֔ԍ�1").Range("B1:B65001")
  With Workbooks(abook).Sheets("�X�֔ԍ�1").Range("C:C")
    Set obj = .Find(zip, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                   MatchCase:=False, MatchByte:=False)
  End With
   
  If Not obj Is Nothing Then     '�V�[�g��X�֔ԍ�1��ŗX�֔ԍ������������ꍇ
    With Workbooks(abook).Sheets("�X�֔ԍ�1")
      nrow = .Range(obj.Address).Row
      
     '== Mar.15,2014 Yuji Ogihara ==
     ' �Z���̗�ԍ���7-9 �ɕύX
     'addrs(0) = .Cells(nrow, 3)
     'addrs(1) = .Cells(nrow, 4)
     'addrs(2) = .Cells(nrow, 5)
      addrs(0) = .Cells(nrow, 7)
      addrs(1) = .Cells(nrow, 8)
      addrs(2) = .Cells(nrow, 9)
    End With
      
  Else                           '�V�[�g��X�֔ԍ�1��ŗX�֔ԍ�����������Ȃ��ꍇ
     '== Mar.15,2014 Yuji Ogihara ==
     ' �V�[�g�u�X�֔ԍ�2�v�p�~
    'With Workbooks(abook).Sheets("�X�֔ԍ�2").Range("B1:B65001")
    '  Set obj = .Find(zip, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
    '                  MatchCase:=False, MatchByte:=False)
    'End With
      
    'If obj Is Nothing Then     '�V�[�g��N�G��2��ŗX�֔ԍ����������Ȃ��ꍇ
      MsgBox "�Y���̗X�֔ԍ��͌��o����܂���ł����I"
      End
             
    'Else                         '�V�[�g��N�G��2��ŗX�֔ԍ����������ꂽ�ꍇ
    '  With Workbooks("�X�֔ԍ��ް��y�S���Łz.xls").Sheets("�X�֔ԍ�2")
    '    nrow = .Range(obj.Address).Row
    '    addrs(0) = .Cells(nrow, 3)
    '    addrs(1) = .Cells(nrow, 4)
    '    addrs(2) = .Cells(nrow, 5)
    '  End With
    'End If
  End If
     
  '�e�L�X�g�{�b�N�X6�`8�Ɍ������ʂ����
  UserForm�o�^.TextBox6.Text = addrs(0)
  UserForm�o�^.TextBox7.Text = addrs(1)
  UserForm�o�^.TextBox8.Text = addrs(2)
   
End Sub

' �o�^��ʂ̢�o�^��{�^���������ꂽ�ꍇ�̓�����L�q
Private Sub CommandButton2_Click()

  Dim obj As Range           '�������ʂ̃I�u�W�F�N�g
  Dim nrow As Long           '�������Ă���s
  Dim msg As String
 
  Select Case UserForm�o�^.Caption
    Case "����o�^"    '���[�U�[�t�H�[���o�^�̖��̂������o�^��̏ꍇ�

      Call GetKiTopRow(nrow)   '�o�^������̐擪�s���擾
      
      Do
        If Cells(nrow, COL_KI) = Cells(nrow + 1, COL_KI) Then   '�P�s���̍s�̊�������
          nrow = nrow + 1
        
        Else    '���ԍ����Ⴆ�΁A���̊��̍ŏI�s��I�����A�J��Ԃ������𒆎~����B
          Range(Cells(nrow, COL_KI), Cells(nrow, COL_KI)).EntireRow.Select
          msg = "���̍s�̉��ɂP�s�ǉ����āA����o�^���܂����H"
          
          If MsgBox(msg, 4 + 32, "����") = 6 Then   '��͂�����N���b�N���ꂽ�ꍇ
            Application.Calculation = xlManual      '�Čv�Z���~
            nrow = nrow + 1
            Call InsertNewRow(nrow)      '���Y���̍ŏI�s��1�s�ǉ�
            Call RegNewData(nrow)        '�ǉ������s�ɐV�K�f�[�^���Z�b�g
            Call RegNewDataToList(nrow)  '�u���މ�ꗗ�v�ɐV�K�f�[�^��o�^
       
            Application.Calculation = xlAutomatic  '�Čv�Z���ĊJ
            Cells(nrow, COL_NAME).Select
          End If
          Exit Do
        End If
      
      Loop Until nrow = DOUKI_MAX  'DOUKI_MAX�ɂȂ�܂ŌJ��Ԃ�
      Unload UserForm�o�^
   
    Case "�ύX�o�^"    '���[�U�[�t�H�[���o�^�̖��̂���ύX�o�^��̏ꍇ
 
      msg = "���̍s�ɕύX����o�^���܂����H"
      If MsgBox(msg, 4 + 32, "�ύX") = 6 Then  '��͂���̏ꍇ
        Application.Calculation = xlManual     '�Čv�Z���~
        nrow = ActiveCell.Row
   
        Call RegEditData(nrow)       '�C���f�[�^��o�^
     
        Application.Calculation = xlAutomatic  '�Čv�Z���ĊJ
   
      End If
      Unload UserForm�o�^
      
  End Select

End Sub

'�o�^�t�H�[���ɓ��͂��ꂽ�u���v�̐擪�s���擾����v���V�[�W��
Sub GetKiTopRow(trow As Long)

  Dim ki As String       '���[�U�t�H�[���́u���v�ɓ��͂��ꂽ����
  Dim obj As Range       '�������ʂ̃Z��
  Dim i As Integer
  
  ki = UserForm�o�^.TextBox1.Text
    
  '�͈͂��w�肵����������
  With Sheets("����").Range("A1:A" & MEMBER_MAX)
    Set obj = .Find(ki, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                   SearchOrder:=xlByColumns, MatchCase:=True, MatchByte:=False)
  End With
 
  If Not obj Is Nothing Then  '�u���v���������ꂽ�ꍇ
    trow = Range(obj.Address).Row
    
  Else       '�u���v����������Ȃ��ꍇ�A�u���v��菬�����ő�́u���v������
    
    i = 1
    Do
      
      '�͈͂��w�肵�A�u���v������
      With Sheets("����").Range("A1:A" & MEMBER_MAX)
        Set obj = .Find(ki - i, LookIn:=xlValues, LookAt:=xlPart, SearchDirection:=xlNext, _
                       SearchOrder:=xlByColumns, MatchCase:=True, MatchByte:=False)
      End With
        
      If Not obj Is Nothing Then    '�u���v���������ꂽ�ꍇ
        trow = Range(obj.Address).Row
        Exit Do
      Else                          '�u���v����������Ȃ��ꍇ
        i = i + 1    '1�������u���v������
      End If
    Loop Until i = 30    '30���������܂Ō���
  End If
End Sub


'���Y�̊��̍ŏI�s�ɐV�K�f�[�^��o�^���邽�߂�1�s�ǉ�����v���V�[�W��
Sub InsertNewRow(nrow As Long)

  Selection.EntireRow.Insert    '�P�s�ǉ�
                                               
  '�Y�����ŏI�s�̂P�s�O�ɍs�ǉ������̂ŁA�Y�����ŏI�s���P�s�O�ɃR�s
  '�`�����W�v�𢓖�Ԋ���A����̑���ɕ����ďW�v���邽�߂ɂ́A����Ԋ���܂��͢���Ԋ��̑O�̊����
  '�@����o�^������ꍇ�A�ŏI�s�̂P�s�O�ɍs�ǉ������Ȃ��Ɗ֐��COUNIF��͈̔�(�s��)��
  '�@�������X�V����Ȃ��B
           
  Range(Cells(nrow, COL_KI), Cells(nrow, COL_COMMENT)).Copy
  Cells(nrow - 1, COL_KI).PasteSpecial Paste:=xlValues, _
              Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False  '�R�s�[���[�h���L�����Z��
  Range(Cells(nrow, COL_KI), Cells(nrow, COL_COMMENT)).ClearContents     '�ŏI�s���N���A
         
  Application.Calculation = xlManual '�Čv�Z���~
            
End Sub


'�ǉ������s�ɐV�K�f�[�^��o�^����v���V�[�W��
Sub RegNewData(nrow As Long)

  Dim data As String       '���[�U�t�H�[���ɓ��͂��ꂽ�o�^�f�[�^

  With UserForm�o�^
             
    '�u���v��2���̏ꍇ�͐擪��0��t����3����
    data = .TextBox1.Text
    If Len(data) = 2 Then
       data = "0" & data
    End If
    Cells(nrow, COL_KI) = "'" & data
            
    '����J�Ŏn�܂�ꍇ�́A�u��ށv��1�A�����łȂ����2
    If Left(data, 1) = "J" Then
      Cells(nrow, COL_CLASS) = 1
    Else
      Cells(nrow, COL_CLASS) = 2
    End If
                    
    '�ǉ�����s�́u���v�ƍŏI�s�́u���v�������ꍇ�A�P���ɁuID�v��+1
    If Cells(nrow, COL_KI) = Cells(nrow - 1, COL_KI) Then
      data = FormatNumber(Right(Cells(nrow - 1, COL_ID), 3)) + 1
    Else   '�Ⴄ�ꍇ�A�V�����u���v�̒ǉ��Ɣ��f���uID�v��������
      data = "001"
    End If
    
    '�uID�v��6�������ɐ��`��2011�N�ύX�\�V��������ID���������ł���悤��
    If Len(data) = 1 Then
      Cells(nrow, COL_ID) = Cells(nrow, COL_KI).Text & "00" & data
    ElseIf Len(data) = 2 Then
      Cells(nrow, COL_ID) = Cells(nrow, COL_KI).Text & "0" & data
    Else
      Cells(nrow, COL_ID) = Cells(nrow, COL_KI).Text & data
    End If
            
    Cells(nrow, COL_NAME) = .TextBox2.Text   '÷���ޯ��2�𢎁�����
    
    Cells(nrow, COL_KANA) = StrConv(.TextBox15.Text, vbNarrow)    '÷���ޯ��15��J�i�������
          
    If OptionButton1.Value = True Then       '����ʣ����j��̏ꍇ
      Cells(nrow, COL_SEX) = OptionButton1.Caption
    ElseIf OptionButton2.Value = True Then   '����ʣ�������̏ꍇ
      Cells(nrow, COL_SEX) = OptionButton2.Caption
    End If
    
    Cells(nrow, COL_ZIP) = .TextBox4.Text & "-" & .TextBox5.Text   '÷���ޯ��4�`5���u���v��
    Cells(nrow, COL_ADDR1) = .TextBox6.Text  '÷���ޯ��6��Z��1���
    Cells(nrow, COL_ADDR2) = .TextBox7.Text  '÷���ޯ��7��Z��2���
    Cells(nrow, COL_ADDR3) = .TextBox8.Text  '÷���ޯ��8��Z��3���
    Cells(nrow, COL_ADDR4) = StrConv(.TextBox9.Text, vbNarrow)  '÷���ޯ��9��Z��4���
    
    '÷���ޯ��10�`12��d�b�ԍ����
    If .TextBox10.Text <> "" Then
      Cells(nrow, COL_TELNO) = StrConv(.TextBox10.Text & "-" & .TextBox11.Text & "-" & .TextBox12.Text, vbNarrow)
    End If
    
    '÷���ޯ��13�`14�𢃁�[���A�h���X���
    If .TextBox13.Text <> "" Then
      Cells(nrow, COL_EMAIL) = StrConv(.TextBox13.Text & "@" & .TextBox14.Text, vbNarrow)
      Cells(nrow, COL_EMAIL).Font.Size = 9
      Cells(nrow, COL_EMAIL).Font.Underline = xlUnderlineStyleNone
    End If

    '÷���ޯ��16�𐮌`���Ģ�������
    Cells(nrow, COL_BUKATSU) = Replace(.TextBox16.Text, "��", "")
              
    '÷���ޯ��17�𐮌`���Ģ�o�g���w���
    Call SetJHScool(nrow)
  
  End With
End Sub


'�V�K�o�^�����f�[�^���y���މ�҈ꗗ�z�t�@�C���ɂ��o�^����v���V�[�W��
Sub RegNewDataToList(nrow As Long)
            
  Dim lrow As Long           '�ŏI�s
            
  Range(Cells(nrow, COL_KI), Cells(nrow, COL_REMARK)).Copy
  With Workbooks("�������}���y���މ�҈ꗗ�z.xls").Sheets("�����")
    lrow = .Range("A" & MEMBER_MAX).End(xlUp).Row + 1
    .Cells(lrow, 1) = Date
    .Cells(lrow, 2).PasteSpecial Paste:=xlValues, _
                       Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  End With
  
  Application.CutCopyMode = False        '�R�s�[���[�h���L�����Z��

End Sub


'�C���o�^�����f�[�^��o�^����v���V�[�W��
Sub RegEditData(nrow As Long)
        
  With UserForm�o�^
    If .TextBox2.Text <> "" Then
      Cells(nrow, COL_NAME) = .TextBox2.Text  '÷���ޯ��2�𢎁�����
    End If
   
    If .TextBox15.Text <> "" Then
      Cells(nrow, COL_KANA) = StrConv(.TextBox15.Text, vbNarrow)  '÷���ޯ��15��J�i�������
    End If
   
    If OptionButton1.Value = True Then      '����ʣ�̢�j��̏ꍇ
      Cells(nrow, COL_SEX) = OptionButton1.Caption
    ElseIf OptionButton2.Value = True Then  '����ʣ�̢����̏ꍇ
      Cells(nrow, COL_SEX) = OptionButton2.Caption
    End If
   
    If .TextBox4.Text <> "" And .TextBox5.Text <> "" Then
      Cells(nrow, COL_ZIP) = .TextBox4.Text & "-" & .TextBox5.Text   '÷���ޯ��4�`5�𢁧���
    End If
   
    If .TextBox6.Text <> "" Then
      Cells(nrow, COL_ADDR1) = .TextBox6.Text  '÷���ޯ��6��Z��1���
    End If
     
    If .TextBox7.Text <> "" Then
      Cells(nrow, COL_ADDR2) = .TextBox7.Text  '÷���ޯ��7��Z��2���
    End If
   
    If .TextBox8.Text <> "" Then
      Cells(nrow, COL_ADDR3) = .TextBox8.Text  '÷���ޯ��8��Z��3���
    End If
   
    If .TextBox9.Text <> "" Then
      Cells(nrow, COL_ADDR4) = StrConv(.TextBox9.Text, vbNarrow)  '÷���ޯ��9��Z��4���
    End If
   
    If .TextBox10.Text <> "" And .TextBox11.Text <> "" And .TextBox12.Text <> "" Then
      Cells(nrow, COL_TELNO) = StrConv(.TextBox10.Text & "-" & .TextBox11.Text & "-" & .TextBox12.Text, vbNarrow)  '÷���ޯ��10�`12��d�b�ԍ����
    End If
   
    If .TextBox13.Text <> "" And .TextBox14.Text <> "" Then
      Cells(nrow, COL_EMAIL) = StrConv(.TextBox13.Text & "@" & .TextBox14.Text, vbNarrow)  '÷���ޯ��13�`14�𢃁�[���A�h���X���
    End If
      
    '÷���ޯ��16�𐮌`���Ģ�������
    Cells(nrow, COL_BUKATSU) = Replace(.TextBox16.Text, "��", "")
              
    '÷���ޯ��17�𐮌`���Ģ�o�g���w���
    Call SetJHScool(nrow)
  End With
End Sub


'���w�Z�̃f�[�^�𐮌`���ēo�^����v���V�[�W��
Sub SetJHScool(nrow As Long)

  Dim data As String       '���w�Z�̃f�[�^
    
  data = TextBox17.Text
  If data = "" Then
  ElseIf Right(data, 3) = "���w�Z" Then
    Cells(nrow, COL_JHSCHOOL) = Replace(data, "���w�Z", "��")
  ElseIf Right(data, 2) = "���w" Then
    Cells(nrow, COL_JHSCHOOL) = Replace(data, "���w", "��")
  ElseIf Right(data, 1) = "��" Then
    Cells(nrow, COL_JHSCHOOL) = data
  Else
    Cells(nrow, COL_JHSCHOOL) = data & "��"
  End If

End Sub


' �o�^��ʂ̢���飃{�^���������ꂽ�ꍇ�̓�����L�q
Private Sub CommandButton3_Click()

  With UserForm�o�^
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



