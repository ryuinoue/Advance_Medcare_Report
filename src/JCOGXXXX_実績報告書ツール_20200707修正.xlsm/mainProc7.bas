Attribute VB_Name = "mainProc7"
Public Sub mainProc7()
  
    '2020.06.02 �쐬
    '�t�@�C���R�s�[Main�֐�
    
    '�@�\
    '�e��ʂ��A�ݒ�擾
    '�ݒ��̃t�H���_�[�Ƀt�@�C�����R�s�[
    
    
    
    '�t�@�C�����i�[��
    Dim FileCol As String
    
    '�ꎞ�t�H���_�[���i�[
    Dim FldName As String

        
    
    Dim rng As Range
        
    With ThisWorkbook.ActiveSheet
    
        '�ꎞ�t�H���_�[�̃t�H���_�[�����擾
        With .Columns(RETUB)
        
            '���S��v
            Set rng = .Find(FOLDERNAME, LookAt:=xlWhole)
            ', xlValues, xlWhole, xlByColumns, xlNext
            
            If rng Is Nothing Then
            
                MsgBox "�ݒ���m�F���Ă�������"
                Exit Sub
                
            End If
            

            FldName = Range(rng.Address).Offset(0, 1).Value
            
        End With
    
    
        '�t�@�C�����̕ۑ�����擾
        With .Columns(RETUC)
        
            Set rng = .Find(filename)
            ', xlValues, xlWhole, xlByColumns, xlNext
            If rng Is Nothing Then
            
                MsgBox "�ݒ���m�F���Ă�������"
                Exit Sub
                
            End If
            

            FileCol = Range(rng.Address).Offset(0, 1).Value
            
        End With
        
        
        
        '�ǂݍ��ފg���q
        Dim kakutyoushi As String
        kakutyoushi = .Range("U19").Value
      
      
      
      
'        '���ѕ񍐏��o�͐�t�H���_�[�쐬
'        Dim OutPutFolder As String
'        OutPutFolder = .Range("C2").Value
'
'        If FindDirectory(WorkBookPath(ThisWorkbook.Path, OutPutFolder, "")) = True Then
''            If MsgBox("�����̃t�H���_�[�����݂��܂��B�t�H���_�[�̒����m�F���Ă�������", vbOKOnly) = vbOK Then
''
''              Exit Sub
''
''
''            End If
'
'        Else
'
'              '�t�H���_�[��������΍쐬
'              MkDir WorkBookPath(ThisWorkbook.Path, OutPutFolder, "")
'
'        End If
      
      
        '�ǂݍ��ސi�����X�g�t�@�C����
        Dim ShintyokuFile As String
        ShintyokuFile = .Range("C5").Value

        If FindFile(WorkBookPath(ThisWorkbook.Path, "", ShintyokuFile)) = False Then

            If MsgBox("�i�����X�g�����݂��Ă��܂���", vbOKOnly) = vbOK Then

                Exit Sub

            End If

        End If
      
        '�ǂݍ��ސi�����X�g�V�[�g��
        Dim ShintyokuSheet As String
        ShintyokuSheet = .Range("C7").Value
'
        '�i�����X�g�ǂݍ��݊J�n�s
        Dim ListStart As Long
        ListStart = 2
'
        '�i�����X�g�ǂݍ��݃t���O��
        Dim ReadRecord As String
        ReadRecord = .Range("C9").Value
      
      
'        '���ѕ񍐏�template
'        Dim Jisitemplate As String
'        Jisitemplate = .Range("C11").Value
'
'        '�i�����X�gLOT�ԍ��ǋL��
'        Dim UpdateLotNum As String
'        UpdateLotNum = .Range("C13").Value
'
'        '�i�����X�gLOT�ԍ��ړ���
'        Dim LOTNAME As String
'        LOTNAME = .Range("C14").Value
'
'
'        '�i�����X�gLOT�ԍ��ڔ���
'        Dim LOTLASTNANE As String
'        LOTLASTNANE = .Range("C15").Value
      
    End With
    
    '�ݒ�ǂݍ��݊J�n�s
    Dim rec As Long
    rec = 18

'
'    '�i�����X�g�̗�����i�[�p�̔z��ϐ�
    Dim adressAry() As String
'    '�i�����X�g�̍��ڐݒ� �u��v�̍s�A��ԍ��A��ʃV�[�g�����w��
    adressAry() = getInitArray(rec, 17, ThisWorkbook.ActiveSheet.Name)
'
'
'    '���ѕ񍐏��̓]�L��̃Z���Ԓn�i�[�p
'    Dim MapingAry() As String
'    '���ѕ񍐂̃Z���Ԓn�ݒ� �u�Ԓn�v�A�s��ԍ��A��ʃV�[�g�����w��
'    MapingAry() = getInitArray(rec, 7, "���ѕ񍐏��쐬")
'

'�t�@�C�����A�o�͐�t�H���_�[���i�[����1�����z����쐬(������́A��ʗv���Ɗm�肵�Ă���̂ŁA2�s�A1��̔z����ŏ�����p�ӂ���)
    Dim adressAry2(1) As String
    adressAry2(0) = FileCol
    adressAry2(1) = adressAry(0)

'
    '�i�����X�g����擾�����A�]�L�p�f�[�^�i�[�p�̔z��ϐ�
    Dim dataAry() As String
'
    '�ݒ�l�̓ǂݍ��ݎ��s
    dataAry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry2(), ListStart, "", "", "", ReadRecord)

'   �t�@�C���R�s�[
    Call FloderToCopy(FldName, dataAry(), kakutyoushi)

'    '���{�񍐏��ւ̓]�L����
'    Call PosToFile(OutPutFolder, "", Jisitemplate, dataAry(), -1, MapingAry(), "���ѕ�", True)

MsgBox "�������I�����܂���"

End Sub


Public Sub FloderToCopy(FldName As String, data() As String, kakutyoushi As String)

    '2020.6.2 �쐬
    
    '�i�����X�g�̔z��f�[�^�����Ɉꎞ�t�H���_�[����A�ړ���̃t�H���_�[�Ƀt�@�C�����R�s�[����
    '�w��̃t�@�C�����ŕۑ�
    
    '
    '�����@�F�ꎞ�t�H���_�[��
    '        �t�@�C�����ƈړ���̃t�H���_�[
    '�@�@�@�@���ѕ񍐃e���v���[�g�t�@�C��
    
  
    '�J��Ԃ��񐔂̏����ݒ�B
    ' �i�����X�g���擾�������R�[�h����
    
    
    Dim RecordNo As Long
    RecordNo = 0
    
    '�u�쐬�v�̏ꍇ�A�i�����X�g�̑S���R�[�h�J��Ԃ����߁A���R�[�h�����J�E���^�[�ϐ��ɑ��
    
    
      RecordNo = UBound(data)


    '�i�J��Ԃ��j�t�@�C���R�s�[����
    Dim i As Long
    Dim j As Long
    
'    Dim TargetBook As Workbook
    
    For i = 0 To RecordNo
      
'        '���ѕ񍐃t�@�C���e���v���[�g���I�[�v��
'        '�u�쐬�v�̏ꍇ�́A�e���v���[�g
'        '�u�X�V�v�̏ꍇ�́A��̂����t�@�C��
'        Set TargetBook = Workbooks.Open(WorkBookPath(ThisWorkbook.Path, AcptFileDir, TmpFile), False, 0)


'        �ꎞ�t�H���_�[���Ƀt�@�C�������鎖���m�F����B
'�@�@�@�@�ꎞ�t�H���_�[�́A�c�[���Ɠ����f�B���N�g���[�ɍ쐬�����B���̗v���͌Œ�Ƃ���B

        If FindFile(WorkBookPath(ThisWorkbook.Path, FldName, data(i, 0) & Trim(kakutyoushi))) = False Then
        
            '�t�@�C�������݂��Ȃ��ꍇ�́A���̏�����
             GoTo Next_Step
            
        End If
    
          Dim strFile2 As String
            strFile2 = WorkBookPath(data(i, 1), "", data(i, 0) & Trim(kakutyoushi))
            
           Dim strFile As String
           strFile = WorkBookPath(ThisWorkbook.Path, FldName, data(i, 0) & Trim(kakutyoushi))
          FileCopy strFile, strFile2
          
          
'        '���ѕ񍐃t�@�C�������̍�ƂɑJ��
'         With TargetBook.Sheets(SheetName)
'
'              'item=-1�̏ꍇ�A�J�E���^�[�ϐ��́A�J�E���g�A�b�v
'              If Item = -1 Then
'              '
'              Else
'              'item<>-1�̏ꍇ�A�J�E���^�[�ϐ��́Aitem�̌Œ�l
'                 i = Item
'              End If
'
'
'              For j = 0 To UBound(Mapping)
'
''               Debug.Print Mapping(j); ":"; data(i, j)
'                .Range(Mapping(j)).Value = data(i, j)
'
'              Next j
'
'              '��ԍŌ�̗v�f�ɓ��͂���LOT�ԍ����t�b�^�[�ɓ���
'              j = j + 1
'
'              If flg Then
'                  .PageSetup.RightFooter = data(i, j)
''                  Debug.Print .PageSetup.RightFooter
'              Else
'                  .PageSetup.RightFooter = Null
'              End If
'
'         End With
         
         'Debug.Print CInt(flg)
         
         '�]�L�z��A
         '�u�쐬�v�̏ꍇ�A�Ōォ���O�̗v�f�@�S�v�f���{True(=-1)
         '�u�X�V�v�̏ꍇ�A�Ō�̗v�f
         '�@�ɓ��͂��ꂽ�l���t�@�C�����Ƃ��ĕۑ�����
'         TargetBook.SaveAs _
'                    FILENAME:=WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg)) & Ext), _
'                    FileFormat:=xlsFmt
'         TargetBook.Close
         
Next_Step:
         
         DoEvents
         
     Next i
'
  
End Sub





