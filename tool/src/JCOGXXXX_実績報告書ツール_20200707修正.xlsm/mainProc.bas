Attribute VB_Name = "mainProc"
Option Explicit

Public Sub mainProc()
  
    '2020.04.15 �쐬
    '���ѕ񍐏��쐬Main�֐�
    
    '�@�\
    '���ѕ񍐏��쐬��ʂ��A�ݒ�擾
    '���ѕ񍐏��쐬�֐����s
    
        
'    With ThisWorkbook.Sheets("���ѕ񍐏��쐬")

     With ThisWorkbook.ActiveSheet
      
        '���ѕ񍐏��o�͐�t�H���_�[�쐬
        Dim OutPutFolder As String
        OutPutFolder = .Range("C2").Value
      
        If FindDirectory(WorkBookPath(ThisWorkbook.Path, OutPutFolder, "")) = True Then
'            If MsgBox("�����̃t�H���_�[�����݂��܂��B�t�H���_�[�̒����m�F���Ă�������", vbOKOnly) = vbOK Then
'
'              Exit Sub
'
'
'            End If
            
        Else
            
              '�t�H���_�[��������΍쐬
              MkDir WorkBookPath(ThisWorkbook.Path, OutPutFolder, "")
        
        End If
      
      
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
      
        '�i�����X�g�ǂݍ��݊J�n�s
        Dim ListStart As Long
        ListStart = 2
      
        '�i�����X�g�ǂݍ��݃t���O��
        Dim ReadRecord As String
        ReadRecord = .Range("C9").Value
      
        '���ѕ񍐏�template�t�H���_�[
        Dim JisitemplateFolder As String
        JisitemplateFolder = .Range("C10").Value
      
        '���ѕ񍐏�template
        Dim Jisitemplate As String
        Jisitemplate = .Range("C11").Value
        
        '���ѕ񍐏�template
        Dim templateSheet As String
        templateSheet = .Range("C12").Value
        
        
        Dim fileExt As String
        fileExt = .Range("G11").Value
        
        Dim fileFmt As Long
        fileFmt = .Range("H11").Value
        
                
        '�i�����X�gLOT�ԍ��ǋL��
        Dim UpdateLotNum As String
        UpdateLotNum = .Range("C13").Value
        
        '�i�����X�gLOT�ԍ��ړ���
        Dim LOTNAME As String
        LOTNAME = .Range("C14").Value
        
        
        '�i�����X�gLOT�ԍ��ڔ���
        Dim LOTLASTNANE As String
        LOTLASTNANE = .Range("C15").Value
      
    End With
    
    '�i�����X�g�A���ѕ񍐏��ݒ�ǂݍ��݊J�n�s
    Dim rec As Long
    rec = 18
    
    '�i�����X�g�̗�����i�[�p�̔z��ϐ�
    Dim adressAry() As String
    '�i�����X�g�̍��ڐݒ� �u��v�̍s�A��ԍ��A��ʃV�[�g�����w��
'    adressAry() = getInitArray(rec, 4, "���ѕ񍐏��쐬")
    adressAry() = getInitArray(rec, 4, ActiveSheet.Name)


    '���ѕ񍐏��̓]�L��̃Z���Ԓn�i�[�p
    Dim MapingAry() As String
    '���ѕ񍐂̃Z���Ԓn�ݒ� �u�Ԓn�v�A�s��ԍ��A��ʃV�[�g�����w��
    MapingAry() = getInitArray(rec, 7, ActiveSheet.Name)


    '�i�����X�g����擾�����A�]�L�p�f�[�^�i�[�p�̔z��ϐ�
    Dim dataAry() As String
    
    '�ݒ�l�̓ǂݍ��ݎ��s
    dataAry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry(), ListStart, UpdateLotNum, LOTNAME, LOTLASTNANE, ReadRecord)
     
    
    '���{�񍐏��ւ̓]�L����
'    Call PosToFile(OutPutFolder, "", Jisitemplate, dataAry(), -1, MapingAry(), "���ѕ�", True)

    If ActiveSheet.Name <> "����0��" Then
    
        Call PosToFile(OutPutFolder, "", Jisitemplate, dataAry(), -1, MapingAry(), templateSheet, True, fileExt, fileFmt)
    Else
        Call PosToFile2(OutPutFolder, "", JisitemplateFolder, dataAry(), -1, MapingAry(), templateSheet, True, fileExt, fileFmt)
    End If
    
    MsgBox "�������I�����܂���"

End Sub


Public Function PosToFile(Filedir As String, AcptFileDir As String, tmpfile As String, data() As String, Item As Long, Mapping() As String, SheetName As String, flg As Boolean, Ext As String, xlsFmt As Long) As Boolean

    '2020.4.15 �쐬
    
    '�i�����X�g�̔z��f�[�^���w��̓]�L��i�Z���Ԓn�j�ɓ����
    '�w��̃t�@�C�����ŕۑ�
    
    '
    '�����@�F�쐬�������ѕ񍐏��̕ۑ���
    '        ��̂������ѕ񍐏��̕ۑ���
    '�@�@�@�@���ѕ񍐃e���v���[�g�t�@�C��

    
    '�@�@�@�@�i�����X�g�]�L�p�z��f�[�^
    
    '�@�@�@�@�u�쐬�v�̔z��\���f�t�H���g
    '�@�@�@�@�@DATA(ONAME,�͏o�󗝔N����,��i��Â̔�p (�͏o��),�l����,�Z�p��,�t�@�C����,LOT�ԍ�)
    '�@�@�@�@�@�Ō�ƍŌォ��1�O�̗v�f�͌Œ�A���͐ݒ�ɂ��ύX��
    
    '�@�@�@�@�u�X�V�v�̔z��\���f�t�H���g
    '�@�@�@�@�@DATA(�ʓY1�F�R�[�h�ԍ�,�t�@�C����)
    '�@�@�@�@�@�Ō�̗v�f�͌Œ�A���͐ݒ�ɂ��ύX��
    
    
    '        �i�����X�g�]�L�p�z��f�[�^�̓Y�����B
    ' �@�@�@�@�@�@-1�̏ꍇ�A�J�E���^�[�ϐ��́A�J�E���g�A�b�v�B
    '           �@-1�ł͖����ꍇ�́A�Œ�l
    
    '�@�@�@�@�]�L��i�Z���Ԓn�j
    '        ���ѕ񍐏��]�L�c�[���̃V�[�g��
    '        �t���O
    '�@�@�@�@�@�@�@�u�쐬�v�̏ꍇ�ATrue(=-1)
    '              �u�X�V�v�̏ꍇ�AFalse(=0)

    '
    '�߂�l�F��ƌ���
    
  
    '�J��Ԃ��񐔂̏����ݒ�B
    '�u�쐬�v�̏ꍇ�A�i�����X�g���擾�������R�[�h����
    '�u�X�V�v�̏ꍇ�A1��̂�
    '�J�E���^�[�ϐ��A�����l��0�B�����A�J��Ԃ��񐔂́A1��
    Dim RecordNo As Long
    RecordNo = 0
    
    '�u�쐬�v�̏ꍇ�A�i�����X�g�̑S���R�[�h�J��Ԃ����߁A���R�[�h�����J�E���^�[�ϐ��ɑ��
    '�u�X�V�v�̏ꍇ�A�i�����X�g1���R�[�h���̃f�[�^�̂݁A�n�����B0����
    If flg Then
      RecordNo = UBound(data)
    End If

    '�i�J��Ԃ��j�]�L����
    Dim i As Long
    Dim j As Long
    Dim TargetBook As Workbook
    
    For i = 0 To RecordNo
      
        '���ѕ񍐃t�@�C���e���v���[�g���I�[�v��
        '�u�쐬�v�̏ꍇ�́A�e���v���[�g
        '�u�X�V�v�̏ꍇ�́A��̂����t�@�C��
        
       ' �Ō�̗v�f�ɓ��͂��ꂽ�l���ANull�������́A"-"�̏ꍇ�́A
       '�������΂��
        If Len(data(i, UBound(data, 2) + CInt(flg))) <= 1 Then
             '�t�@�C�������݂��Ȃ��ꍇ�́A���̏�����
             GoTo Next_Step
            
        End If
        
        Set TargetBook = Workbooks.Open(WorkBookPath(ThisWorkbook.Path, AcptFileDir, tmpfile), False, 0)
    
          
        '���ѕ񍐃t�@�C�������̍�ƂɑJ��
         With TargetBook.Sheets(SheetName)

         
              'item=-1�̏ꍇ�A�J�E���^�[�ϐ��́A�J�E���g�A�b�v
              If Item = -1 Then
              '
              Else
              'item<>-1�̏ꍇ�A�J�E���^�[�ϐ��́Aitem�̌Œ�l
                 i = Item
              End If
              

              For j = 0 To UBound(Mapping)
      
'               Debug.Print Mapping(j); ":"; data(i, j)
                .Range(Mapping(j)).Value = data(i, j)
               
              Next j
              
              '��ԍŌ�̗v�f�ɓ��͂���LOT�ԍ����t�b�^�[�ɓ���
              j = j + 1
              
              If flg Then
                  .PageSetup.RightFooter = data(i, j)
'                  Debug.Print .PageSetup.RightFooter
              Else
                  .PageSetup.RightFooter = Null
              End If
              
         End With
         
         'Debug.Print CInt(flg)
                  
         
         '�]�L�z��A
         '�u�쐬�v�̏ꍇ�A�Ōォ���O�̗v�f�@�S�v�f���{True(=-1)
         '�u�X�V�v�̏ꍇ�A�Ō�̗v�f
         '�@�ɓ��͂��ꂽ�l���t�@�C�����Ƃ��ĕۑ�����
         
        Dim strFileName As String
         
        With TargetBook
        If InStr(data(i, UBound(data, 2) + CInt(flg)), Ext) > 0 Then

'            .SaveAs _
'                    filename:=WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg))), _
'                    FileFormat:=xlsFmt
            strFileName = WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg)))
         
         Else
       
            strFileName = WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg)) & Ext)
       
         End If
         
'            .SaveAs _
'                    filename:=WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg)) & Ext), _
'                    FileFormat:=xlsFmt

             .SaveAs _
                    filename:=strFileName, _
                    FileFormat:=xlsFmt
            .Close
            
         
         End With
         
         DoEvents
         
'�t�@�C�������݂��Ȃ��ꍇ�́A���̏�����
Next_Step:
         
     Next i
'
  
End Function


Public Function PosToFile2(Filedir As String, AcptFileDir As String, tmpfileDir As String, data() As String, Item As Long, Mapping() As String, SheetName As String, flg As Boolean, Ext As String, xlsFmt As Long) As Boolean

    '2020.4.15 �쐬
    
    '�i�����X�g�̔z��f�[�^���w��̓]�L��i�Z���Ԓn�j�ɓ����
    '�w��̃t�@�C�����ŕۑ�
    
    '
    '�����@�F�쐬�������ѕ񍐏��̕ۑ���
    '        ��̂������ѕ񍐏��̕ۑ���
    '�@�@�@�@���ѕ񍐃e���v���[�g�t�@�C��

    
    '�@�@�@�@�i�����X�g�]�L�p�z��f�[�^
    
    '�@�@�@�@�u�쐬�v�̔z��\���f�t�H���g
    '�@�@�@�@�@DATA(ONAME,�͏o�󗝔N����,��i��Â̔�p (�͏o��),�l����,�Z�p��,�t�@�C����,LOT�ԍ�)
    '�@�@�@�@�@�Ō�ƍŌォ��1�O�̗v�f�͌Œ�A���͐ݒ�ɂ��ύX��
    
    '�@�@�@�@�u�X�V�v�̔z��\���f�t�H���g
    '�@�@�@�@�@DATA(�ʓY1�F�R�[�h�ԍ�,�t�@�C����)
    '�@�@�@�@�@�Ō�̗v�f�͌Œ�A���͐ݒ�ɂ��ύX��
    
    
    '        �i�����X�g�]�L�p�z��f�[�^�̓Y�����B
    ' �@�@�@�@�@�@-1�̏ꍇ�A�J�E���^�[�ϐ��́A�J�E���g�A�b�v�B
    '           �@-1�ł͖����ꍇ�́A�Œ�l
    
    '�@�@�@�@�]�L��i�Z���Ԓn�j
    '        ���ѕ񍐏��]�L�c�[���̃V�[�g��
    '        �t���O
    '�@�@�@�@�@�@�@�u�쐬�v�̏ꍇ�ATrue(=-1)
    '              �u�X�V�v�̏ꍇ�AFalse(=0)

    '
    '�߂�l�F��ƌ���
    
  
    '�J��Ԃ��񐔂̏����ݒ�B
    '�u�쐬�v�̏ꍇ�A�i�����X�g���擾�������R�[�h����
    '�u�X�V�v�̏ꍇ�A1��̂�
    '�J�E���^�[�ϐ��A�����l��0�B�����A�J��Ԃ��񐔂́A1��
    Dim RecordNo As Long
    RecordNo = 0
    
    '�u�쐬�v�̏ꍇ�A�i�����X�g�̑S���R�[�h�J��Ԃ����߁A���R�[�h�����J�E���^�[�ϐ��ɑ��
    '�u�X�V�v�̏ꍇ�A�i�����X�g1���R�[�h���̃f�[�^�̂݁A�n�����B0����
    If flg Then
      RecordNo = UBound(data)
    End If

    '�i�J��Ԃ��j�]�L����
    Dim i As Long
    Dim j As Long
    Dim TargetBook As Workbook
    
    Dim templatefile As String
    
    For i = 0 To RecordNo
      
        '���ѕ񍐃t�@�C���e���v���[�g���I�[�v��
        '�u�쐬�v�̏ꍇ�́A�e���v���[�g
        '�u�X�V�v�̏ꍇ�́A��̂����t�@�C��
        
       ' �Ō�̗v�f�ɓ��͂��ꂽ�l���ANull�������́A"-"�̏ꍇ�́A
       '�������΂��
        If Len(data(i, UBound(data, 2) + CInt(flg))) <= 1 Then
             '�t�@�C�������݂��Ȃ��ꍇ�́A���̏�����
             GoTo Next_Step
            
        End If
        
        
       If FindFile(WorkBookPath(ThisWorkbook.Path, tmpfileDir, data(i, 1))) = False Then
           GoTo Next_Step
       End If
        
        Set TargetBook = Workbooks.Open(WorkBookPath(ThisWorkbook.Path, tmpfileDir, data(i, 1)), False, 0)
          
        '���ѕ񍐃t�@�C�������̍�ƂɑJ��
         With TargetBook.Sheets(SheetName)

         
              'item=-1�̏ꍇ�A�J�E���^�[�ϐ��́A�J�E���g�A�b�v
              If Item = -1 Then
              '
              Else
              'item<>-1�̏ꍇ�A�J�E���^�[�ϐ��́Aitem�̌Œ�l
                 i = Item
              End If
              

              For j = 0 To UBound(Mapping)
      
'               Debug.Print Mapping(j); ":"; data(i, j)
                .Range(Mapping(j)).Value = data(i, j)
               
              Next j
              
              '��ԍŌ�̗v�f�ɓ��͂���LOT�ԍ����t�b�^�[�ɓ���
              j = j + 1
              
              If flg Then
                  .PageSetup.RightFooter = data(i, j)
'                  Debug.Print .PageSetup.RightFooter
              Else
                  .PageSetup.RightFooter = Null
              End If
              
         End With
         
         'Debug.Print CInt(flg)
                  
         
         '�]�L�z��A
         '�u�쐬�v�̏ꍇ�A�Ōォ���O�̗v�f�@�S�v�f���{True(=-1)
         '�u�X�V�v�̏ꍇ�A�Ō�̗v�f
         '�@�ɓ��͂��ꂽ�l���t�@�C�����Ƃ��ĕۑ�����
         
        Dim strFileName As String
         
        With TargetBook
        If InStr(data(i, UBound(data, 2) + CInt(flg)), Ext) > 0 Then

'            .SaveAs _
'                    filename:=WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg))), _
'                    FileFormat:=xlsFmt
            strFileName = WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg)))
         
         Else
       
            strFileName = WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg)) & Ext)
       
         End If
         
'            .SaveAs _
'                    filename:=WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg)) & Ext), _
'                    FileFormat:=xlsFmt

             .SaveAs _
                    filename:=strFileName, _
                    FileFormat:=xlsFmt
            .Close
            
         
         End With
         
         DoEvents
         
'�t�@�C�������݂��Ȃ��ꍇ�́A���̏�����
Next_Step:
         
     Next i
'
  
End Function

