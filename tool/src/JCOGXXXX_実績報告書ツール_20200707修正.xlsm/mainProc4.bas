Attribute VB_Name = "mainProc4"
Public Sub mainProc4()
  
    '2020.05.15 �쐬
    '�����񍐏��쐬Main�֐�
    
    '�@�\
    '���ѕ񍐏��쐬��ʂ��A�ݒ�擾
    '�����񍐏��֐����s
    
        
    With ThisWorkbook.Sheets("�����񍐏��쐬")
      
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
      
      
        '���ѕ񍐏�template
        Dim Jisitemplate As String
        Jisitemplate = .Range("C11").Value

'�����񍐏�Word�ł́A�쐬�̂�
'        '�i�����X�gLOT�ԍ��ǋL��
'        Dim UpdateLotNum As String
'        UpdateLotNum = .Range("C13").Value
'
'        '�i�����X�gLOT�ԍ��ړ���
'        Dim LOTNAME As String
'        LOTNAME = .Range("C14").Value

'�����񍐏�Word�ł́ALOTNO�K�v

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
    adressAry() = getInitArray(rec, 4, "�����񍐏��쐬")


    '���ѕ񍐏��̓]�L��̃Z���Ԓn�i�[�p
    Dim MapingAry() As String
    '���ѕ񍐂̃Z���Ԓn�ݒ� �u�Ԓn�v�A�s��ԍ��A��ʃV�[�g�����w��
    MapingAry() = getInitArray(rec, 7, "�����񍐏��쐬")


    '�i�����X�g����擾�����A�]�L�p�f�[�^�i�[�p�̔z��ϐ�
    Dim dataAry() As String
    
    '�ݒ�l�̓ǂݍ��ݎ��s
    'dataAry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry(), ListStart, "", "", "", ReadRecord)
    dataAry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry(), ListStart, UpdateLotNum, LOTNAME, LOTLASTNANE, ReadRecord)
     
    
    '���{�񍐏��ւ̓]�L����
    Call PosToFileWrd(OutPutFolder, Jisitemplate, dataAry(), -1, True)

MsgBox "�������I�����܂���"

End Sub

Public Function PosToFileWrd(Filedir As String, tmpfile As String, data() As String, Item As Long, flg As Boolean) As Boolean

    '2020.4.15 �쐬
    
    '�i�����X�g�̔z��f�[�^���w��̓]�L��i�Z���Ԓn�j�ɓ����
    '�w��̃t�@�C�����ŕۑ�
    
    '
    '�����@�F�쐬�������ѕ񍐏��̕ۑ���
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
'    Dim TargetBook As Workbook
    
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


'        Set TargetBook = Workbooks.Open(WorkBookPath(ThisWorkbook.Path, AcptFileDir, TmpFile), False, 0)

        Dim objWord As Word.Application
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = True
        
        Dim TargetDoc As Word.Document
        
        Set TargetDoc = objWord.Documents.Open(WorkBookPath(ThisWorkbook.Path, "", tmpfile), ReadOnly:=False)
    
          
'        '�i�����X�g�t�@�C�������̍�ƂɑJ��
'         With TargetBook.Sheets(SheetName)

         
              'item=-1�̏ꍇ�A�J�E���^�[�ϐ��́A�J�E���g�A�b�v
              If Item = -1 Then
              '
              Else
              'item<>-1�̏ꍇ�A�J�E���^�[�ϐ��́Aitem�̌Œ�l
                 i = Item
              End If
              
'
'              For j = 0 To UBound(Mapping)
'
'               Debug.Print Mapping(j); ":"; data(i, j)
''                .Range(Mapping(j)).Value = data(i, j)
'
'              Next j
              
              
               TargetDoc.Tables(1).Cell(2, 2).Range.Text = data(i, j)
              
              
              
'              '��ԍŌ�̗v�f�ɓ��͂���LOT�ԍ����t�b�^�[�ɓ���
'              UBound(data, 2)
'              j = j + 1

              If flg Then
'                  .PageSetup.RightFooter = data(i, j)

                    
                      Dim Sec As Object
                      Dim ftr As Object
                    
                      For Each Sec In TargetDoc.Sections
                      For Each ftr In Sec.Footers
                        With ftr.Range
                          .Text = data(i, UBound(data, 2))
                          .Paragraphs.Alignment = wdAlignParagraphRight
                        End With
                      Next ftr
                    Next Sec
                    
                    
                    
'                  Debug.Print .PageSetup.RightFooter
              Else
                  
''                  .PageSetup.RightFooter = Null
'                    TargetDoc.Sections.Footers.Range.Text = data(i, j)
              
              End If
              
'         End With
         
         'Debug.Print CInt(flg)
         
         '�i�����X�g�]�L�z��A
         '�u�쐬�v�̏ꍇ�A�Ōォ���O�̗v�f�@�S�v�f���{True(=-1)
         '�u�X�V�v�̏ꍇ�A�Ō�̗v�f
         '�@�ɓ��͂��ꂽ�l���t�@�C�����Ƃ��ĕۑ�����
         
         '2020/06/10
         '�u�����񍐏��쐬�v�����s���܂������A�Z���Ԓn�FD20�Ŏw�肵���t�@�C���������f����܂���ł����B
         'D19�Ɏw�肵����Ë@�֖��Ńt�@�C�����������Ă���悤�ɂ������܂��
         '���m�F�����肢�v���܂��
         '
'         Debug.Print UBound(data, 2) + CInt(flg)
'        TRUE�́u-1�v�Ȃ̂Ł@1-1 = 0 ���@data(i,0)���Q�Ƃ��Ă����@�AD20�����Ă������肪�AD19��


         
'         TargetDoc.SaveAs _
'                    filename:=WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg)))


'
'         TargetDoc.SaveAs _
'                    filename:=WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg)))
                    
                   
'        TargetDoc.SaveAs _
'                    FILENAME:=WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2)) & ".doc")
                   
                   
        Dim strFileName As String
         
        With TargetDoc
        If InStr(data(i, UBound(data, 2) + CInt(flg)), wExt) > 0 Then


            strFileName = WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg)))
         
         Else
       
            strFileName = WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg)) & wExt)
       
         End If
         

            .SaveAs2 _
                    filename:=strFileName, _
                    FileFormat:=wdFormatDocument
                    
            .Close
            
         
         End With
                   
    
'
        Set TargetDoc = Nothing
        
        objWord.Quit

            
        Set objWord = Nothing


         DoEvents


Next_Step:

     Next i
'
  
End Function


