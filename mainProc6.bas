Attribute VB_Name = "mainProc6"
Public Sub mainProc6()
  
    '2020.05.27 �쐬
    '�t�H���_�[�쐬Main�֐�
    
    '�@�\
    '���ѕ񍐏��쐬��ʂ��A�ݒ�擾
    '�t�H���_�[�\����i�����X�g�ɓ��͗p�֐����s
    
        
    With ThisWorkbook.Sheets("�Ɩ��t�H���_�[�쐬")
      
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
      
        '�i�����X�g�ǂݍ��݊J�n�s
        Dim ListStart As Long
        ListStart = 2
      
        '�i�����X�g�ǂݍ��݃t���O��
        Dim ReadRecord As String
        ReadRecord = .Range("C9").Value
      
      
'        '���ѕ񍐏�template
'        Dim Jisitemplate As String
'        Jisitemplate = .Range("C11").Value
                
        '�i�����X�gLOT�ԍ��ǋL��
        Dim UpdateLotNum As String
        UpdateLotNum = .Range("C13").Value
        
        '�i�����X�gLOT�ԍ��ړ���
        Dim LOTNAME As String
        LOTNAME = .Range("C14").Value
        
        
        '�i�����X�gLOT�ԍ��ڔ���
        Dim LOTLASTNANE As String
        LOTLASTNANE = .Range("C15").Value
      
        'TOP�̃t�H���_�[
        Dim folderTop As String
        folderTop = .Range("L2").Value
        
        '������
        Dim folderPrtName As String
        folderPrtName = .Range("L3").Value
        
        '�i�����X�gOCODE��
        Dim OCODE As String
        OCODE = .Range("L4").Value
        
        '�i�����X�gONAME��
        Dim ONAME As String
        ONAME = .Range("L5").Value
        
        '�i�����X�g�t�H���_�[�ǋL��
        Dim UpdateFolder As String
        UpdateFolder = .Range("L8").Value
        
        
      
    End With
    
    '�i�����X�g�A���ѕ񍐏��ݒ�ǂݍ��݊J�n�s
    Dim rec As Long
    rec = 10
    
'    '�i�����X�g�̗�����i�[�p�̔z��ϐ�
'    Dim adressAry() As String
'    '�i�����X�g�̍��ڐݒ� �u��v�̍s�A��ԍ��A��ʃV�[�g�����w��
'    adressAry() = getInitArray(rec, 4, "���ѕ񍐏��쐬")
    

    '�i�����X�g�ɏo�͂�����STEP�̏��擾
    Dim adressAry() As String
    '�i�����X�g�̍��ڐݒ� �u��v�̍s�A��ԍ��A��ʃV�[�g�����w��
    adressAry() = getInitArray(rec, 11, "�Ɩ��t�H���_�[�쐬")

 
     '�Ɩ��t�H���_�[TOP�̕�����쐬
     '�N�x�̒ǉ�
            
        folderTop = WorkBookPath(folderTop, "", "")

      If FindDirectory(folderTop) = True Then
        '
      Else
            MkDir folderTop
      End If
        
     '�������̒ǉ�
        folderTop = WorkBookPath(folderTop, folderPrtName, "")
        
      If FindDirectory(folderTop) = True Then
        '
      Else
        MkDir folderTop
     End If
        
    Dim adressAry2() As String
    Dim aA2 As Variant
     '�i�����X�g�X�V�@�t�H���_�[������
    adressAry2 = MkFolder(ShintyokuFile, ShintyokuSheet, adressAry(), ListStart, UpdateLotNum, LOTNAME, LOTLASTNANE, ReadRecord, folderTop, OCODE, ONAME, UpdateFolder)
    

'    �i�����X�g�ɒl��}���������Ԃ�

    With ThisWorkbook.Sheets("�Ɩ��t�H���_�[�쐬")
    Dim i As Long
    For Each aA2 In adressAry2
        .Cells(rec + 1 + i, 11 + 1).Value = aA2
        i = i + 1
    Next
    End With

'    '���ѕ񍐏��̓]�L��̃Z���Ԓn�i�[�p
'    Dim MapingAry() As String
'    '���ѕ񍐂̃Z���Ԓn�ݒ� �u�Ԓn�v�A�s��ԍ��A��ʃV�[�g�����w��
'    MapingAry() = getInitArray(rec, 7, "���ѕ񍐏��쐬")


'    '�i�����X�g����擾�����A�]�L�p�f�[�^�i�[�p�̔z��ϐ�
'    Dim dataAry() As String
    
'    '�ݒ�l�̓ǂݍ��ݎ��s
'    dataAry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry(), ListStart, UpdateLotNum, LOTNAME, LOTLASTNANE, ReadRecord)
     
    '���{�񍐏��ւ̓]�L����
'    Call PosToFile(OutPutFolder, "", Jisitemplate, dataAry(), -1, MapingAry(), "���ѕ�", True)

MsgBox "�������I�����܂���"

End Sub


Public Function MkFolder(filename As String, SheetName As String, Mapping() As String, rec As Long, LotCol As String, LotNames As String, LotLastName As String, ReadingFlg As String, folderTop As String, OCODE As String, ONAME As String, UpdateFolder As String) As String()

    '2020.5.25 �쐬

    
    '�i�����X�g���J���i�����N�X�V�̃��b�Z�[�W��\�������Ȃ��j
    '�i�����X�g����ǂݍ��ރ��R�[�h�����m�F����
    '�w��̗�̃f�[�^��1�s���z��ɓ����
    
    '
    '�����@�F�i�����X�g�̃t�@�C����
    '        �i�����X�g�̃V�[�g��
    '�@�@�@�@�i�����X�g�ɓ��͂���f�[�^�i�G�N�Z��:��j
    '�@�@�@  �i�����X�g�̓ǂݍ��݊J�n�ʒu
    '�@�@�@�@�i�����X�g�ɒǋL����Lot�ԍ��X�V��
    '�@�@�@�@�i�����X�g�ɒǋL����Lot�ԍ��̐ړ���
    '�@�@�@�@�i�����X�g�ɒǋL����Lot�ԍ��̐ڔ���
    '        �i�����X�g����ǂݍ��ރ��R�[�h�𔻒肷��t���O��
    '�@�@�@�@�t�H���_�[���쐬����ꏊ
    '�@�@�@�@�i�����X�g��OCODE�̗�
    '�@�@�@�@�i�����X�g��ONAME�̗�
    '�@�@�@�@�i�����X�g�ɒǋL����t�H���_�[
    
    
'    '��Ɨp�̓]�L�p�f�[�^�i�[�p�z��ϐ�
'    Dim Work() As String
'    Dim Work2() As String
'
    '�i�����X�g����ǂݍ��ރ��R�[�h��
    Dim l As Long
    
    '�J�E���^�[�ϐ�
    Dim r As Long
    r = rec
    
    '�i�����X�g�t�@�C���I�[�v��
    Dim TargetBook As Workbook
    Set TargetBook = Workbooks.Open(WorkBookPath(ThisWorkbook.Path, "", filename), False, 0)


    '�i�����X�g�t�@�C�������̍�ƂɑJ��
    With TargetBook.Sheets(SheetName)
          
        '�i�����X�g����ǂݍ��ރ��R�[�h���m�F
        .Select
        .Cells(1048576, rec).Activate
        
        Selection.End(xlUp).Select
        l = Selection.Row
        
        
'        '���R�[�h���ƃf�[�^���ځi�����j�̐�����A�z����Đݒ�
'        '���ѕ񍐏��X�V�̏ꍇ�́A�v�f�́A���̂܂܁B
'        '���ѕ񍐏��쐬�̏ꍇ�́A�v�f��1�ǉ����܂��B
'        '��5�����A��6�������󗓂ł͖����ꍇ�ALot�ԍ��p�̗v�f��ǉ�����B
'
'        If Len(LotCol) = 0 And Len(LotNames) = 0 Then
'            ReDim Work(l - r, UBound(Mapping))
'        Else
'            ReDim Work(l - r, UBound(Mapping) + 1)
'        End If
'
        '�J�E���^�[�ϐ�
        Dim i As Long
        
        '�J�E���^�[�ϐ��A�ǂݍ��܂Ȃ����R�[�h��
        Dim j As Long
        j = 0
        
        
        '�Ăяo�����ɁA�i�����X�g�̗�ԍ���Ԃ��l���i�[����z��
        Dim oMp() As String
        ReDim oMp(UBound(Mapping))
        
        '���[�v�I�������F�������郌�R�[�h���A�i�����X�g�̃��R�[�h���𒴂�����I��
        Do Until l < r
           
            i = 0
            
            
            '����[�i�����X�g����ǂݍ��ރ��R�[�h�𔻒肷��t���O��]�ɒl������ꍇ�A
            '�t���O�Ɏw�肳�ꂽ���ڂ̒l�ɂ���āA�f�[�^�̓ǂݍ��ނ��ۂ��̔�����s��
            If Len(ReadingFlg) > 0 Then
            
                '�w�肳�ꂽ��̃f�[�^�ɒl���݂�΁A�������s��
                '�Ȃ���΁A�ǂݍ��܂Ȃ��B
                
                If Len(.Range(ReadingFlg & r).Value) = 0 Then
                
                    '�ǂݍ��܂Ȃ����R�[�h��
                    j = j + 1
                    '�f�[�^��ǂݍ��܂ȂȂ����̏�����
                    GoTo Next_Step
                
                End If
            
            End If
            
'            '�Z���Ԓn�p�̕ϐ�
'            Dim adress As Variant
'
'            For Each adress In Mapping
'
'               '�f�[�^���ځi�����j����Z���Ԓn�𐶐�
'               '�]�L�p�f�[�^��z��Ɋi�[
'               Work(r - rec, i) = .Range(adress & r).Value
'
'               i = i + 1
'
'            Next adress
            
            '���ѕ񍐏��X�V�̏ꍇ�́A�������Ȃ��B
            '���ѕ񍐏��쐬�̏ꍇ�́ALOT�ԍ��������āA�]�L�p�̔z��A�i�����X�g�ɒǋL
            
            If Len(LotCol) = 0 And Len(LotNames) = 0 Then
            
               '
            
'            '2020/5/19 �ǉ�
'            ElseIf Len(LotCol) = 0 And Len(LotNames) <> 0 Then
'
'                Work(r - rec, i) = LotNames & "_" & .Range(LotLastName & r).Value
            
            Else
               
               '��5�����A��6�������󗓂ł͖����ꍇ�i���ѕ񍐏��쐬�j�A��Ԍ��̗v�f��Lot�ԍ���ǉ�����B
               
'                Work(r - rec, i) = LotNum(LotNames, CStr(r - rec + 1))
'               .Range(LotCol & r).Value = LotNum(LotNames, CStr(r - rec + 1))

                Dim workPath As String
                Dim workPathSub As String
                
                workPath = WorkBookPath(folderTop, .Range(OCODE & r).Value & "_" & .Range(ONAME & r).Value, "")
                
                If FindDirectory(workPath) = True Then
                    '
                Else
                    MkDir workPath
                End If
                
'                .Range(UpdateFolder & r).Value = workPath


                Dim ofset As Long
                ofset = 0
                Dim mp As Variant
                                
                For Each mp In Mapping

                                       
                    If r = rec Then
                    


                        Dim w As String
                        Dim k As Long
                        w = .Range(.Range(UpdateFolder & r).Address).Offset(0, ofset).Address
'                        Debug.Print Mid(w, InStr(w, "$") + 1, InStr(InStr(w, "$") + 1, w, "$") - 2)
                        oMp(k) = Mid(w, InStr(w, "$") + 1, InStr(InStr(w, "$") + 1, w, "$") - 2)
                        k = k + 1
                    End If
                    
                    workPathSub = WorkBookPath(workPath, CStr(mp), "")
                    
                    If FindDirectory(workPathSub) = True Then
                    '
                    Else
            
            
'                    Debug.Print workPathSub
                      '�t�H���_�[��������΍쐬
                      MkDir workPathSub
                      
                      
                      .Range(.Range(UpdateFolder & r).Address).Offset(0, ofset).Value = workPathSub
        
        
                    End If
      
                    
                    
                    ofset = ofset + 1
                Next
                


'                Work(r - rec, i) = LotNames & "_" & .Range(LotLastName & r).Value
               .Range(LotCol & r).Value = LotNames & "_" & .Range(LotLastName & r).Value


            End If
            
Next_Step:

            r = r + 1

         Loop
    
    End With
    
    '�i�����X�g�����B�������[�J��
    TargetBook.Close SaveChanges:=True
    Set TargetBook = Nothing
    
    DoEvents
    
'
'    '�z��Work����s�v�ȋ󔒍s���폜���邽�߂̏���
'    ReDim Work2(UBound(Work, 1) - j, UBound(Work, 2))
'
'    Dim k As Long
'    Dim m As Long
'    m = 0
'    Dim n As Long
'    n = 0
'
'    For k = 0 To UBound(Work, 1)
'
'        If Len(Work(k, 0)) <> 0 Then
'
'            For m = 0 To UBound(Work, 2)
'                Work2(n, m) = Work(k, m)
'            Next m
'
'            n = n + 1
'
'        End If
'    Next k
    
    '��Ɨp�̔z��f�[�^���Ăяo�����ɕԂ��B
    MkFolder = oMp

End Function
