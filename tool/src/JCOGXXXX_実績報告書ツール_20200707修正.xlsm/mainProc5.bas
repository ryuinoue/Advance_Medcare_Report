Attribute VB_Name = "mainProc5"
Public Sub mainProc5()
  
    '2020.05.19 �쐬
    '���ғo�^�󋵈ꗗ�쐬Main�֐�
    
    '�@�\
    '���ғo�^�󋵈ꗗ�쐬��ʂ��A�ݒ�擾
    '���ғo�^�󋵈ꗗ�쐬�֐����s
    
        
    With ThisWorkbook.Sheets("���ғo�^�󋵈ꗗ�쐬")
      
        '���ғo�^�󋵈ꗗ�o�͐�t�H���_�[�쐬
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
      
        '�m�F���ʏo�͐�i�I�t�Z�b�g�Ԓn�j
        Dim ofset As Long
        ofset = .Range("C10").Value
      
      
        '���ғo�^�󋵈ꗗtemplate
        Dim Jisitemplate As String
        Jisitemplate = .Range("C11").Value
                
        '�i�����X�gLOT�ԍ��ǋL��
        Dim UpdateLotNum As String
        UpdateLotNum = .Range("C13").Value
        
        '�i�����X�gLOT�ԍ��ړ���
        Dim LOTNAME As String
        LOTNAME = .Range("C14").Value
        
        
        '�i�����X�gLOT�ԍ��ڔ���
        Dim LOTLASTNANE As String
        LOTLASTNANE = .Range("C15").Value
        
        
        '------------------------------------------------------------------------------------------------
        'MDB���
        'SAS�o�^���f�[�^.MDB�ۑ���t�H���_�[
        Dim MDBDir As String
        MDBDir = .Range("G3").Value


        'SAS�o�^���f�[�^.MDB�̃t�@�C����
        Dim mdbFile As String
        mdbFile = .Range("H4").Value


        '�e�[�u����
        Dim mdbTbl As String
        mdbTbl = .Range("H5").Value
        
        '���я�
        Dim tblorder As String
        tblorder = .Range("I13").Value
        
        '���ғo�^�󋵈ꗗ�̒��o�����Z���Ԓn�i�[�p
        Dim whereAry() As String
        
        '���ѕ񍐂̃Z���Ԓn�ݒ� �u�Ԓn�v�A�s��ԍ��A��ʃV�[�g�����w��
        whereAry() = getInitArray(7, 10, "���ғo�^�󋵈ꗗ�쐬")
        
        '------------------------------------------------------------------------------------------------
      
    End With
    
    '�i�����X�g�A���ѕ񍐏��ݒ�ǂݍ��݊J�n�s
    Dim rec As Long
    rec = 19
    
    '�i�����X�g�̗�����i�[�p�̔z��ϐ�
    Dim adressAry() As String
    '�i�����X�g�̍��ڐݒ� �u��v�̍s�A��ԍ��A��ʃV�[�g�����w��
    adressAry() = getInitArray(rec, 4, "���ғo�^�󋵈ꗗ�쐬")


    '���ѕ񍐏��̓]�L��̃Z���Ԓn�i�[�p
    Dim MapingAry() As String
    '���ѕ񍐂̃Z���Ԓn�ݒ� �u�Ԓn�v�A�s��ԍ��A��ʃV�[�g�����w��
    MapingAry() = getInitArray(rec, 7, "���ғo�^�󋵈ꗗ�쐬")


    '------------------------------------------------------------------------------------------------
    '���ѕ񍐏��̓]�L��̃Z���Ԓn�i�[�p
    Dim itemAry() As String
    '���ѕ񍐂̃Z���Ԓn�ݒ� �u�Ԓn�v�A�s��ԍ��A��ʃV�[�g�����w��
    itemAry() = getInitArray(18, 9, "���ғo�^�󋵈ꗗ�쐬")


    '�i�����X�g����擾�����A�{�ݏ��̔z��ϐ�
    Dim dataAry() As String
    
    
    '�ݒ�l�̓ǂݍ��ݎ��s
    '���i�����X�g����AMDB�t�@�C�����{�ݖ��̃f�[�^�𒊏o���邽�߂̃L�[�R�[�h�uOCODE�v
    '  LOTNUM�̔z����擾
    
'     dataAry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry(), ListStart, "", LOTNAME, LOTLASTNANE, ReadRecord)
    dataAry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry(), ListStart, "", "", "", ReadRecord)
     
    
    'MDB�iSAS�o�^���f�[�^�j����擾�����]�L�p�f�[�^�i�[�p�̔z��ϐ�
    Dim mdbdataAry() As String
    Dim mdbrec As Long
    Dim o As Long
    Dim LOTNUM As String

    For o = 0 To UBound(dataAry, 1)
    
        'MDB����荞��
        Call InputArrayMDB(MDBDir, mdbFile, mdbTbl, tblorder, itemAry(), whereAry(), dataAry(), o, mdbdataAry(), LOTNUM, mdbrec)
        
        If mdbrec > 0 Then
        
            '���ғo�^�󋵈ꗗ�ւ̓]�L����
            Call PosToFileMult(OutPutFolder, Jisitemplate, mdbdataAry(), dataAry(), o, "Sheet1")

        End If
        
        '�z��̉��
        Erase mdbdataAry
        
'�@�@�@�@�i�����X�g�ɏo�͂���o�^���́A���l�Ƃ��Ďg�p����́u���v�́A�폜
'        Call InPutQuery(ShintyokuFile, ShintyokuSheet, LOTNUM, mdbrec & "��", adressAry(UBound(adressAry)), ofset)
        Call InPutQuery(ShintyokuFile, ShintyokuSheet, LOTNUM, CStr(mdbrec), adressAry(UBound(adressAry)), ofset)
        
        
    Next o
'Public Function PosToFileMult(Filedir As String, TmpFile As String, DATA() As String, LOTNUM As String, SheetName As String) As Boolean

MsgBox "�������I�����܂���"

End Sub

Public Function InputArrayMDB(mdbir As String, mdbFile As String, mdbTbl As String, tblOdr As String, itemAry() As String, whereAry() As String, dataAry() As String, o As Long, rsAry() As String, LOTNUM As String, mdbrec As Long) As String

    '2020.5.19 �쐬
    
    'MDB�iSAS�o�^���f�[�^�j���J�����R�[�h�Z�b�g���쐬
    'MDB����ǂݍ��ރ��R�[�h�����m�F����
    '�w��̗�̃f�[�^��1�s���z��ɓ����
    
    '
    '�����@�FMDB�̃t�@�C���̕ۑ���f�B���N�g��
    '        MDB�̃t�@�C����
    '�@�@�@�@�e�[�u����
    '�@�@�@  ���o����
    '�@�@�@�@���o����
    '�@�@�@�@�ݒ�l�@��OCODE
    '
    '
    
    '�߂�l�FSQL��
    

    'MDB�t�@�C�����f�[�^���o�pSQL���쐬

  
    Dim workItem As String
    
    Dim Item As Variant
    
    For Each Item In itemAry
    
        workItem = workItem & "," & Item
    
    Next
     
     workItem = Mid(workItem, Len(",") + 1, 1000)
    
    Dim workFrom As String
    
    Dim wh As Variant
    
    For Each wh In whereAry
    
        workFrom = workFrom & " and " & wh
    
    Next
    
    workFrom = Mid(workFrom, Len(" and ") + 1, 1000)
    
    'OCODE�́A�z��̈�ԍŏ��ɓ����Ă��邱�Ƃ�O��Ƃ���B
    Dim strsql As String
    strsql = "select " & workItem & " from " & mdbTbl & " where " & workFrom & " and OCODE='" & dataAry(o, 0) & "' order by " & tblOdr
    
'    Debug.Print strsql
    
    '------------------------------------------------------------------------------------------------
    
    '------------------------------------------------------------------------------------------------
    'MDB�iSAS�o�^���f�[�^�j�ւ̐ڑ�
    
      Dim cn As New ADODB.Connection
      Dim rs As New ADODB.Recordset
      Dim ConString As String
      
      ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & WorkBookPath(mdbir, "", Trim(mdbFile))
      
      cn.Open ConnectionString:=ConString
      
     '���R�[�h�Z�b�g�쐬
     rs.Open Source:=strsql, ActiveConnection:=cn, CursorType:=adOpenKeyset, LockType:=adLockOptimistic
     
'    Debug.Print rs.RecordCount & "," & rs.Fields.Count
      
    '------------------------------------------------------------------------------------------------
        
    '���o�������ʁA���R�[�h�J�E���g0�̏ꍇ�́A���������Ȃ��B
    If rs.RecordCount = 0 Then
     
'        �f�[�^��ǂݍ��܂ȂȂ����̏�����
        GoTo Next_Step
     
    End If

    '���R�[�h���ƃf�[�^���ڂ̐�����A�z����Đݒ�
     ReDim rsAry(rs.RecordCount - 1, rs.Fields.Count - 1)


    '�J�E���^�[�ϐ�
    Dim r As Long
    Dim i As Long

    '�z��Ƀf�[�^������
    Do Until rs.EOF
    
        For i = 0 To rs.Fields.Count - 1
        
            rsAry(r, i) = rs.Fields(i).Value
        
        Next i
        
        i = 0
        r = r + 1
        rs.MoveNext
    Loop
    
'    Debug.Print
    

Next_Step:
 
 InputArrayMDB = strsql
    
    mdbrec = rs.RecordCount
    
    'LOTNUM�擾
    LOTNUM = dataAry(o, UBound(dataAry, 2))
    
    rs.Close
    Set rs = Nothing
    
    cn.Close
    Set cn = Nothing
    
    DoEvents


End Function


'Public Function PosToFileWrd(Filedir As String, AcptFileDir As String, TmpFile As String, DATA() As String, Item As Long, flg As Boolean) As Boolean

Public Function PosToFileMult(Filedir As String, tmpfile As String, mdbDATA() As String, data() As String, o As Long, SheetName As String) As Boolean

    '2020.5.19 �쐬
    
    '�z��f�[�^��]�L����
    '�w��̃t�@�C�����ŕۑ�
    
    '
    '�����@�F�쐬����t�@�C���̕ۑ���
    '�@�@�@�@�e���v���[�g�t�@�C��
    '�@�@�@�@�]�L�z��f�[�^
    '�@�@�@�@�V�[�g��
    '
    '�߂�l�F��ƌ���
    
    Dim r As Long
    Dim i As Long
    
    
    ' �Ō�̗v�f�ɓ��͂��ꂽ�l���ANull�������́A"-"�̏ꍇ�́A
    '�������΂��
    
    If Len(data(o, UBound(data, 2) + -1)) <= 1 Then
     '�t�@�C�������݂��Ȃ��ꍇ�́A���̏�����
     GoTo Next_Step
        
    End If
    
    Dim TargetBook As Workbook
    
    Set TargetBook = Workbooks.Open(WorkBookPath(ThisWorkbook.Path, "", tmpfile), False, 0)
    
          
        '�t�@�C�������̍�ƂɑJ��
         With TargetBook.Sheets(SheetName)
         
              For r = 0 To UBound(mdbDATA, 1)
      
                For i = 0 To UBound(mdbDATA, 2)
                .Cells(r + 3, i + 1).Value = mdbDATA(r, i)
                Next i
                
              Next r
              
              'LOT�ԍ����t�b�^�[�ɓ���
                .PageSetup.RightFooter = data(o, UBound(data, 2))
                

        
            With .Range("A3").CurrentRegion
                .Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeRight).LineStyle = xlContinuous
                .Borders(xlInsideVertical).LineStyle = xlContinuous
                .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            End With
                  
         End With
         
        'XLSX�ŗǂ��̂����́A�Œ�
'         TargetBook.SaveAs _
'                    filename:=WorkBookPath(ThisWorkbook.Path, Filedir, data(o, UBound(data, 2) + -1) & ".xlsx"), _
'                    FileFormat:=51
'         TargetBook.Close


        With TargetBook
        
        If InStr(data(o, UBound(data, 2) + CInt(True)), ".xlsx") > 0 Then


            strFileName = WorkBookPath(ThisWorkbook.Path, Filedir, data(o, UBound(data, 2) + CInt(True)))
         
         Else
       
            strFileName = WorkBookPath(ThisWorkbook.Path, Filedir, data(o, UBound(data, 2) + CInt(True)) & ".xlsx")
       
         End If
         

             .SaveAs _
                    filename:=strFileName, _
                    FileFormat:=51
            .Close
            
         
         End With

         
         DoEvents
         
'�t�@�C�������݂��Ȃ��ꍇ�́A���̏�����
Next_Step:

End Function
