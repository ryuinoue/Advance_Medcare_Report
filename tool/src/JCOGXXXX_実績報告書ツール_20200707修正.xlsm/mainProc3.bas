Attribute VB_Name = "mainProc3"

Public Sub mainProc3()

    '2020.04.24 �쐬
    '���ѕ񍐏��m�FMain�֐�
    
    '�@�\
    '���ѕ񍐏��X�V��ʂ��A�ݒ�擾
    '���ѕ񍐏��쐬�֐����s


    With ThisWorkbook.Sheets("���ѕ񍐏��m�F")
    
    
        '���ѕ񍐏���̐�t�H���_�[�m�F
        Dim AcptFolder As String
        AcptFolder = .Range("C2").Value
      
        If FindDirectory(WorkBookPath(ThisWorkbook.Path, AcptFolder, "")) = False Then
        
            If MsgBox("��̐�̃t�H���_�[�����݂��Ă��܂���", vbOKOnly) = vbOK Then
              
              Exit Sub
            
            End If
        
        End If
    
    
'���{�񍐏��ɂ͕��o�������Ȃ��d�l�ɕύX
'        '���ѕ񍐏�(�X�V)�o�͐�t�H���_�[�쐬
'        Dim OutPutFolder As String
'        OutPutFolder = .Range("C11").Value
'
'        If FindDirectory(WorkBookPath(ThisWorkbook.Path, OutPutFolder, "")) = True Then
'
'            If MsgBox("�����̃t�H���_�[�����݂��܂��B�t�H���_�[�̒����m�F���Ă�������", vbOKOnly) = vbOK Then
'
''                Exit Sub
'
'
'            End If
''
'        Else
'
'                '�t�H���_�[��������΍쐬
'                MkDir WorkBookPath(ThisWorkbook.Path, OutPutFolder, "")
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
        
        '�m�F���ʏo�͐�i�I�t�Z�b�g�Ԓn�j
        Dim ofset As Long
        ofset = .Range("C10").Value
        
    End With

    '�i�����X�g�A���ѕ񍐏��ݒ�ǂݍ��݊J�n�s
    Dim rec As Long
    rec = 15

    '��������t�@�C���̃t�@�C�����ALOTNO�̔z��ϐ�
    Dim dataAry() As String
    '���ѕ񍐏���̐�t�H���_�[�����w��
    dataAry() = LotFileAry(AcptFolder)


    '�i�����X�g�̗�����i�[�p�̔z��ϐ�
    Dim adressAry() As String
    '�i�����X�g�̍��ڐݒ� �u��v�̍s�A��ԍ��A��ʃV�[�g�����w��
    adressAry() = getInitArray(rec, 4, "���ѕ񍐏��m�F")

    '�i�����X�g����擾�����A�]�L�p�f�[�^�i�[�p�̔z��ϐ�
    Dim dataAry2() As String
    dataAry2() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry(), 2, "", "", "", ReadRecord)



    '���W�b�N�m�F
    '�]�LRAW�f�[�^ = �S��̂����t�@�C���̃f�[�^(�t�@�C�����ALOT�ԍ�)+�i�����X�g�̃f�[�^(�i�����X�g�̍���...,LOT�ԍ�) ��LOT�ԍ����L�[�Ƃ��ă}�[�W

    '���]�L���R�[�h��=��������t�@�C����
    '�]�LRAW�f�[�^�́A��������t�@�C������擾�����f�[�^�̍��ڐ�(LOT�ԍ��A�t�@�C����)�{�i�����X�g�̍��ڐ�
    '�]�LRAW�f�[�^�i�[�p�z��f�[�^(�t�@�C����,�i�����X�g�̍���1,�i�����X�g�̍���2,�c,�c)�@��LOT�ԍ��́A������g�p���Ȃ��f�[�^�Ȃ̂Ŋi�[���Ȃ��B
    Dim dataAry3() As String
    ReDim dataAry3(UBound(dataAry()), (UBound(dataAry(), 2) - 1) + UBound(adressAry(), 1))

    
    Dim i As Long
    Dim j As Long

    '�J��Ԃ��̏����̐��́A���R�[�h���́��t�@�C���̐��A�����AdataAry�̓Y����1�Ԗڂ̍ő�v�f��
    For i = 0 To UBound(dataAry(), 1)
    
     'dataAry:��������t�@�C���̔z��
     'dataAry2:�i�����X�g����擾�����f�[�^�̔z��
     '2�̔z��̒��ɂ���LOT�ԍ����������̒T��
     
     '�T���񐔂́A�i�����X�g�̃��R�[�h��
      For j = 0 To UBound(dataAry2(), 1)
        
           Dim k As Long
           k = 0
           '2�̔z��̒��ɂ���LOT�ԍ��������ꍇ�A
           If Trim(dataAry(i, 0)) = Trim(dataAry2(j, UBound(dataAry2(), 2))) Then
           
               '�]�LRAW�f�[�^�i�[�p�z��f�[�^�̓Y����1�ɁA�t�@�C���������
                dataAry3(i, k) = dataAry(i, 1)
                
                '�]�LRAW�f�[�^�i�[�p�z��f�[�^�̓Y����2�ȍ~�́A�i�����X�g�̃f�[�^
                Dim l As Long
                l = 0
                For k = 1 To UBound(dataAry3(), 2)
                
                    dataAry3(i, k) = dataAry2(j, k - 1)
                    
                Next k
           
           End If
      Next j
    Next i
    
    '����
    'dataAry3�́A�X�V�����ɕK�v�ȃf�[�^���ׂĂ��������z�񂪊���
    '���؂��K�v�ɂȂ����ꍇ�A�܂�dataAry3�̓��e���m�F����B


    '�]�LRAW�f�[�^�i�[�p�z����A�t�@�C�������ڂ𔲂��o�����z��
    Dim dataAry4() As String
    ReDim dataAry4(UBound(dataAry3(), 1))
    
    '�]�LRAW�f�[�^�i�[�p�z����A�i�����X�g�����]�L�p�̍��ڂ𔲂��o�����z��
    Dim dataAry5() As String
    ReDim dataAry5(UBound(dataAry3()), UBound(dataAry3(), 2) - 1)
    
    'Debug.Print UBound(dataAry3(), 2) - 1
    
    Dim m As Long
    Dim n As Long
    
    '�J��Ԃ������̉񐔂́A�]�LRAW�f�[�^�i�[�p�z��̃��R�[�h���A�����Y����1�Ԗڂ̍ő�v�f��
    For m = 0 To UBound(dataAry3, 1)
    
        '�t�@�C���������
        dataAry4(m) = dataAry3(m, 0)
        
        '�t�@�C�����ȊO�̐i�����X�g�̍��ڂ����
        For n = 0 To UBound(dataAry3, 2) - 1
            dataAry5(m, n) = dataAry3(m, n + 1)
        Next n
        
    Next m


    '���ѕ񍐏��̓]�L��̃Z���Ԓn�i�[�p
    Dim MapingAry() As String
    '���ѕ񍐂̃Z���Ԓn�ݒ� �u�Ԓn�v�A�s��ԍ��A��ʃV�[�g�����w��
    MapingAry() = getInitArray(rec, 7, "���ѕ񍐏��m�F")


    '���ѕ񍐏��̓]�L��̃Z���Ԓn�i�[�p
    Dim itemAry() As String
    '���ѕ񍐂̃Z���Ԓn�ݒ� �u�Ԓn�v�A�s��ԍ��A��ʃV�[�g�����w��
    itemAry() = getInitArray(rec, 6, "���ѕ񍐏��m�F")


    Dim o As Long
    
    Dim LOTNUM As String
    Dim alrtMsg As String
    
    For o = 0 To UBound(dataAry4)

        'Call PosToFile(OutPutFolder, Jisitemplate, dataAry(), MapingAry(), "���ѕ�", "�쐬")
        'Call PosToCheckFile(OutPutFolder, AcptFolder, dataAry4(o), dataAry5(), o, MapingAry(), ItemAry(), "���ѕ�", False)
        
    
        Call PosToCheckFile(AcptFolder, dataAry4(o), dataAry5(), o, MapingAry(), itemAry(), "���ѕ�", False, LOTNUM, alrtMsg)
        
        '    dataAry2() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry(), 2, "", "", ReadRecord)
        
        Call InPutQuery(ShintyokuFile, ShintyokuSheet, LOTNUM, alrtMsg, adressAry(UBound(adressAry)), ofset)
        
    Next o

'Debug.Print '

MsgBox "�������I�����܂���"

End Sub


'Public Function PosToCheckFile(Filedir As String, AcptFileDir As String, TmpFile As String, DATA() As String, Item As Long, Mapping() As String, ItemAry() As String, SheetName As String, flg As Boolean) As Boolean
    '
    '�����@�F�쐬�������ѕ񍐏��̕ۑ���(�~)
    '�w��̃t�@�C�����ŕۑ�

    '�߂�l�F��ƌ���

Public Function PosToCheckFile(AcptFileDir As String, tmpfile As String, data() As String, Item As Long, Mapping() As String, itemAry() As String, SheetName As String, flg As Boolean, LOTNUM As String, alrtMsg As String)

    '2020.4.15 �쐬
    
    '�㕔�i��{���j�̃`�F�b�N���e
    '�i�����X�g�̔z��f�[�^�Ǝw��̃Z���̓]�L��i�Z���Ԓn�j�l���r������Ă���΁A�Z���̖Ԋ|����Ԃ��B
    '�i�����X�g�̒l���R�����g�ɓ����
    

    
    '����
    '       ����
    '        1)��̂������ѕ񍐏��̕ۑ���
    '�@�@�@�@2)���ѕ񍐃e���v���[�g�t�@�C��
    
    '�@�@�@�@3)�i�����X�g�]�L�p�z��f�[�^
    
    '�@�@�@�@�u�m�F�v�̔z��\���f�t�H���g���u�쐬�v�ɏ�����
    '�@�@�@�@�@DATA(ONAME,�͏o�󗝔N����,��i��Â̔�p (�͏o��),�l����,�Z�p��,�t�@�C����,LOT�ԍ�)
    '�@�@�@�@�@�Ō�ƍŌォ��1�O�̗v�f�͌Œ�A���͐ݒ�ɂ��ύX��
    
    '        4)�i�����X�g�]�L�p�z��f�[�^�̓Y�����B
    ' �@�@�@�@�@�@-1�̏ꍇ�A�J�E���^�[�ϐ��́A�J�E���g�A�b�v�B
    '           �@-1�ł͖����ꍇ�́A�Œ�l
    
    
    '�@�@�@�@5)�]�L��i�Z���Ԓn�j
    '        6)���ѕ񍐏��]�L�c�[���̃V�[�g��
    '        7)�t���O
    '�@�@�@�@�@�@�@�u�쐬�v�̏ꍇ�ATrue(=-1)
    '              �u�X�V�v�̏ꍇ�AFalse(=0)
            

    '        �o��
    '�@�@�@�@8)LOT�ԍ� �u�����N�œn�����B�֐��̒��ŁA���ѕ񍐏��̃t�b�^�[�̒l�����͂����B�߂��̃v���V�W���[�Ŏg�p�����B
    '�@�@�@�@9)�A���[�g���b�Z�[�W�@�u�����N�œn�����B�֐��̒��ŁA�A���[�g���b�Z�[�W�����͂����B�߂��̃v���V�W���[�Ŏg�p�����B
             
    
  
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


alrtMsg = ""

    For i = 0 To RecordNo
      
        '���ѕ񍐃t�@�C���e���v���[�g���I�[�v��
        '�u�쐬�v�̏ꍇ�́A�e���v���[�g
        '�u�X�V�v�̏ꍇ�́A��̂����t�@�C��
        Set TargetBook = Workbooks.Open(WorkBookPath(ThisWorkbook.Path, AcptFileDir, tmpfile), False, 0)
    
          
        '�i�����X�g�t�@�C�������̍�ƂɑJ��
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
                
                If .Range(Mapping(j)).Value <> data(i, j) Then
                

'                    ���ѕ񍐏��ɕ��o��������ꍇ�̏����BExcel�̃o�[�W�����ɂ���Ă̓G���[�ƂȂ邽�߃R�����g�A�E�g
'                    .Range(Mapping(j)).AddCommentThreaded ("�i�����X�g�̒l:" & CStr(DATA(i, j)))

'                    Debug.Print ItemAry(j) & " �i�����X�g:" & DATA(i, j) & " <-> ���ѕ񍐏���:" & .Range(Mapping(j)).Value
                    
                    alrtMsg = alrtMsg & "�y" & itemAry(j) & "�z �i�����X�g�u" & data(i, j) & "�v  ���ѕ񍐏����i�{�ݓ��͒l�j�u" & IIf(Len(.Range(Mapping(j)).Value) = 0, "null", .Range(Mapping(j)).Value) & "�v" & vbCrLf
                
                End If
                
               
              Next j

              
'�m�F�p���W���[���Ȃ̂ŁALOT�ԍ��̏����͉������Ȃ��B
'              '��ԍŌ�̗v�f�ɓ��͂���LOT�ԍ����t�b�^�[�ɓ���
'              j = j + 1
'
'              If flg Then
'                  .PageSetup.RightFooter = DATA(i, j)
''                  Debug.Print .PageSetup.RightFooter
'              Else
'                  .PageSetup.RightFooter = Null
'              End If
              
              LOTNUM = .PageSetup.RightFooter
              
         End With
         
         'Debug.Print CInt(flg)
         
         '�i�����X�g�]�L�z��A
         '�u�쐬�v�̏ꍇ�A�Ōォ���O�̗v�f�@�S�v�f���{True(=-1)
         '�u�X�V�v�̏ꍇ�A�Ō�̗v�f
         '�@�ɓ��͂��ꂽ�l���t�@�C�����Ƃ��ĕۑ�����
         
'         TargetBook.SaveAs _
'                    FileName:=WorkBookPath(ThisWorkbook.Path, Filedir, DATA(i, UBound(DATA, 2) + CInt(flg)) & Ext), _
'                    FileFormat:=xlsFmt
         
         TargetBook.Close False
         
         DoEvents
         
     Next i
'
'  Debug.Print alrtMsg
  
End Function

