Attribute VB_Name = "mainProc2"
Public Sub mainProc2()

    '2020.04.24 �쐬
    '���ѕ񍐏��X�VMain�֐�
    
    '�@�\
    '���ѕ񍐏��X�V��ʂ��A�ݒ�擾
    '���ѕ񍐏��쐬�֐����s


    With ThisWorkbook.Sheets("���ѕ񍐏��X�V")
    
    
        '���ѕ񍐏���̐�t�H���_�[�m�F
        Dim AcptFolder As String
        AcptFolder = .Range("C2").Value
      
        If FindDirectory(WorkBookPath(ThisWorkbook.Path, AcptFolder, "")) = False Then
        
            If MsgBox("��̐�̃t�H���_�[�����݂��Ă��܂���", vbOKOnly) = vbOK Then
              
              Exit Sub
            
            End If
        
        End If
    
    
        Dim fileExt As String
        fileExt = .Range("G2").Value
        
        Dim fileFmt As Long
        fileFmt = .Range("H2").Value
    
    
        '���ѕ񍐏�(�X�V)�o�͐�t�H���_�[�쐬
        Dim OutPutFolder As String
        OutPutFolder = .Range("C11").Value
      
        If FindDirectory(WorkBookPath(ThisWorkbook.Path, OutPutFolder, "")) = True Then
            
            If MsgBox("�����̃t�H���_�[�����݂��܂��B�t�H���_�[�̒����m�F���Ă�������", vbOKOnly) = vbOK Then

'                Exit Sub


            End If
'
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
    adressAry() = getInitArray(rec, 4, "���ѕ񍐏��X�V")

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
    MapingAry() = getInitArray(rec, 7, "���ѕ񍐏��X�V")


    Dim o As Long
    For o = 0 To UBound(dataAry4)

        'Call PosToFile(OutPutFolder, Jisitemplate, dataAry(), MapingAry(), "���ѕ�", "�쐬")
        Call PosToFile(OutPutFolder, AcptFolder, dataAry4(o), dataAry5(), o, MapingAry(), "���ѕ�", False, fileExt, fileFmt)
 
    Next o

'Debug.Print '

MsgBox "�������I�����܂���"

End Sub


