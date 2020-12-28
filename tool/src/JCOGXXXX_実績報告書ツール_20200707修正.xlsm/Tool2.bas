Attribute VB_Name = "Tool2"

Public Function getInitArray(rec As Long, Col As Long, SheetName As String) As String()

     '2020.4.16 �쐬
     '���ѕ񍐏��]�L�c�[��������ѕ񍐏��i�]�L��j�̔Ԓn�̐����m�F
     '�w��̍s�̃f�[�^��1�s���z��ɓ����
     
     '����    :���ѕ񍐏��]�L�c�[���̃V�[�g��
     '�@�@�@�@ �ǂݍ��݊J�n�ʒu�i�s�j
     '�@�@     �ǂݍ��݊J�n�ʒu�i��j
    
     '�߂�l�@:�]�L��̃Z���Ԓn
     
     '��Ɨp�̓]�L�p�f�[�^�i�[�p�z��ϐ�
     Dim work() As String
     '�i�����X�g����ǂݍ��ރ��R�[�h��
     Dim l As Long
     '�J�E���^�[�ϐ�
     Dim r As Long
    
    '
     With ThisWorkbook.Sheets(SheetName)
     
        .Select
        .Cells(1048576, Col).Activate
        
        '�i�����X�g����ǂݍ��ރ��R�[�h���m�F
        Selection.End(xlUp).Select
        l = Selection.Row
        
        r = rec
        
        '�J�E���^�[�ϐ�
        Dim i As Long
        i = 0
        
        '���R�[�h������A�z����Đݒ�
        ReDim work(l - r - 1)
        
        Do Until l <= r
         
             work(i) = .Cells(r + 1, Col).Value
             r = r + 1
             i = i + 1
         
        Loop
     
     End With
     
     getInitArray = work
     
End Function


Public Function InputArray(filename As String, SheetName As String, Mapping() As String, rec As Long, LotCol As String, LotNames As String, LotLastName As String, ReadingFlg As String) As String()

    '2020.4.15 �쐬
    '2020.4.24 �X�V
    
    '�i�����X�g���J���i�����N�X�V�̃��b�Z�[�W��\�������Ȃ��j
    '�i�����X�g����ǂݍ��ރ��R�[�h�����m�F����
    '�w��̗�̃f�[�^��1�s���z��ɓ����
    
    '
    '�����@�F�i�����X�g�̃t�@�C����
    '        �i�����X�g�̃V�[�g��
    '�@�@�@�@�i�����X�g���甲�����f�[�^���ځi�G�N�Z��:��j
    '�@�@�@  �i�����X�g�̓ǂݍ��݊J�n�ʒu
    '�@�@�@�@�i�����X�g�ɒǋL����Lot�ԍ��X�V��
    '�@�@�@�@�i�����X�g�ɒǋL����Lot�ԍ��̐ړ���
    '�@�@�@�@�i�����X�g�ɒǋL����Lot�ԍ��̐ڔ���
    '        �i�����X�g����ǂݍ��ރ��R�[�h�𔻒肷��t���O��
    '
    '
    
    '�߂�l�F�i�����X�g����]�L�p�̔z��f�[�^
    
    
    '��Ɨp�̓]�L�p�f�[�^�i�[�p�z��ϐ�
    Dim work() As String
    Dim Work2() As String
    
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
        
        
        '���R�[�h���ƃf�[�^���ځi�����j�̐�����A�z����Đݒ�
        '���ѕ񍐏��X�V�̏ꍇ�́A�v�f�́A���̂܂܁B
        '���ѕ񍐏��쐬�̏ꍇ�́A�v�f��1�ǉ����܂��B
        '��5�����A��6�������󗓂ł͖����ꍇ�ALot�ԍ��p�̗v�f��ǉ�����B
        
        If Len(LotCol) = 0 And Len(LotNames) = 0 Then
            ReDim work(l - r, UBound(Mapping))
        Else
            ReDim work(l - r, UBound(Mapping) + 1)
        End If
          
        '�J�E���^�[�ϐ�
        Dim i As Long
        
        '�J�E���^�[�ϐ��A�ǂݍ��܂Ȃ����R�[�h��
        Dim j As Long
        j = 0
        
        '���[�v�I�������F�������郌�R�[�h���A�i�����X�g�̃��R�[�h���𒴂�����I��
        Do Until l < r
           
            i = 0
            
            
            '����[�i�����X�g����ǂݍ��ރ��R�[�h�𔻒肷��t���O��]�ɒl������ꍇ�A
            '�t���O�Ɏw�肳�ꂽ���ڂ̒l�ɂ���āA�f�[�^�̓ǂݍ��ނ��ۂ��̔�����s��
            If Len(ReadingFlg) > 0 Then
            
                '�w�肳�ꂽ��̃f�[�^�ɒl���݂�΁A�ǂݍ��݁A
                '�Ȃ���΁A�ǂݍ��܂Ȃ��B
                
                If Len(.Range(ReadingFlg & r).Value) = 0 Then
                
                    '�ǂݍ��܂Ȃ����R�[�h��
                    j = j + 1
                    '�f�[�^��ǂݍ��܂ȂȂ����̏�����
                    GoTo Next_Step
                
                End If
            
            End If
            
            '�Z���Ԓn�p�̕ϐ�
            Dim adress As Variant
            
            For Each adress In Mapping
               
               '�f�[�^���ځi�����j����Z���Ԓn�𐶐�
               '�]�L�p�f�[�^��z��Ɋi�[
               work(r - rec, i) = .Range(adress & r).Value
            
               i = i + 1
            
            Next adress
            
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

                work(r - rec, i) = LotNames & "_" & .Range(LotLastName & r).Value
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
    
    
    '�z��Work����s�v�ȋ󔒍s���폜���邽�߂̏���
    ReDim Work2(UBound(work, 1) - j, UBound(work, 2))
    
    Dim k As Long
    Dim m As Long
    m = 0
    Dim n As Long
    n = 0
    
    Dim RecThum As Long
    Dim o  As Long
    
    For k = 0 To UBound(work, 1)
        
'       If UBound(work, 2) > 0 Then
''      �z��̍ŏ��̐�����Null�̏ꍇ�A�폜�Ƃ����v������
''�@�@�@�z��̑S�Ă̒l��Null�̏ꍇ�A�폜�ɗv����ύX
'        For o = 0 To UBound(work, 2)
'            RecThum = RecThum + Len(work(k, o))
'        Next o
'
'       Else
       
'        RecThum = Len(work(k, 0))

'       End If
       
'        If RecThum <> 0 Then
        
         If Len(work(k, 0)) <> 0 Then
         
            For m = 0 To UBound(work, 2)
                Work2(n, m) = work(k, m)
            Next m
            
            n = n + 1
            
        End If
    Next k
    
    '��Ɨp�̔z��f�[�^���Ăяo�����ɕԂ��B
    InputArray = Work2

End Function



Public Function LotFileAry(FOLDERNAME As String) As String()

    '2020.4.22 �쐬
    '��̐�t�H���_���̃t�@�C����LOT�ԍ��̔z����쐬
    '�����@�F��̐�t�H���_��
    '�߂�l�F��̐�t�H���_���̃t�@�C����LOT�ԍ��̔z��


    'Debug.Print WorkBookPath(ThisWorkbook.Path, FolderName, "")
    
    Dim buf As String
    
'    buf = dir(WorkBookPath(ThisWorkbook.Path, FolderName, "") & "\*.xls")
'    buf = dir(Trim(WorkBookPath(ThisWorkbook.Path, FolderName, "") & "\*" & Ext))
    
    buf = Dir(WorkBookPath(ThisWorkbook.Path, FOLDERNAME, "\*" & Ext))
  
    
    Dim work() As String


    '�t�H���_�[���̃t�@�C�������J�E���g
    Dim cnt As Long
    cnt = 0
    
    Do While buf <> ""
        buf = Dir()
        cnt = cnt + 1
    Loop
    
    
    ReDim work(cnt - 1, 1)
    
    
'    buf = dir(WorkBookPath(ThisWorkbook.Path, FolderName, "") & "\*.xls")
    
    buf = Dir(WorkBookPath(ThisWorkbook.Path, FOLDERNAME, "\*" & Ext))
    
    
    Dim filename As String
    Dim i As Long
    i = 0
    
    
    '
    Do While buf <> ""
        
        'Debug.Print getLotNum(WorkBookPath(ThisWorkbook.Path, FolderName, buf)) & ":"; buf
        
        work(i, 0) = getLotNum(WorkBookPath(ThisWorkbook.Path, FOLDERNAME, buf))
        work(i, 1) = buf
        buf = Dir()
        i = i + 1
        
    Loop
    
    LotFileAry = work

End Function
