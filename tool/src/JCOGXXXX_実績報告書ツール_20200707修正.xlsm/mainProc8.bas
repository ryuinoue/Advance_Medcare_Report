Attribute VB_Name = "mainProc8"
Option Explicit


Public Sub mainProc8()
  
    '2020.04.15 �쐬
    '���ѕ񍐏��쐬Main�֐�
    
    '�@�\
    '���ѕ񍐏��쐬��ʂ��A�ݒ�擾
    '���ѕ񍐏��쐬�֐����s
    
    
'    ��Ɨp�ϐ�
    Dim adressAry1() As String
    Dim adressAry2() As String
    Dim adressAry3() As String
    Dim WordTemp() As String
    Dim WordReplace() As String
    Dim WordMapping() As String
    
'    �J�E���^�[�ϐ�
    Dim l As Long
    Dim m As Long
    Dim j As Long
    Dim i As Long
    Dim k As Long
    
'    word�e���v���[�g�i�[�p
    Dim w1() As String
    Dim w2() As String
    Dim w3() As String

    Dim TempWord() As String

    
     With ThisWorkbook.ActiveSheet
      

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
      
      
        'word���[���e���v���[�g�t�@�C���ۑ���
        Dim wrdTmp As String
        wrdTmp = .Range("C8").Value
      
      
        '�i�����X�g�ǂݍ��݊J�n�s
        Dim ListStart As Long
        ListStart = 2
      
        '�i�����X�g�ǂݍ��݃t���O��
        Dim ReadRecord As String
        ReadRecord = .Range("C9").Value
      
      
        Dim AttFolder As String
        AttFolder = .Range("F11").Value
      
    
    End With
    
    
'    '--- Outlook����̃I�u�W�F�N�g ---'
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")

    '--- Outlook����̃I�u�W�F�N�g�A�������Ƀt�H���_�[�쐬 ---'

    Dim myNameSpace As Object
    Dim DraftBox As Object
    Dim oNewFolder As Object

    Set myNameSpace = objOutlook.GetNamespace("mapi")
    Set DraftBox = myNameSpace.GetDefaultFolder(16)
    Set oNewFolder = DraftBox.Folders.Add(ShintyokuSheet & Format(Now(), "YYYYMMDD") & Format(Now(), "HHMM"))
'
'    '--- Outlook����̃I�u�W�F�N�g�A�������Ƀt�H���_�[�쐬 ---'
    
    
    
    '�i�����X�g�ݒ�ǂݍ��݊J�n�s
    Dim rec As Long
    rec = 20


    'TO�i�[�p�̔z��ϐ�
    adressAry1() = getInitArray(rec, 15, ActiveSheet.Name)
    Dim ToArry() As String
    ToArry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry1(), ListStart, "", "", "", ReadRecord)
    
    
    'CC�i�[�p�̔z��ϐ�
    adressAry1() = getInitArray(rec, 18, ActiveSheet.Name)
    Dim CcArry() As String
    CcArry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry1(), ListStart, "", "", "", ReadRecord)
    
    
    '�Y�t�t�@�C���i�[�p�̔z��ϐ�����
    adressAry1() = getInitArray(rec, 20, ActiveSheet.Name)
    Dim AttachArryLbl() As String
    AttachArryLbl = adressAry1
    
    '�Y�t�t�@�C���i�[�p�̔z��ϐ�
    adressAry1() = getInitArray(rec, 21, ActiveSheet.Name)
    Dim AttachArry() As String
    AttachArry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry1(), ListStart, "", "", "", ReadRecord)


    '�Y�t�t�@�C���i�[��t�H���_�[�̔z��ϐ�
    Dim adressAry6(0) As String
    adressAry6(0) = AttFolder
    Dim AttFolderAry() As String
    AttFolderAry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry6(), ListStart, "", "", "", ReadRecord)

    'WORD�e���v���[�g�t�@�C��,�i�[�p�̔z��ϐ�
    adressAry1() = getInitArray(rec, 2, ActiveSheet.Name)
    adressAry2() = getInitArray(rec, 3, ActiveSheet.Name)
    adressAry3() = getInitArray(rec, 4, ActiveSheet.Name)
    
    '1�����z�񂩂畡�������z���
    WordTemp() = arry2(adressAry1, adressAry2)
    WordTemp() = arry2(WordTemp, adressAry3)

  
    w1() = WordTempMail(wrdTmp, WordTemp(0, 1))
    w2() = WordTempMail(wrdTmp, WordTemp(1, 1))
    w3() = WordTempMail(wrdTmp, WordTemp(2, 1))


    'WORD�e���v���[�g�t�@�C������`�����u���@WORD�{�����̒u������,�i�[�p�̔z��ϐ�
    adressAry1() = getInitArray(rec, 6, ActiveSheet.Name)
    adressAry2() = getInitArray(rec, 7, ActiveSheet.Name)
    
    '1�����z�񂩂畡�������z���
    WordReplace() = arry2(adressAry1, adressAry2)


    'Word��������KeyWord
    adressAry1() = getInitArray(rec, 10, ActiveSheet.Name)
    '�i�����X�g��ԍ�
    adressAry2() = getInitArray(rec, 11, ActiveSheet.Name)
    

    '�i�����X�g����擾�����A�]�L�p�f�[�^�i�[�p�̔z��ϐ�
    Dim dataAry() As String

    '�i�����X�g����ݒ�l�̓ǂݍ��ݎ��s
    dataAry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry2(), ListStart, "", "", "", ReadRecord)

    'dataAry()��WordTemp���̌����u���p�L�[���[�h�i�ϐ����FadressAry1�j��1��ǉ����A3�����z���
    Dim mailBody() As String
    
    Dim kw As Long
    kw = 1
    
    '��`�����u���@WORD�{�����̒u�������ǉ�
    '�z��̃J�E���g�́A0�n�܂�Ȃ̂ŁA����2�Ȃ̂�+1
     
    ReDim mailBody(UBound(dataAry, 1), UBound(adressAry1, 1) + UBound(WordReplace, 1) + 1, kw)

        
    For l = 0 To UBound(dataAry, 1)
        
        j = 0
        
        For m = 0 To UBound(dataAry, 2)
        
        
            mailBody(i, j, 0) = adressAry1(j)
        
            mailBody(i, j, 1) = dataAry(i, j)
        
            
            j = j + 1
        
        Next m


        For m = 0 To UBound(WordReplace, 2)
        
            mailBody(i, j, 0) = WordReplace(m, 0)
        
            mailBody(i, j, 1) = WordReplace(m, 1)
            
            j = j + 1
            
        Next m


        i = i + 1

        
    Next l
    
    
'outPutText Now(), Hour(Now) & Minute(Now)

    Dim workBody() As String

    
    For i = 0 To UBound(mailBody, 1)
    
   'Debug.Print work(i, 0, 1)
    
      For j = 0 To UBound(WordTemp, 1)
       
            If mailBody(i, 0, 1) = WordTemp(j, 0) Then
            
'                Debug.Print WordTemp(j, 2)
                  
                Select Case WordTemp(j, 2)
                  Case "w1": workBody = w1
                  Case "w2": workBody = w2
                  Case "w3": workBody = w3
                End Select
              
                   
                k = 0
                Do Until Len(workBody(k)) = 0
                
                     For l = 1 To UBound(mailBody, 2)
    
                       workBody(k) = Replace(workBody(k), mailBody(i, l, 0), mailBody(i, l, 1))
                       
                     Next l
                     
                     outPutText workBody(k), Hour(Now) & Minute(Now)
    
                    k = k + 1
                
                Loop
                
              
            End If
       
      Next j
      
'
'          Dim aryTo() As String
'          aryTo = aryTrance(ToArry, i)
'
'          Dim aryCc() As String
'          aryCc = aryTrance(CcArry, i)
'
'          Dim aryAtach() As String
'          aryAtach = aryTrance(AttachArry, i)

       
      '�����̍s���T������
      Dim s As Long
      For s = 0 To UBound(workBody, 1)
       If InStr(workBody(s), "����") > 0 Then
        Exit For
       End If
      Next s
          
      SendMail oNewFolder, aryTrance(ToArry, i), aryTrance(CcArry, i), workBody(s), workBody, s, AttFolderAry(i, 0), aryTrance(AttachArry, i), AttachArryLbl()
'     SendMail aryTrance(ToArry, i), aryTrance(CcArry, i), workBody(s), workBody, s, AttFolderAry(i, 0), aryTrance(AttachArry, i)
      
      Erase workBody
      
    Next i
    
    
    
    
    
    
    
    
'   For l = 0 To UBound(WordTrance, 1)
'
'        For m = 0 To UBound(WordTrance, 2)
'
'            work(i, j, 0) = WordTrance(m, 0)
'
'            work(i, j, 1) = WordTrance(m, 1)
'
'
'        Next m
'        i = i + 1
'   Next l

'    Dim WordTrance2() As String
'    WordTrance2() = arry2(WordTrance, adressAry1_3)
'
'    Dim WordTrance3() As String
'    WordTrance3() = arry2(adressAry1_3, WordTrance)

'    '���ѕ񍐏��̓]�L��̃Z���Ԓn�i�[�p
'    Dim MapingAry() As String
'    '���ѕ񍐂̃Z���Ԓn�ݒ� �u�Ԓn�v�A�s��ԍ��A��ʃV�[�g�����w��
'    MapingAry() = getInitArray(rec, 7, ActiveSheet.Name)
'
'

'
'
'    '���{�񍐏��ւ̓]�L����
''    Call PosToFile(OutPutFolder, "", Jisitemplate, dataAry(), -1, MapingAry(), "���ѕ�", True)
'
'    Call PosToFile(OutPutFolder, "", Jisitemplate, dataAry(), -1, MapingAry(), templateSheet, True, fileExt, fileFmt)

MsgBox "�������I�����܂���"

End Sub

Public Function arry2(ary1() As String, ary2() As String) As String()

   '1�����z��

    Dim i As Long
    Dim j As Long
    Dim l As Long
    
    Dim k As Long
    Dim k2 As Long
    Dim work() As String
    
    
    k = AD(ary1())
    k2 = AD(ary2())
    
    ReDim work(UBound(ary1, 1), k + k2 - 1)
    
    For i = 0 To UBound(ary1, 1)
    
        j = 0
        For l = 0 To k - 1
                 
            If k = 1 Then
                 
                 work(i, j) = ary1(i)
        
            Else
            
                 work(i, j) = ary1(i, j)
        
            End If
            
            j = j + 1
            
        Next l

        For l = 0 To k2 - 1

            If k2 = 1 Then
                
                work(i, j) = ary2(i)
                
            Else
            
                work(i, j) = ary2(i, l)
                
            End If
            
            j = j + 1
        
        Next l

    Next i
    
    arry2 = work
    
End Function

Public Function AD(ary() As String) As Long


    Dim TempData As Variant
    Dim i As Long
    i = 0
    On Error Resume Next
    Do While Err.Number = 0
          i = i + 1
        TempData = UBound(ary, i)

    Loop
    On Error GoTo 0
    
    AD = i - 1

End Function
