Attribute VB_Name = "Tool4"
Public Sub TEST_SendMail2()
    
    '--- Outlook����̃I�u�W�F�N�g ---'
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
'    Set NewItem = objOutlook.CreateItem(0)
    
Set myNameSpace = objOutlook.GetNamespace("mapi")
Set DraftBox = myNameSpace.GetDefaultFolder(16)



Set oNewFolder = DraftBox.Folders.Add("�W�v")   '��M�g���C�̒����Ɂu�W�v�v�t�H���_���쐬���܂�
'oFolder.Folders.Add (�g�W�v�h) �f������ł��쐬�ł��܂�
'
'Set oNewFolder = Nothing
'Set oFolder = Nothing
'Set myNameSpace = Nothing


    
'    NewItem.Add ("JCOG***")
    
    '--- ���[���I�u�W�F�N�g ---'
    Dim objMail As Object
    Set objMail = objOutlook.CreateItem(olMailItem)
        
    '--- ���[���̓��e���i�[����ϐ� ---'
    Dim ToStr As String
    Dim CcStr As String
    Dim BccStr As String
    Dim subjectStr As String
    Dim BodyStr As String
    
    '--- ����̓��e ---'
    ToStr = "ryuinoue@ncc.go.jp;ryuinoue@ncc.go.jp"
    CcStr = "ryuinoue@ncc.go.jp"
    BccStr = "ryuinoue@ncc.go.jp"
'    toStr = "ryuinoue2@ncc.go.jp"
    
    '--- �����̓��e ---'
    subjectStr = "TEST"
    
    '--- �{���̓��e ---'
    BodyStr = "TEST"
        
    '--- ������ݒ� ---'
    objMail.To = ToStr
    objMail.To = ToStr
    objMail.CC = CcStr
    objMail.BCC = BccStr
    objMail.subject = subjectStr
    objMail.BodyFormat = olFormatPlain
    objMail.Body = BodyStr
    
    '--- �Y�t�t�@�C���̃p�X ---'
    Dim attachmentPath As String
    attachmentPath = "\\Mac\Home\Desktop\JCOG�ݑ���\20200622_���\2020\JCOG1502C\1102_��ʌ�������Z���^�[\�쐬\0219_53B_��ʌ�������Z���^�[_�l����1��(�ʓY1).xls"
    
    '--- �Y�t�t�@�C����ݒ� ---'
    Call objMail.Attachments.Add(attachmentPath)
    
    attachmentPath = "\\Mac\Home\Desktop\JCOG�ݑ���\20200622_���\2020\JCOG1502C\0901_�Ȗ،�������Z���^�[\�쐬\JCOG1502_53B_�n���Ɨ��s���@�l�Ȗ،�������Z���^�[_�l����1��.doc"
         Call objMail.Attachments.Add(attachmentPath)
    attachmentPath = "\\Mac\Home\Desktop\JCOG�ݑ���\20200622_���\2020\JCOG1502C\1102_��ʌ�������Z���^�[\�쐬\0219_53B_��ʌ�������Z���^�[_�l����1��(�ʓY4).xlsx"
         Call objMail.Attachments.Add(attachmentPath)
    attachmentPath = "\\Mac\Home\Desktop\JCOG�ݑ���\20200622_���\2020\JCOG1502C\1102_��ʌ�������Z���^�[\�쐬\�l����1��(�ʓY1)_�L�ڎ����_���тȂ��̏ꍇ.pdf"
    
     Call objMail.Attachments.Add(attachmentPath)
    
    '--- ���[����\�� ---'
    objMail.Save
'    objMail.Display
    
    '--- ���[���𑗕t ---'
'    objMail.Send

objMail.Move oNewFolder
        
End Sub


'Public Sub SendMail(ToStr() As String, CcStr() As String, subjectStr As String, BodyStr() As String, s As Long, attchDir As String, attachmentPath() As String)
Public Sub SendMail(oNewFolder As Object, ToStr() As String, CcStr() As String, subjectStr As String, BodyStr() As String, s As Long, attchDir As String, attachmentPath() As String, attachmentLbl() As String)

    Dim i As Long

    '--- Outlook����̃I�u�W�F�N�g ---'
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")

'    '--- Outlook����̃I�u�W�F�N�g�A�������Ƀt�H���_�[�쐬 ---'
'
'    Set myNameSpace = objOutlook.GetNamespace("mapi")
'    Set DraftBox = myNameSpace.GetDefaultFolder(16)
'    Set oNewFolder = DraftBox.Folders.Add(shiken & Format(Now(), "YYYYMMDD"))


    '--- ���[���I�u�W�F�N�g ---'
    Dim objMail As Object
    Set objMail = objOutlook.CreateItem(olMailItem)
        

    '--- ������ݒ� ---'
     'To
     i = 0
     objMail.To = ToStr(i) & ";"
     For i = 1 To UBound(ToStr, 1)
        objMail.To = objMail.To & ";" & ToStr(i)
     Next i
     
     'Cc
     i = 0
     objMail.CC = CcStr(i) & ";"
     For i = 1 To UBound(CcStr, 1)
        objMail.CC = objMail.CC & ";" & CcStr(i)
     Next i
     
     '����
     objMail.subject = Trim(Replace(subjectStr, "�����F", ""))
     objMail.BodyFormat = olFormatPlain
     
     '�{��
     s = s + 2
     objMail.Body = BodyStr(s)
     For i = s + 1 To UBound(BodyStr, 1)
         objMail.Body = objMail.Body & vbCrLf & BodyStr(i)
     Next i

     '�Y�t�t�@�C��
     For i = 0 To UBound(attachmentPath, 1)
     
          If FindFile(WorkBookPath(attchDir, "", attachmentPath(i))) And Len(attachmentPath(i)) > 0 Then
          
'          Debug.Print FindFile(WorkBookPath(attchDir, "", attachmentPath(i)))
            
            Call objMail.Attachments.Add(WorkBookPath(attchDir, "", attachmentPath(i)))

          Else
            
'            outPutText Now, "missFile"
            outPutText Now & "  " & attachmentLbl(i) & " " & WorkBookPath(attchDir, "", attachmentPath(i)), "missFile"
            
          End If

     Next i
     
     
     '--- ���[����ۑ��A�t�H���_�[�Ɉړ� ---'
     objMail.Save
'    objMail.Display

    '--- ���[���𑗕t ---'
'''    objMail.Send
     objMail.Move oNewFolder
     
     Set objMail = Nothing
     Set objMail = Nothing

End Sub

Public Function aryTrance(Arr() As String, i As Long) As String()

    Dim work() As String
    ReDim work(UBound(Arr, 2))
    
    Dim j As Long
    
    ReDim work(UBound(Arr, 2))
    For j = 0 To UBound(Arr, 2)
    
        work(j) = Arr(i, j)
        
    Next j

    aryTrance = work

End Function

