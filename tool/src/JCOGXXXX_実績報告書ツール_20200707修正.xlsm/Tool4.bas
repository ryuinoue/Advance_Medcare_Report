Attribute VB_Name = "Tool4"
Public Sub TEST_SendMail2()
    
    '--- Outlook操作のオブジェクト ---'
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
'    Set NewItem = objOutlook.CreateItem(0)
    
Set myNameSpace = objOutlook.GetNamespace("mapi")
Set DraftBox = myNameSpace.GetDefaultFolder(16)



Set oNewFolder = DraftBox.Folders.Add("集計")   '受信トレイの直下に「集計」フォルダを作成します
'oFolder.Folders.Add (“集計”) ’こちらでも作成できます
'
'Set oNewFolder = Nothing
'Set oFolder = Nothing
'Set myNameSpace = Nothing


    
'    NewItem.Add ("JCOG***")
    
    '--- メールオブジェクト ---'
    Dim objMail As Object
    Set objMail = objOutlook.CreateItem(olMailItem)
        
    '--- メールの内容を格納する変数 ---'
    Dim ToStr As String
    Dim CcStr As String
    Dim BccStr As String
    Dim subjectStr As String
    Dim BodyStr As String
    
    '--- 宛先の内容 ---'
    ToStr = "ryuinoue@ncc.go.jp;ryuinoue@ncc.go.jp"
    CcStr = "ryuinoue@ncc.go.jp"
    BccStr = "ryuinoue@ncc.go.jp"
'    toStr = "ryuinoue2@ncc.go.jp"
    
    '--- 件名の内容 ---'
    subjectStr = "TEST"
    
    '--- 本文の内容 ---'
    BodyStr = "TEST"
        
    '--- 条件を設定 ---'
    objMail.To = ToStr
    objMail.To = ToStr
    objMail.CC = CcStr
    objMail.BCC = BccStr
    objMail.subject = subjectStr
    objMail.BodyFormat = olFormatPlain
    objMail.Body = BodyStr
    
    '--- 添付ファイルのパス ---'
    Dim attachmentPath As String
    attachmentPath = "\\Mac\Home\Desktop\JCOG在宅作業\20200622_作業\2020\JCOG1502C\1102_埼玉県立がんセンター\作成\0219_53B_埼玉県立がんセンター_様式第1号(別添1).xls"
    
    '--- 添付ファイルを設定 ---'
    Call objMail.Attachments.Add(attachmentPath)
    
    attachmentPath = "\\Mac\Home\Desktop\JCOG在宅作業\20200622_作業\2020\JCOG1502C\0901_栃木県立がんセンター\作成\JCOG1502_53B_地方独立行政法人栃木県立がんセンター_様式第1号.doc"
         Call objMail.Attachments.Add(attachmentPath)
    attachmentPath = "\\Mac\Home\Desktop\JCOG在宅作業\20200622_作業\2020\JCOG1502C\1102_埼玉県立がんセンター\作成\0219_53B_埼玉県立がんセンター_様式第1号(別添4).xlsx"
         Call objMail.Attachments.Add(attachmentPath)
    attachmentPath = "\\Mac\Home\Desktop\JCOG在宅作業\20200622_作業\2020\JCOG1502C\1102_埼玉県立がんセンター\作成\様式第1号(別添1)_記載手引き_実績なしの場合.pdf"
    
     Call objMail.Attachments.Add(attachmentPath)
    
    '--- メールを表示 ---'
    objMail.Save
'    objMail.Display
    
    '--- メールを送付 ---'
'    objMail.Send

objMail.Move oNewFolder
        
End Sub


'Public Sub SendMail(ToStr() As String, CcStr() As String, subjectStr As String, BodyStr() As String, s As Long, attchDir As String, attachmentPath() As String)
Public Sub SendMail(oNewFolder As Object, ToStr() As String, CcStr() As String, subjectStr As String, BodyStr() As String, s As Long, attchDir As String, attachmentPath() As String, attachmentLbl() As String)

    Dim i As Long

    '--- Outlook操作のオブジェクト ---'
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")

'    '--- Outlook操作のオブジェクト、下書きにフォルダー作成 ---'
'
'    Set myNameSpace = objOutlook.GetNamespace("mapi")
'    Set DraftBox = myNameSpace.GetDefaultFolder(16)
'    Set oNewFolder = DraftBox.Folders.Add(shiken & Format(Now(), "YYYYMMDD"))


    '--- メールオブジェクト ---'
    Dim objMail As Object
    Set objMail = objOutlook.CreateItem(olMailItem)
        

    '--- 条件を設定 ---'
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
     
     '件名
     objMail.subject = Trim(Replace(subjectStr, "件名：", ""))
     objMail.BodyFormat = olFormatPlain
     
     '本文
     s = s + 2
     objMail.Body = BodyStr(s)
     For i = s + 1 To UBound(BodyStr, 1)
         objMail.Body = objMail.Body & vbCrLf & BodyStr(i)
     Next i

     '添付ファイル
     For i = 0 To UBound(attachmentPath, 1)
     
          If FindFile(WorkBookPath(attchDir, "", attachmentPath(i))) And Len(attachmentPath(i)) > 0 Then
          
'          Debug.Print FindFile(WorkBookPath(attchDir, "", attachmentPath(i)))
            
            Call objMail.Attachments.Add(WorkBookPath(attchDir, "", attachmentPath(i)))

          Else
            
'            outPutText Now, "missFile"
            outPutText Now & "  " & attachmentLbl(i) & " " & WorkBookPath(attchDir, "", attachmentPath(i)), "missFile"
            
          End If

     Next i
     
     
     '--- メールを保存、フォルダーに移動 ---'
     objMail.Save
'    objMail.Display

    '--- メールを送付 ---'
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

