Attribute VB_Name = "mainProc8"
Option Explicit


Public Sub mainProc8()
  
    '2020.04.15 作成
    '実績報告書作成Main関数
    
    '機能
    '実績報告書作成画面より、設定取得
    '実績報告書作成関数実行
    
    
'    作業用変数
    Dim adressAry1() As String
    Dim adressAry2() As String
    Dim adressAry3() As String
    Dim WordTemp() As String
    Dim WordReplace() As String
    Dim WordMapping() As String
    
'    カウンター変数
    Dim l As Long
    Dim m As Long
    Dim j As Long
    Dim i As Long
    Dim k As Long
    
'    wordテンプレート格納用
    Dim w1() As String
    Dim w2() As String
    Dim w3() As String

    Dim TempWord() As String


     With ThisWorkbook.ActiveSheet
      

        '読み込む進捗リストファイル名
        Dim ShintyokuFile As String
        ShintyokuFile = .Range("C5").Value
      
        If FindFile(WorkBookPath(ThisWorkbook.Path, "", ShintyokuFile)) = False Then
      
            If MsgBox("進捗リストが存在していません", vbOKOnly) = vbOK Then
                
                Exit Sub
                
            End If
            
        End If
      
        '読み込む進捗リストシート名
        Dim ShintyokuSheet As String
        ShintyokuSheet = .Range("C7").Value
      
      
        'wordメールテンプレートファイル保存先
        Dim wrdTmp As String
        wrdTmp = .Range("C8").Value
      
      
        '進捗リスト読み込み開始行
        Dim ListStart As Long
        ListStart = 2
      
        '進捗リスト読み込みフラグ列
        Dim ReadRecord As String
        ReadRecord = .Range("C9").Value
      
      
        Dim AttFolder As String
        AttFolder = .Range("F11").Value
      
    
    End With
    
    
'    '--- Outlook操作のオブジェクト ---'
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")

    '--- Outlook操作のオブジェクト、下書きにフォルダー作成 ---'

    Dim myNameSpace As Object
    Dim DraftBox As Object
    Dim oNewFolder As Object

    Set myNameSpace = objOutlook.GetNamespace("mapi")
    Set DraftBox = myNameSpace.GetDefaultFolder(16)
    Set oNewFolder = DraftBox.Folders.Add(ShintyokuSheet & Format(Now(), "YYYYMMDD") & Format(Now(), "HHMM"))
'
'    '--- Outlook操作のオブジェクト、下書きにフォルダー作成 ---'
    
    
    
    '進捗リスト設定読み込み開始行
    Dim rec As Long
    rec = 20


    'TO格納用の配列変数
    adressAry1() = getInitArray(rec, 15, ActiveSheet.Name)
    Dim ToArry() As String
    ToArry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry1(), ListStart, "", "", "", ReadRecord)
    
    
    'CC格納用の配列変数
    adressAry1() = getInitArray(rec, 18, ActiveSheet.Name)
    Dim CcArry() As String
    CcArry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry1(), ListStart, "", "", "", ReadRecord)
    
    
    '添付ファイル格納用の配列変数項目
    adressAry1() = getInitArray(rec, 20, ActiveSheet.Name)
    Dim AttachArryLbl() As String
    AttachArryLbl = adressAry1
    
    '添付ファイル格納用の配列変数
    adressAry1() = getInitArray(rec, 21, ActiveSheet.Name)
    Dim AttachArry() As String
    AttachArry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry1(), ListStart, "", "", "", ReadRecord)


    '添付ファイル格納先フォルダーの配列変数
    Dim adressAry6(0) As String
    adressAry6(0) = AttFolder
    Dim AttFolderAry() As String
    AttFolderAry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry6(), ListStart, "", "", "", ReadRecord)

    'WORDテンプレートファイル,格納用の配列変数
    adressAry1() = getInitArray(rec, 2, ActiveSheet.Name)
    adressAry2() = getInitArray(rec, 3, ActiveSheet.Name)
    adressAry3() = getInitArray(rec, 4, ActiveSheet.Name)
    
    '1次元配列から複数次元配列へ
    WordTemp() = arry2(adressAry1, adressAry2)
    WordTemp() = arry2(WordTemp, adressAry3)

  
    w1() = WordTempMail(wrdTmp, WordTemp(0, 1))
    w2() = WordTempMail(wrdTmp, WordTemp(1, 1))
    w3() = WordTempMail(wrdTmp, WordTemp(2, 1))


    'WORDテンプレートファイル内定形文字置換　WORD本文中の置換文字,格納用の配列変数
    adressAry1() = getInitArray(rec, 6, ActiveSheet.Name)
    adressAry2() = getInitArray(rec, 7, ActiveSheet.Name)
    
    '1次元配列から複数次元配列へ
    WordReplace() = arry2(adressAry1, adressAry2)


    'Word差し込みKeyWord
    adressAry1() = getInitArray(rec, 10, ActiveSheet.Name)
    '進捗リスト列番号
    adressAry2() = getInitArray(rec, 11, ActiveSheet.Name)
    

    '進捗リストから取得した、転記用データ格納用の配列変数
    Dim dataAry() As String

    '進捗リストから設定値の読み込み実行
    dataAry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry2(), ListStart, "", "", "", ReadRecord)

    'dataAry()にWordTemp内の検索置換用キーワード（変数名：adressAry1）を1列追加し、3次元配列に
    Dim mailBody() As String
    
    Dim kw As Long
    kw = 1
    
    '定形文字置換　WORD本文中の置換文字追加
    '配列のカウントは、0始まりなので、個数は2なので+1
     
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

       
      '件名の行数探索処理
      Dim s As Long
      For s = 0 To UBound(workBody, 1)
       If InStr(workBody(s), "件名") > 0 Then
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

'    '実績報告書の転記先のセル番地格納用
'    Dim MapingAry() As String
'    '実績報告のセル番地設定 「番地」、行列番号、画面シート名を指定
'    MapingAry() = getInitArray(rec, 7, ActiveSheet.Name)
'
'

'
'
'    '実施報告書への転記処理
''    Call PosToFile(OutPutFolder, "", Jisitemplate, dataAry(), -1, MapingAry(), "実績報告", True)
'
'    Call PosToFile(OutPutFolder, "", Jisitemplate, dataAry(), -1, MapingAry(), templateSheet, True, fileExt, fileFmt)

MsgBox "処理が終了しました"

End Sub

Public Function arry2(ary1() As String, ary2() As String) As String()

   '1次元配列

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
