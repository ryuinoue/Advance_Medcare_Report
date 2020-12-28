Attribute VB_Name = "mainProc5"
Public Sub mainProc5()
  
    '2020.05.19 作成
    '患者登録状況一覧作成Main関数
    
    '機能
    '患者登録状況一覧作成画面より、設定取得
    '患者登録状況一覧作成関数実行
    
        
    With ThisWorkbook.Sheets("患者登録状況一覧作成")
      
        '患者登録状況一覧出力先フォルダー作成
        Dim OutPutFolder As String
        OutPutFolder = .Range("C2").Value
      
        If FindDirectory(WorkBookPath(ThisWorkbook.Path, OutPutFolder, "")) = True Then
'            If MsgBox("同名のフォルダーが存在します。フォルダーの中を確認してください", vbOKOnly) = vbOK Then
'
'              Exit Sub
'
'
'            End If
            
        Else
            
              'フォルダーが無ければ作成
              MkDir WorkBookPath(ThisWorkbook.Path, OutPutFolder, "")
        
        End If
      
      
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
      
        '進捗リスト読み込み開始行
        Dim ListStart As Long
        ListStart = 2
      
        '進捗リスト読み込みフラグ列
        Dim ReadRecord As String
        ReadRecord = .Range("C9").Value
      
        '確認結果出力先（オフセット番地）
        Dim ofset As Long
        ofset = .Range("C10").Value
      
      
        '患者登録状況一覧template
        Dim Jisitemplate As String
        Jisitemplate = .Range("C11").Value
                
        '進捗リストLOT番号追記列
        Dim UpdateLotNum As String
        UpdateLotNum = .Range("C13").Value
        
        '進捗リストLOT番号接頭字
        Dim LOTNAME As String
        LOTNAME = .Range("C14").Value
        
        
        '進捗リストLOT番号接尾字
        Dim LOTLASTNANE As String
        LOTLASTNANE = .Range("C15").Value
        
        
        '------------------------------------------------------------------------------------------------
        'MDB情報
        'SAS登録数データ.MDB保存先フォルダー
        Dim MDBDir As String
        MDBDir = .Range("G3").Value


        'SAS登録数データ.MDBのファイル名
        Dim mdbFile As String
        mdbFile = .Range("H4").Value


        'テーブル名
        Dim mdbTbl As String
        mdbTbl = .Range("H5").Value
        
        '並び順
        Dim tblorder As String
        tblorder = .Range("I13").Value
        
        '患者登録状況一覧の抽出条件セル番地格納用
        Dim whereAry() As String
        
        '実績報告のセル番地設定 「番地」、行列番号、画面シート名を指定
        whereAry() = getInitArray(7, 10, "患者登録状況一覧作成")
        
        '------------------------------------------------------------------------------------------------
      
    End With
    
    '進捗リスト、実績報告書設定読み込み開始行
    Dim rec As Long
    rec = 19
    
    '進捗リストの列情報を格納用の配列変数
    Dim adressAry() As String
    '進捗リストの項目設定 「列」の行、列番号、画面シート名を指定
    adressAry() = getInitArray(rec, 4, "患者登録状況一覧作成")


    '実績報告書の転記先のセル番地格納用
    Dim MapingAry() As String
    '実績報告のセル番地設定 「番地」、行列番号、画面シート名を指定
    MapingAry() = getInitArray(rec, 7, "患者登録状況一覧作成")


    '------------------------------------------------------------------------------------------------
    '実績報告書の転記先のセル番地格納用
    Dim itemAry() As String
    '実績報告のセル番地設定 「番地」、行列番号、画面シート名を指定
    itemAry() = getInitArray(18, 9, "患者登録状況一覧作成")


    '進捗リストから取得した、施設情報の配列変数
    Dim dataAry() As String
    
    
    '設定値の読み込み実行
    '※進捗リストから、MDBファイルより施設毎のデータを抽出するためのキーコード「OCODE」
    '  LOTNUMの配列を取得
    
'     dataAry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry(), ListStart, "", LOTNAME, LOTLASTNANE, ReadRecord)
    dataAry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry(), ListStart, "", "", "", ReadRecord)
     
    
    'MDB（SAS登録数データ）から取得した転記用データ格納用の配列変数
    Dim mdbdataAry() As String
    Dim mdbrec As Long
    Dim o As Long
    Dim LOTNUM As String

    For o = 0 To UBound(dataAry, 1)
    
        'MDBより取り込み
        Call InputArrayMDB(MDBDir, mdbFile, mdbTbl, tblorder, itemAry(), whereAry(), dataAry(), o, mdbdataAry(), LOTNUM, mdbrec)
        
        If mdbrec > 0 Then
        
            '患者登録状況一覧への転記処理
            Call PosToFileMult(OutPutFolder, Jisitemplate, mdbdataAry(), dataAry(), o, "Sheet1")

        End If
        
        '配列の解放
        Erase mdbdataAry
        
'　　　　進捗リストに出力する登録数は、数値として使用するの「件」は、削除
'        Call InPutQuery(ShintyokuFile, ShintyokuSheet, LOTNUM, mdbrec & "件", adressAry(UBound(adressAry)), ofset)
        Call InPutQuery(ShintyokuFile, ShintyokuSheet, LOTNUM, CStr(mdbrec), adressAry(UBound(adressAry)), ofset)
        
        
    Next o
'Public Function PosToFileMult(Filedir As String, TmpFile As String, DATA() As String, LOTNUM As String, SheetName As String) As Boolean

MsgBox "処理が終了しました"

End Sub

Public Function InputArrayMDB(mdbir As String, mdbFile As String, mdbTbl As String, tblOdr As String, itemAry() As String, whereAry() As String, dataAry() As String, o As Long, rsAry() As String, LOTNUM As String, mdbrec As Long) As String

    '2020.5.19 作成
    
    'MDB（SAS登録数データ）を開きレコードセットを作成
    'MDBから読み込むレコード数を確認する
    '指定の列のデータを1行ずつ配列に入れる
    
    '
    '引数　：MDBのファイルの保存先ディレクトリ
    '        MDBのファイル名
    '　　　　テーブル名
    '　　　  抽出項目
    '　　　　抽出条件
    '　　　　設定値　※OCODE
    '
    '
    
    '戻り値：SQL文
    

    'MDBファイルよりデータ抽出用SQL文作成

  
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
    
    'OCODEは、配列の一番最初に入ってくることを前提とする。
    Dim strsql As String
    strsql = "select " & workItem & " from " & mdbTbl & " where " & workFrom & " and OCODE='" & dataAry(o, 0) & "' order by " & tblOdr
    
'    Debug.Print strsql
    
    '------------------------------------------------------------------------------------------------
    
    '------------------------------------------------------------------------------------------------
    'MDB（SAS登録数データ）への接続
    
      Dim cn As New ADODB.Connection
      Dim rs As New ADODB.Recordset
      Dim ConString As String
      
      ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & WorkBookPath(mdbir, "", Trim(mdbFile))
      
      cn.Open ConnectionString:=ConString
      
     'レコードセット作成
     rs.Open Source:=strsql, ActiveConnection:=cn, CursorType:=adOpenKeyset, LockType:=adLockOptimistic
     
'    Debug.Print rs.RecordCount & "," & rs.Fields.Count
      
    '------------------------------------------------------------------------------------------------
        
    '抽出した結果、レコードカウント0の場合は、処理をしない。
    If rs.RecordCount = 0 Then
     
'        データを読み込まなない次の処理へ
        GoTo Next_Step
     
    End If

    'レコード数とデータ項目の数から、配列を再設定
     ReDim rsAry(rs.RecordCount - 1, rs.Fields.Count - 1)


    'カウンター変数
    Dim r As Long
    Dim i As Long

    '配列にデータを入れる
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
    
    'LOTNUM取得
    LOTNUM = dataAry(o, UBound(dataAry, 2))
    
    rs.Close
    Set rs = Nothing
    
    cn.Close
    Set cn = Nothing
    
    DoEvents


End Function


'Public Function PosToFileWrd(Filedir As String, AcptFileDir As String, TmpFile As String, DATA() As String, Item As Long, flg As Boolean) As Boolean

Public Function PosToFileMult(Filedir As String, tmpfile As String, mdbDATA() As String, data() As String, o As Long, SheetName As String) As Boolean

    '2020.5.19 作成
    
    '配列データを転記する
    '指定のファイル名で保存
    
    '
    '引数　：作成するファイルの保存先
    '　　　　テンプレートファイル
    '　　　　転記配列データ
    '　　　　シート名
    '
    '戻り値：作業結果
    
    Dim r As Long
    Dim i As Long
    
    
    ' 最後の要素に入力された値が、Nullもしくは、"-"の場合は、
    '処理を飛ばす｡
    
    If Len(data(o, UBound(data, 2) + -1)) <= 1 Then
     'ファイルが存在しない場合は、次の処理へ
     GoTo Next_Step
        
    End If
    
    Dim TargetBook As Workbook
    
    Set TargetBook = Workbooks.Open(WorkBookPath(ThisWorkbook.Path, "", tmpfile), False, 0)
    
          
        'ファイル内部の作業に遷移
         With TargetBook.Sheets(SheetName)
         
              For r = 0 To UBound(mdbDATA, 1)
      
                For i = 0 To UBound(mdbDATA, 2)
                .Cells(r + 3, i + 1).Value = mdbDATA(r, i)
                Next i
                
              Next r
              
              'LOT番号をフッターに入力
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
         
        'XLSXで良いのここは、固定
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
         
'ファイルが存在しない場合は、次の処理へ
Next_Step:

End Function
