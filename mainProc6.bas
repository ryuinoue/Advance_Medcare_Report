Attribute VB_Name = "mainProc6"
Public Sub mainProc6()
  
    '2020.05.27 作成
    'フォルダー作成Main関数
    
    '機能
    '実績報告書作成画面より、設定取得
    'フォルダー構造を進捗リストに入力用関数実行
    
        
    With ThisWorkbook.Sheets("業務フォルダー作成")
      
'        '実績報告書出力先フォルダー作成
'        Dim OutPutFolder As String
'        OutPutFolder = .Range("C2").Value
'
'        If FindDirectory(WorkBookPath(ThisWorkbook.Path, OutPutFolder, "")) = True Then
''            If MsgBox("同名のフォルダーが存在します。フォルダーの中を確認してください", vbOKOnly) = vbOK Then
''
''              Exit Sub
''
''
''            End If
'
'        Else
'
'              'フォルダーが無ければ作成
'              MkDir WorkBookPath(ThisWorkbook.Path, OutPutFolder, "")
'
'        End If
      
      
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
      
      
'        '実績報告書template
'        Dim Jisitemplate As String
'        Jisitemplate = .Range("C11").Value
                
        '進捗リストLOT番号追記列
        Dim UpdateLotNum As String
        UpdateLotNum = .Range("C13").Value
        
        '進捗リストLOT番号接頭字
        Dim LOTNAME As String
        LOTNAME = .Range("C14").Value
        
        
        '進捗リストLOT番号接尾字
        Dim LOTLASTNANE As String
        LOTLASTNANE = .Range("C15").Value
      
        'TOPのフォルダー
        Dim folderTop As String
        folderTop = .Range("L2").Value
        
        '試験名
        Dim folderPrtName As String
        folderPrtName = .Range("L3").Value
        
        '進捗リストOCODE列
        Dim OCODE As String
        OCODE = .Range("L4").Value
        
        '進捗リストONAME列
        Dim ONAME As String
        ONAME = .Range("L5").Value
        
        '進捗リストフォルダー追記列
        Dim UpdateFolder As String
        UpdateFolder = .Range("L8").Value
        
        
      
    End With
    
    '進捗リスト、実績報告書設定読み込み開始行
    Dim rec As Long
    rec = 10
    
'    '進捗リストの列情報を格納用の配列変数
'    Dim adressAry() As String
'    '進捗リストの項目設定 「列」の行、列番号、画面シート名を指定
'    adressAry() = getInitArray(rec, 4, "実績報告書作成")
    

    '進捗リストに出力する作業STEPの情報取得
    Dim adressAry() As String
    '進捗リストの項目設定 「列」の行、列番号、画面シート名を指定
    adressAry() = getInitArray(rec, 11, "業務フォルダー作成")

 
     '業務フォルダーTOPの文字列作成
     '年度の追加
            
        folderTop = WorkBookPath(folderTop, "", "")

      If FindDirectory(folderTop) = True Then
        '
      Else
            MkDir folderTop
      End If
        
     '試験名の追加
        folderTop = WorkBookPath(folderTop, folderPrtName, "")
        
      If FindDirectory(folderTop) = True Then
        '
      Else
        MkDir folderTop
     End If
        
    Dim adressAry2() As String
    Dim aA2 As Variant
     '進捗リスト更新　フォルダー情報入力
    adressAry2 = MkFolder(ShintyokuFile, ShintyokuSheet, adressAry(), ListStart, UpdateLotNum, LOTNAME, LOTLASTNANE, ReadRecord, folderTop, OCODE, ONAME, UpdateFolder)
    

'    進捗リストに値を挿入した列を返す

    With ThisWorkbook.Sheets("業務フォルダー作成")
    Dim i As Long
    For Each aA2 In adressAry2
        .Cells(rec + 1 + i, 11 + 1).Value = aA2
        i = i + 1
    Next
    End With

'    '実績報告書の転記先のセル番地格納用
'    Dim MapingAry() As String
'    '実績報告のセル番地設定 「番地」、行列番号、画面シート名を指定
'    MapingAry() = getInitArray(rec, 7, "実績報告書作成")


'    '進捗リストから取得した、転記用データ格納用の配列変数
'    Dim dataAry() As String
    
'    '設定値の読み込み実行
'    dataAry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry(), ListStart, UpdateLotNum, LOTNAME, LOTLASTNANE, ReadRecord)
     
    '実施報告書への転記処理
'    Call PosToFile(OutPutFolder, "", Jisitemplate, dataAry(), -1, MapingAry(), "実績報告", True)

MsgBox "処理が終了しました"

End Sub


Public Function MkFolder(filename As String, SheetName As String, Mapping() As String, rec As Long, LotCol As String, LotNames As String, LotLastName As String, ReadingFlg As String, folderTop As String, OCODE As String, ONAME As String, UpdateFolder As String) As String()

    '2020.5.25 作成

    
    '進捗リストを開く（リンク更新のメッセージを表示させない）
    '進捗リストから読み込むレコード数を確認する
    '指定の列のデータを1行ずつ配列に入れる
    
    '
    '引数　：進捗リストのファイル名
    '        進捗リストのシート名
    '　　　　進捗リストに入力するデータ（エクセル:列）
    '　　　  進捗リストの読み込み開始位置
    '　　　　進捗リストに追記するLot番号更新列
    '　　　　進捗リストに追記するLot番号の接頭字
    '　　　　進捗リストに追記するLot番号の接尾字
    '        進捗リストから読み込むレコードを判定するフラグ列
    '　　　　フォルダーを作成する場所
    '　　　　進捗リストのOCODEの列
    '　　　　進捗リストのONAMEの列
    '　　　　進捗リストに追記するフォルダー
    
    
'    '作業用の転記用データ格納用配列変数
'    Dim Work() As String
'    Dim Work2() As String
'
    '進捗リストから読み込むレコード数
    Dim l As Long
    
    'カウンター変数
    Dim r As Long
    r = rec
    
    '進捗リストファイルオープン
    Dim TargetBook As Workbook
    Set TargetBook = Workbooks.Open(WorkBookPath(ThisWorkbook.Path, "", filename), False, 0)


    '進捗リストファイル内部の作業に遷移
    With TargetBook.Sheets(SheetName)
          
        '進捗リストから読み込むレコード数確認
        .Select
        .Cells(1048576, rec).Activate
        
        Selection.End(xlUp).Select
        l = Selection.Row
        
        
'        'レコード数とデータ項目（引数）の数から、配列を再設定
'        '実績報告書更新の場合は、要素は、そのまま。
'        '実績報告書作成の場合は、要素を1つ追加します。
'        '第5引数、第6引数が空欄では無い場合、Lot番号用の要素を追加する。
'
'        If Len(LotCol) = 0 And Len(LotNames) = 0 Then
'            ReDim Work(l - r, UBound(Mapping))
'        Else
'            ReDim Work(l - r, UBound(Mapping) + 1)
'        End If
'
        'カウンター変数
        Dim i As Long
        
        'カウンター変数、読み込まないレコード数
        Dim j As Long
        j = 0
        
        
        '呼び出し元に、進捗リストの列番号を返す値を格納する配列
        Dim oMp() As String
        ReDim oMp(UBound(Mapping))
        
        'ループ終了条件：処理するレコードが、進捗リストのレコード数を超えたら終了
        Do Until l < r
           
            i = 0
            
            
            '引数[進捗リストから読み込むレコードを判定するフラグ列]に値がある場合、
            'フラグに指定された項目の値によって、データの読み込むか否かの判定を行う
            If Len(ReadingFlg) > 0 Then
            
                '指定された列のデータに値が在れば、処理を行う
                'なければ、読み込まない。
                
                If Len(.Range(ReadingFlg & r).Value) = 0 Then
                
                    '読み込まないレコード数
                    j = j + 1
                    'データを読み込まなない次の処理へ
                    GoTo Next_Step
                
                End If
            
            End If
            
'            'セル番地用の変数
'            Dim adress As Variant
'
'            For Each adress In Mapping
'
'               'データ項目（引数）からセル番地を生成
'               '転記用データを配列に格納
'               Work(r - rec, i) = .Range(adress & r).Value
'
'               i = i + 1
'
'            Next adress
            
            '実績報告書更新の場合は、何もしない。
            '実績報告書作成の場合は、LOT番号生成して、転記用の配列、進捗リストに追記
            
            If Len(LotCol) = 0 And Len(LotNames) = 0 Then
            
               '
            
'            '2020/5/19 追加
'            ElseIf Len(LotCol) = 0 And Len(LotNames) <> 0 Then
'
'                Work(r - rec, i) = LotNames & "_" & .Range(LotLastName & r).Value
            
            Else
               
               '第5引数、第6引数が空欄では無い場合（実績報告書作成）、一番後ろの要素にLot番号を追加する。
               
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
                      'フォルダーが無ければ作成
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
    
    '進捗リストを閉じる。メモリー開放
    TargetBook.Close SaveChanges:=True
    Set TargetBook = Nothing
    
    DoEvents
    
'
'    '配列Workから不要な空白行を削除するための処理
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
    
    '作業用の配列データを呼び出し元に返す。
    MkFolder = oMp

End Function
