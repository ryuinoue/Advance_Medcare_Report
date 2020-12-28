Attribute VB_Name = "Tool2"

Public Function getInitArray(rec As Long, Col As Long, SheetName As String) As String()

     '2020.4.16 作成
     '実績報告書転記ツールから実績報告書（転記先）の番地の数を確認
     '指定の行のデータを1行ずつ配列に入れる
     
     '引数    :実績報告書転記ツールのシート名
     '　　　　 読み込み開始位置（行）
     '　　     読み込み開始位置（列）
    
     '戻り値　:転記先のセル番地
     
     '作業用の転記用データ格納用配列変数
     Dim work() As String
     '進捗リストから読み込むレコード数
     Dim l As Long
     'カウンター変数
     Dim r As Long
    
    '
     With ThisWorkbook.Sheets(SheetName)
     
        .Select
        .Cells(1048576, Col).Activate
        
        '進捗リストから読み込むレコード数確認
        Selection.End(xlUp).Select
        l = Selection.Row
        
        r = rec
        
        'カウンター変数
        Dim i As Long
        i = 0
        
        'レコード数から、配列を再設定
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

    '2020.4.15 作成
    '2020.4.24 更新
    
    '進捗リストを開く（リンク更新のメッセージを表示させない）
    '進捗リストから読み込むレコード数を確認する
    '指定の列のデータを1行ずつ配列に入れる
    
    '
    '引数　：進捗リストのファイル名
    '        進捗リストのシート名
    '　　　　進捗リストから抜き取るデータ項目（エクセル:列）
    '　　　  進捗リストの読み込み開始位置
    '　　　　進捗リストに追記するLot番号更新列
    '　　　　進捗リストに追記するLot番号の接頭字
    '　　　　進捗リストに追記するLot番号の接尾字
    '        進捗リストから読み込むレコードを判定するフラグ列
    '
    '
    
    '戻り値：進捗リストから転記用の配列データ
    
    
    '作業用の転記用データ格納用配列変数
    Dim work() As String
    Dim Work2() As String
    
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
        
        
        'レコード数とデータ項目（引数）の数から、配列を再設定
        '実績報告書更新の場合は、要素は、そのまま。
        '実績報告書作成の場合は、要素を1つ追加します。
        '第5引数、第6引数が空欄では無い場合、Lot番号用の要素を追加する。
        
        If Len(LotCol) = 0 And Len(LotNames) = 0 Then
            ReDim work(l - r, UBound(Mapping))
        Else
            ReDim work(l - r, UBound(Mapping) + 1)
        End If
          
        'カウンター変数
        Dim i As Long
        
        'カウンター変数、読み込まないレコード数
        Dim j As Long
        j = 0
        
        'ループ終了条件：処理するレコードが、進捗リストのレコード数を超えたら終了
        Do Until l < r
           
            i = 0
            
            
            '引数[進捗リストから読み込むレコードを判定するフラグ列]に値がある場合、
            'フラグに指定された項目の値によって、データの読み込むか否かの判定を行う
            If Len(ReadingFlg) > 0 Then
            
                '指定された列のデータに値が在れば、読み込み、
                'なければ、読み込まない。
                
                If Len(.Range(ReadingFlg & r).Value) = 0 Then
                
                    '読み込まないレコード数
                    j = j + 1
                    'データを読み込まなない次の処理へ
                    GoTo Next_Step
                
                End If
            
            End If
            
            'セル番地用の変数
            Dim adress As Variant
            
            For Each adress In Mapping
               
               'データ項目（引数）からセル番地を生成
               '転記用データを配列に格納
               work(r - rec, i) = .Range(adress & r).Value
            
               i = i + 1
            
            Next adress
            
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

                work(r - rec, i) = LotNames & "_" & .Range(LotLastName & r).Value
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
    
    
    '配列Workから不要な空白行を削除するための処理
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
''      配列の最初の数字がNullの場合、削除という要件から
''　　　配列の全ての値がNullの場合、削除に要件を変更
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
    
    '作業用の配列データを呼び出し元に返す。
    InputArray = Work2

End Function



Public Function LotFileAry(FOLDERNAME As String) As String()

    '2020.4.22 作成
    '受領先フォルダ内のファイルとLOT番号の配列を作成
    '引数　：受領先フォルダ名
    '戻り値：受領先フォルダ内のファイルとLOT番号の配列


    'Debug.Print WorkBookPath(ThisWorkbook.Path, FolderName, "")
    
    Dim buf As String
    
'    buf = dir(WorkBookPath(ThisWorkbook.Path, FolderName, "") & "\*.xls")
'    buf = dir(Trim(WorkBookPath(ThisWorkbook.Path, FolderName, "") & "\*" & Ext))
    
    buf = Dir(WorkBookPath(ThisWorkbook.Path, FOLDERNAME, "\*" & Ext))
  
    
    Dim work() As String


    'フォルダー内のファイル数をカウント
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
