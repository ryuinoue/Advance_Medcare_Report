Attribute VB_Name = "mainProc2"
Public Sub mainProc2()

    '2020.04.24 作成
    '実績報告書更新Main関数
    
    '機能
    '実績報告書更新画面より、設定取得
    '実績報告書作成関数実行


    With ThisWorkbook.Sheets("実績報告書更新")
    
    
        '実績報告書受領先フォルダー確認
        Dim AcptFolder As String
        AcptFolder = .Range("C2").Value
      
        If FindDirectory(WorkBookPath(ThisWorkbook.Path, AcptFolder, "")) = False Then
        
            If MsgBox("受領先のフォルダーが存在していません", vbOKOnly) = vbOK Then
              
              Exit Sub
            
            End If
        
        End If
    
    
        Dim fileExt As String
        fileExt = .Range("G2").Value
        
        Dim fileFmt As Long
        fileFmt = .Range("H2").Value
    
    
        '実績報告書(更新)出力先フォルダー作成
        Dim OutPutFolder As String
        OutPutFolder = .Range("C11").Value
      
        If FindDirectory(WorkBookPath(ThisWorkbook.Path, OutPutFolder, "")) = True Then
            
            If MsgBox("同名のフォルダーが存在します。フォルダーの中を確認してください", vbOKOnly) = vbOK Then

'                Exit Sub


            End If
'
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
        
        
        
    End With

    '進捗リスト、実績報告書設定読み込み開始行
    Dim rec As Long
    rec = 15

    '回収したファイルのファイル名、LOTNOの配列変数
    Dim dataAry() As String
    '実績報告書受領先フォルダー名を指定
    dataAry() = LotFileAry(AcptFolder)


    '進捗リストの列情報を格納用の配列変数
    Dim adressAry() As String
    '進捗リストの項目設定 「列」の行、列番号、画面シート名を指定
    adressAry() = getInitArray(rec, 4, "実績報告書更新")

    '進捗リストから取得した、転記用データ格納用の配列変数
    Dim dataAry2() As String
    dataAry2() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry(), 2, "", "", "", ReadRecord)



    'ロジック確認
    '転記RAWデータ = 全受領したファイルのデータ(ファイル名、LOT番号)+進捗リストのデータ(進捗リストの項目...,LOT番号) ※LOT番号をキーとしてマージ

    '総転記レコード数=回収したファイル数
    '転記RAWデータは、回収したファイルから取得したデータの項目数(LOT番号、ファイル名)＋進捗リストの項目数
    '転記RAWデータ格納用配列データ(ファイル名,進捗リストの項目1,進捗リストの項目2,…,…)　※LOT番号は、併合後使用しないデータなので格納しない。
    Dim dataAry3() As String
    ReDim dataAry3(UBound(dataAry()), (UBound(dataAry(), 2) - 1) + UBound(adressAry(), 1))

    
    Dim i As Long
    Dim j As Long

    '繰り返しの処理の数は、レコード数は＝ファイルの数、即ち、dataAryの添え字1番目の最大要素数
    For i = 0 To UBound(dataAry(), 1)
    
     'dataAry:回収したファイルの配列
     'dataAry2:進捗リストから取得したデータの配列
     '2つの配列の中にあるLOT番号が同じもの探索
     
     '探索回数は、進捗リストのレコード数
      For j = 0 To UBound(dataAry2(), 1)
        
           Dim k As Long
           k = 0
           '2つの配列の中にあるLOT番号が同じ場合、
           If Trim(dataAry(i, 0)) = Trim(dataAry2(j, UBound(dataAry2(), 2))) Then
           
               '転記RAWデータ格納用配列データの添え字1に、ファイル名を入力
                dataAry3(i, k) = dataAry(i, 1)
                
                '転記RAWデータ格納用配列データの添え字2以降は、進捗リストのデータ
                Dim l As Long
                l = 0
                For k = 1 To UBound(dataAry3(), 2)
                
                    dataAry3(i, k) = dataAry2(j, k - 1)
                    
                Next k
           
           End If
      Next j
    Next i
    
    'メモ
    'dataAry3は、更新処理に必要なデータすべてが入った配列が完成
    '検証が必要になった場合、まずdataAry3の内容を確認する。


    '転記RAWデータ格納用配列より、ファイル名項目を抜き出した配列
    Dim dataAry4() As String
    ReDim dataAry4(UBound(dataAry3(), 1))
    
    '転記RAWデータ格納用配列より、進捗リスト即ち転記用の項目を抜き出した配列
    Dim dataAry5() As String
    ReDim dataAry5(UBound(dataAry3()), UBound(dataAry3(), 2) - 1)
    
    'Debug.Print UBound(dataAry3(), 2) - 1
    
    Dim m As Long
    Dim n As Long
    
    '繰り返し処理の回数は、転記RAWデータ格納用配列のレコード数、即ち添え字1番目の最大要素数
    For m = 0 To UBound(dataAry3, 1)
    
        'ファイル名を入力
        dataAry4(m) = dataAry3(m, 0)
        
        'ファイル名以外の進捗リストの項目を入力
        For n = 0 To UBound(dataAry3, 2) - 1
            dataAry5(m, n) = dataAry3(m, n + 1)
        Next n
        
    Next m


    '実績報告書の転記先のセル番地格納用
    Dim MapingAry() As String
    '実績報告のセル番地設定 「番地」、行列番号、画面シート名を指定
    MapingAry() = getInitArray(rec, 7, "実績報告書更新")


    Dim o As Long
    For o = 0 To UBound(dataAry4)

        'Call PosToFile(OutPutFolder, Jisitemplate, dataAry(), MapingAry(), "実績報告", "作成")
        Call PosToFile(OutPutFolder, AcptFolder, dataAry4(o), dataAry5(), o, MapingAry(), "実績報告", False, fileExt, fileFmt)
 
    Next o

'Debug.Print '

MsgBox "処理が終了しました"

End Sub


