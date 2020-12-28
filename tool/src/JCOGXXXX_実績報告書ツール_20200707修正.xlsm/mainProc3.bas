Attribute VB_Name = "mainProc3"

Public Sub mainProc3()

    '2020.04.24 作成
    '実績報告書確認Main関数
    
    '機能
    '実績報告書更新画面より、設定取得
    '実績報告書作成関数実行


    With ThisWorkbook.Sheets("実績報告書確認")
    
    
        '実績報告書受領先フォルダー確認
        Dim AcptFolder As String
        AcptFolder = .Range("C2").Value
      
        If FindDirectory(WorkBookPath(ThisWorkbook.Path, AcptFolder, "")) = False Then
        
            If MsgBox("受領先のフォルダーが存在していません", vbOKOnly) = vbOK Then
              
              Exit Sub
            
            End If
        
        End If
    
    
'実施報告書には噴出しをつけない仕様に変更
'        '実績報告書(更新)出力先フォルダー作成
'        Dim OutPutFolder As String
'        OutPutFolder = .Range("C11").Value
'
'        If FindDirectory(WorkBookPath(ThisWorkbook.Path, OutPutFolder, "")) = True Then
'
'            If MsgBox("同名のフォルダーが存在します。フォルダーの中を確認してください", vbOKOnly) = vbOK Then
'
''                Exit Sub
'
'
'            End If
''
'        Else
'
'                'フォルダーが無ければ作成
'                MkDir WorkBookPath(ThisWorkbook.Path, OutPutFolder, "")
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
        
        '確認結果出力先（オフセット番地）
        Dim ofset As Long
        ofset = .Range("C10").Value
        
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
    adressAry() = getInitArray(rec, 4, "実績報告書確認")

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
    MapingAry() = getInitArray(rec, 7, "実績報告書確認")


    '実績報告書の転記先のセル番地格納用
    Dim itemAry() As String
    '実績報告のセル番地設定 「番地」、行列番号、画面シート名を指定
    itemAry() = getInitArray(rec, 6, "実績報告書確認")


    Dim o As Long
    
    Dim LOTNUM As String
    Dim alrtMsg As String
    
    For o = 0 To UBound(dataAry4)

        'Call PosToFile(OutPutFolder, Jisitemplate, dataAry(), MapingAry(), "実績報告", "作成")
        'Call PosToCheckFile(OutPutFolder, AcptFolder, dataAry4(o), dataAry5(), o, MapingAry(), ItemAry(), "実績報告", False)
        
    
        Call PosToCheckFile(AcptFolder, dataAry4(o), dataAry5(), o, MapingAry(), itemAry(), "実績報告", False, LOTNUM, alrtMsg)
        
        '    dataAry2() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry(), 2, "", "", ReadRecord)
        
        Call InPutQuery(ShintyokuFile, ShintyokuSheet, LOTNUM, alrtMsg, adressAry(UBound(adressAry)), ofset)
        
    Next o

'Debug.Print '

MsgBox "処理が終了しました"

End Sub


'Public Function PosToCheckFile(Filedir As String, AcptFileDir As String, TmpFile As String, DATA() As String, Item As Long, Mapping() As String, ItemAry() As String, SheetName As String, flg As Boolean) As Boolean
    '
    '引数　：作成した実績報告書の保存先(×)
    '指定のファイル名で保存

    '戻り値：作業結果

Public Function PosToCheckFile(AcptFileDir As String, tmpfile As String, data() As String, Item As Long, Mapping() As String, itemAry() As String, SheetName As String, flg As Boolean, LOTNUM As String, alrtMsg As String)

    '2020.4.15 作成
    
    '上部（基本情報）のチェック内容
    '進捗リストの配列データと指定のセルの転記先（セル番地）値を比較し違っていれば、セルの網掛けを赤く。
    '進捗リストの値をコメントに入れる
    

    
    '引数
    '       入力
    '        1)受領した実績報告書の保存先
    '　　　　2)実績報告テンプレートファイル
    
    '　　　　3)進捗リスト転記用配列データ
    
    '　　　　「確認」の配列構造デフォルト※「作成」に準ずる
    '　　　　　DATA(ONAME,届出受理年月日,先進医療の費用 (届出時),人件費,技術名,ファイル名,LOT番号)
    '　　　　　最後と最後から1つ前の要素は固定、他は設定により変更可
    
    '        4)進捗リスト転記用配列データの添え字。
    ' 　　　　　　-1の場合、カウンター変数は、カウントアップ。
    '           　-1では無い場合は、固定値
    
    
    '　　　　5)転記先（セル番地）
    '        6)実績報告書転記ツールのシート名
    '        7)フラグ
    '　　　　　　　「作成」の場合、True(=-1)
    '              「更新」の場合、False(=0)
            

    '        出力
    '　　　　8)LOT番号 ブランクで渡される。関数の中で、実績報告書のフッターの値が入力される。戻り先のプロシジャーで使用される。
    '　　　　9)アラートメッセージ　ブランクで渡される。関数の中で、アラートメッセージが入力される。戻り先のプロシジャーで使用される。
             
    
  
    '繰り返し回数の条件設定。
    '「作成」の場合、進捗リストより取得したレコード数分
    '「更新」の場合、1回のみ
    'カウンター変数、初期値は0。即ち、繰り返し回数は、1回
    Dim RecordNo As Long
    RecordNo = 0
    
    '「作成」の場合、進捗リストの全レコード繰り返すため、レコード数をカウンター変数に代入
    '「更新」の場合、進捗リスト1レコード分のデータのみ、渡される。0を代入
    If flg Then
      RecordNo = UBound(data)
    End If

    '（繰り返し）転記処理
    Dim i As Long
    Dim j As Long
    Dim TargetBook As Workbook


alrtMsg = ""

    For i = 0 To RecordNo
      
        '実績報告ファイルテンプレートをオープン
        '「作成」の場合は、テンプレート
        '「更新」の場合は、受領したファイル
        Set TargetBook = Workbooks.Open(WorkBookPath(ThisWorkbook.Path, AcptFileDir, tmpfile), False, 0)
    
          
        '進捗リストファイル内部の作業に遷移
         With TargetBook.Sheets(SheetName)
         
              'item=-1の場合、カウンター変数は、カウントアップ
              If Item = -1 Then
              '
              Else
              'item<>-1の場合、カウンター変数は、itemの固定値
                 i = Item
              End If
              

              For j = 0 To UBound(Mapping)
      
'               Debug.Print Mapping(j); ":"; data(i, j)
                
                If .Range(Mapping(j)).Value <> data(i, j) Then
                

'                    実績報告書に噴出しをつける場合の処理。Excelのバージョンによってはエラーとなるためコメントアウト
'                    .Range(Mapping(j)).AddCommentThreaded ("進捗リストの値:" & CStr(DATA(i, j)))

'                    Debug.Print ItemAry(j) & " 進捗リスト:" & DATA(i, j) & " <-> 実績報告書書:" & .Range(Mapping(j)).Value
                    
                    alrtMsg = alrtMsg & "【" & itemAry(j) & "】 進捗リスト「" & data(i, j) & "」  実績報告書書（施設入力値）「" & IIf(Len(.Range(Mapping(j)).Value) = 0, "null", .Range(Mapping(j)).Value) & "」" & vbCrLf
                
                End If
                
               
              Next j

              
'確認用モジュールなので、LOT番号の処理は何もしない。
'              '一番最後の要素に入力したLOT番号をフッターに入力
'              j = j + 1
'
'              If flg Then
'                  .PageSetup.RightFooter = DATA(i, j)
''                  Debug.Print .PageSetup.RightFooter
'              Else
'                  .PageSetup.RightFooter = Null
'              End If
              
              LOTNUM = .PageSetup.RightFooter
              
         End With
         
         'Debug.Print CInt(flg)
         
         '進捗リスト転記配列、
         '「作成」の場合、最後から一つ前の要素　全要素数＋True(=-1)
         '「更新」の場合、最後の要素
         '　に入力された値をファイル名として保存する｡
         
'         TargetBook.SaveAs _
'                    FileName:=WorkBookPath(ThisWorkbook.Path, Filedir, DATA(i, UBound(DATA, 2) + CInt(flg)) & Ext), _
'                    FileFormat:=xlsFmt
         
         TargetBook.Close False
         
         DoEvents
         
     Next i
'
'  Debug.Print alrtMsg
  
End Function

