Attribute VB_Name = "mainProc7"
Public Sub mainProc7()
  
    '2020.06.02 作成
    'ファイルコピーMain関数
    
    '機能
    '各画面より、設定取得
    '設定先のフォルダーにファイルをコピー
    
    
    
    'ファイル名格納列
    Dim FileCol As String
    
    '一時フォルダー名格納
    Dim FldName As String

        
    
    Dim rng As Range
        
    With ThisWorkbook.ActiveSheet
    
        '一時フォルダーのフォルダー名を取得
        With .Columns(RETUB)
        
            '完全一致
            Set rng = .Find(FOLDERNAME, LookAt:=xlWhole)
            ', xlValues, xlWhole, xlByColumns, xlNext
            
            If rng Is Nothing Then
            
                MsgBox "設定を確認してください"
                Exit Sub
                
            End If
            

            FldName = Range(rng.Address).Offset(0, 1).Value
            
        End With
    
    
        'ファイル名の保存先を取得
        With .Columns(RETUC)
        
            Set rng = .Find(filename)
            ', xlValues, xlWhole, xlByColumns, xlNext
            If rng Is Nothing Then
            
                MsgBox "設定を確認してください"
                Exit Sub
                
            End If
            

            FileCol = Range(rng.Address).Offset(0, 1).Value
            
        End With
        
        
        
        '読み込む拡張子
        Dim kakutyoushi As String
        kakutyoushi = .Range("U19").Value
      
      
      
      
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
'
        '進捗リスト読み込み開始行
        Dim ListStart As Long
        ListStart = 2
'
        '進捗リスト読み込みフラグ列
        Dim ReadRecord As String
        ReadRecord = .Range("C9").Value
      
      
'        '実績報告書template
'        Dim Jisitemplate As String
'        Jisitemplate = .Range("C11").Value
'
'        '進捗リストLOT番号追記列
'        Dim UpdateLotNum As String
'        UpdateLotNum = .Range("C13").Value
'
'        '進捗リストLOT番号接頭字
'        Dim LOTNAME As String
'        LOTNAME = .Range("C14").Value
'
'
'        '進捗リストLOT番号接尾字
'        Dim LOTLASTNANE As String
'        LOTLASTNANE = .Range("C15").Value
      
    End With
    
    '設定読み込み開始行
    Dim rec As Long
    rec = 18

'
'    '進捗リストの列情報を格納用の配列変数
    Dim adressAry() As String
'    '進捗リストの項目設定 「列」の行、列番号、画面シート名を指定
    adressAry() = getInitArray(rec, 17, ThisWorkbook.ActiveSheet.Name)
'
'
'    '実績報告書の転記先のセル番地格納用
'    Dim MapingAry() As String
'    '実績報告のセル番地設定 「番地」、行列番号、画面シート名を指定
'    MapingAry() = getInitArray(rec, 7, "実績報告書作成")
'

'ファイル名、出力先フォルダーを格納する1次元配列を作成(※これは、画面要件と確定しているので、2行、1列の配列を最初から用意する)
    Dim adressAry2(1) As String
    adressAry2(0) = FileCol
    adressAry2(1) = adressAry(0)

'
    '進捗リストから取得した、転記用データ格納用の配列変数
    Dim dataAry() As String
'
    '設定値の読み込み実行
    dataAry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry2(), ListStart, "", "", "", ReadRecord)

'   ファイルコピー
    Call FloderToCopy(FldName, dataAry(), kakutyoushi)

'    '実施報告書への転記処理
'    Call PosToFile(OutPutFolder, "", Jisitemplate, dataAry(), -1, MapingAry(), "実績報告", True)

MsgBox "処理が終了しました"

End Sub


Public Sub FloderToCopy(FldName As String, data() As String, kakutyoushi As String)

    '2020.6.2 作成
    
    '進捗リストの配列データを元に一時フォルダーから、移動先のフォルダーにファイルをコピーする
    '指定のファイル名で保存
    
    '
    '引数　：一時フォルダー名
    '        ファイル名と移動先のフォルダー
    '　　　　実績報告テンプレートファイル
    
  
    '繰り返し回数の条件設定。
    ' 進捗リストより取得したレコード数分
    
    
    Dim RecordNo As Long
    RecordNo = 0
    
    '「作成」の場合、進捗リストの全レコード繰り返すため、レコード数をカウンター変数に代入
    
    
      RecordNo = UBound(data)


    '（繰り返し）ファイルコピー処理
    Dim i As Long
    Dim j As Long
    
'    Dim TargetBook As Workbook
    
    For i = 0 To RecordNo
      
'        '実績報告ファイルテンプレートをオープン
'        '「作成」の場合は、テンプレート
'        '「更新」の場合は、受領したファイル
'        Set TargetBook = Workbooks.Open(WorkBookPath(ThisWorkbook.Path, AcptFileDir, TmpFile), False, 0)


'        一時フォルダー内にファイルがある事を確認する。
'　　　　一時フォルダーは、ツールと同じディレクトリーに作成される。この要件は固定とする。

        If FindFile(WorkBookPath(ThisWorkbook.Path, FldName, data(i, 0) & Trim(kakutyoushi))) = False Then
        
            'ファイルが存在しない場合は、次の処理へ
             GoTo Next_Step
            
        End If
    
          Dim strFile2 As String
            strFile2 = WorkBookPath(data(i, 1), "", data(i, 0) & Trim(kakutyoushi))
            
           Dim strFile As String
           strFile = WorkBookPath(ThisWorkbook.Path, FldName, data(i, 0) & Trim(kakutyoushi))
          FileCopy strFile, strFile2
          
          
'        '実績報告ファイル内部の作業に遷移
'         With TargetBook.Sheets(SheetName)
'
'              'item=-1の場合、カウンター変数は、カウントアップ
'              If Item = -1 Then
'              '
'              Else
'              'item<>-1の場合、カウンター変数は、itemの固定値
'                 i = Item
'              End If
'
'
'              For j = 0 To UBound(Mapping)
'
''               Debug.Print Mapping(j); ":"; data(i, j)
'                .Range(Mapping(j)).Value = data(i, j)
'
'              Next j
'
'              '一番最後の要素に入力したLOT番号をフッターに入力
'              j = j + 1
'
'              If flg Then
'                  .PageSetup.RightFooter = data(i, j)
''                  Debug.Print .PageSetup.RightFooter
'              Else
'                  .PageSetup.RightFooter = Null
'              End If
'
'         End With
         
         'Debug.Print CInt(flg)
         
         '転記配列、
         '「作成」の場合、最後から一つ前の要素　全要素数＋True(=-1)
         '「更新」の場合、最後の要素
         '　に入力された値をファイル名として保存する｡
'         TargetBook.SaveAs _
'                    FILENAME:=WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg)) & Ext), _
'                    FileFormat:=xlsFmt
'         TargetBook.Close
         
Next_Step:
         
         DoEvents
         
     Next i
'
  
End Sub





