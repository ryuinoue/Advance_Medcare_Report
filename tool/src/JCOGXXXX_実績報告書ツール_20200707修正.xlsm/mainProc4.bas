Attribute VB_Name = "mainProc4"
Public Sub mainProc4()
  
    '2020.05.15 作成
    '総括報告書作成Main関数
    
    '機能
    '実績報告書作成画面より、設定取得
    '総括報告書関数実行
    
        
    With ThisWorkbook.Sheets("総括報告書作成")
      
        '実績報告書出力先フォルダー作成
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
      
      
        '実績報告書template
        Dim Jisitemplate As String
        Jisitemplate = .Range("C11").Value

'総括報告書Word版は、作成のみ
'        '進捗リストLOT番号追記列
'        Dim UpdateLotNum As String
'        UpdateLotNum = .Range("C13").Value
'
'        '進捗リストLOT番号接頭字
'        Dim LOTNAME As String
'        LOTNAME = .Range("C14").Value

'総括報告書Word版は、LOTNO必要

        '進捗リストLOT番号追記列
        Dim UpdateLotNum As String
        UpdateLotNum = .Range("C13").Value
        
        '進捗リストLOT番号接頭字
        Dim LOTNAME As String
        LOTNAME = .Range("C14").Value
        
        
        '進捗リストLOT番号接尾字
        Dim LOTLASTNANE As String
        LOTLASTNANE = .Range("C15").Value


      
    End With
    
    '進捗リスト、実績報告書設定読み込み開始行
    Dim rec As Long
    rec = 18
    
    '進捗リストの列情報を格納用の配列変数
    Dim adressAry() As String
    '進捗リストの項目設定 「列」の行、列番号、画面シート名を指定
    adressAry() = getInitArray(rec, 4, "総括報告書作成")


    '実績報告書の転記先のセル番地格納用
    Dim MapingAry() As String
    '実績報告のセル番地設定 「番地」、行列番号、画面シート名を指定
    MapingAry() = getInitArray(rec, 7, "総括報告書作成")


    '進捗リストから取得した、転記用データ格納用の配列変数
    Dim dataAry() As String
    
    '設定値の読み込み実行
    'dataAry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry(), ListStart, "", "", "", ReadRecord)
    dataAry() = InputArray(ShintyokuFile, ShintyokuSheet, adressAry(), ListStart, UpdateLotNum, LOTNAME, LOTLASTNANE, ReadRecord)
     
    
    '実施報告書への転記処理
    Call PosToFileWrd(OutPutFolder, Jisitemplate, dataAry(), -1, True)

MsgBox "処理が終了しました"

End Sub

Public Function PosToFileWrd(Filedir As String, tmpfile As String, data() As String, Item As Long, flg As Boolean) As Boolean

    '2020.4.15 作成
    
    '進捗リストの配列データを指定の転記先（セル番地）に入れる
    '指定のファイル名で保存
    
    '
    '引数　：作成した実績報告書の保存先
    '　　　　実績報告テンプレートファイル

    
    '　　　　進捗リスト転記用配列データ
    
    '　　　　「作成」の配列構造デフォルト
    '　　　　　DATA(ONAME,届出受理年月日,先進医療の費用 (届出時),人件費,技術名,ファイル名,LOT番号)
    '　　　　　最後と最後から1つ前の要素は固定、他は設定により変更可
    
    '　　　　「更新」の配列構造デフォルト
    '　　　　　DATA(別添1：コード番号,ファイル名)
    '　　　　　最後の要素は固定、他は設定により変更可
    
    
    '        進捗リスト転記用配列データの添え字。
    ' 　　　　　　-1の場合、カウンター変数は、カウントアップ。
    '           　-1では無い場合は、固定値
    
    '        フラグ
    '　　　　　　　「作成」の場合、True(=-1)
    '              「更新」の場合、False(=0)

    '
    '戻り値：作業結果
    
  
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
'    Dim TargetBook As Workbook
    
    For i = 0 To RecordNo
      
        '実績報告ファイルテンプレートをオープン
        '「作成」の場合は、テンプレート
        '「更新」の場合は、受領したファイル


       ' 最後の要素に入力された値が、Nullもしくは、"-"の場合は、
       '処理を飛ばす｡
        
        If Len(data(i, UBound(data, 2) + CInt(flg))) <= 1 Then
             'ファイルが存在しない場合は、次の処理へ
             GoTo Next_Step
            
        End If


'        Set TargetBook = Workbooks.Open(WorkBookPath(ThisWorkbook.Path, AcptFileDir, TmpFile), False, 0)

        Dim objWord As Word.Application
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = True
        
        Dim TargetDoc As Word.Document
        
        Set TargetDoc = objWord.Documents.Open(WorkBookPath(ThisWorkbook.Path, "", tmpfile), ReadOnly:=False)
    
          
'        '進捗リストファイル内部の作業に遷移
'         With TargetBook.Sheets(SheetName)

         
              'item=-1の場合、カウンター変数は、カウントアップ
              If Item = -1 Then
              '
              Else
              'item<>-1の場合、カウンター変数は、itemの固定値
                 i = Item
              End If
              
'
'              For j = 0 To UBound(Mapping)
'
'               Debug.Print Mapping(j); ":"; data(i, j)
''                .Range(Mapping(j)).Value = data(i, j)
'
'              Next j
              
              
               TargetDoc.Tables(1).Cell(2, 2).Range.Text = data(i, j)
              
              
              
'              '一番最後の要素に入力したLOT番号をフッターに入力
'              UBound(data, 2)
'              j = j + 1

              If flg Then
'                  .PageSetup.RightFooter = data(i, j)

                    
                      Dim Sec As Object
                      Dim ftr As Object
                    
                      For Each Sec In TargetDoc.Sections
                      For Each ftr In Sec.Footers
                        With ftr.Range
                          .Text = data(i, UBound(data, 2))
                          .Paragraphs.Alignment = wdAlignParagraphRight
                        End With
                      Next ftr
                    Next Sec
                    
                    
                    
'                  Debug.Print .PageSetup.RightFooter
              Else
                  
''                  .PageSetup.RightFooter = Null
'                    TargetDoc.Sections.Footers.Range.Text = data(i, j)
              
              End If
              
'         End With
         
         'Debug.Print CInt(flg)
         
         '進捗リスト転記配列、
         '「作成」の場合、最後から一つ前の要素　全要素数＋True(=-1)
         '「更新」の場合、最後の要素
         '　に入力された値をファイル名として保存する｡
         
         '2020/06/10
         '「総括報告書作成」を実行しましたが、セル番地：D20で指定したファイル名が反映されませんでした。
         'D19に指定した医療機関名でファイル名がつけられているようにおもいます｡
         'ご確認をお願い致します｡
         '
'         Debug.Print UBound(data, 2) + CInt(flg)
'        TRUEは「-1」なので　1-1 = 0 →　data(i,0)を参照していた　、D20を見ていたつもりが、D19に


         
'         TargetDoc.SaveAs _
'                    filename:=WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg)))


'
'         TargetDoc.SaveAs _
'                    filename:=WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg)))
                    
                   
'        TargetDoc.SaveAs _
'                    FILENAME:=WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2)) & ".doc")
                   
                   
        Dim strFileName As String
         
        With TargetDoc
        If InStr(data(i, UBound(data, 2) + CInt(flg)), wExt) > 0 Then


            strFileName = WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg)))
         
         Else
       
            strFileName = WorkBookPath(ThisWorkbook.Path, Filedir, data(i, UBound(data, 2) + CInt(flg)) & wExt)
       
         End If
         

            .SaveAs2 _
                    filename:=strFileName, _
                    FileFormat:=wdFormatDocument
                    
            .Close
            
         
         End With
                   
    
'
        Set TargetDoc = Nothing
        
        objWord.Quit

            
        Set objWord = Nothing


         DoEvents


Next_Step:

     Next i
'
  
End Function


