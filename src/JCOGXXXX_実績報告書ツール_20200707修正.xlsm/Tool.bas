Attribute VB_Name = "Tool"
Option Explicit

'実績報告書の拡張子　97-2003の場合
Public Const Ext As String = ".xls"
Public Const xlsFmt As Integer = 56


'実績報告書の拡張子　2007以降の場合
'Public Const Ext As String = ".xlsx"
'Public Const xlsFmt As Long = 51

'総括報告書の拡張子
Public Const wExt As String = ""
Public Const wrdFmt As Long = 0#


'ファイルコピーツール
'定数:ファイル名
Public Const filename As String = "ファイル名"
'定数:C
Public Const RETUC As String = "C"

'定数:フォルダー名
Public Const FOLDERNAME As String = "フォルダー名"
'定数:C
Public Const RETUB As String = "B"


Public Function WorkBookPath(FilePath As String, Filedir As String, filename As String) As String

    '2020.4.15 作成
    'ネットワーク上のフルパスのアドレスを作成すると「\」によって文字列が作成されないため共通関数
    'フルパスのファイル名を作成
    
    '
    '引数　:ツールファイルパス（ツールの場所、上位）
    '　　　:ツールファイル下のディレクトリ
    '      :ツールファイル名　または、作成する報告書ファイル
    
    
    '戻り値：フルパスファイル名
    
    
    If Len(Filedir) = 0 Then
      
      WorkBookPath = Trim(FilePath & "\" & filename)
    
    ElseIf Len(filename) = 0 Then
    
      WorkBookPath = Trim(FilePath & "\" & Filedir)
    
    Else
      
      WorkBookPath = Trim(FilePath & "\" & Filedir & "\" & filename)
    
    End If
  
End Function


Public Function FindDirectory(FOLDERNAME As String) As Boolean
  
    '2020.4.15 作成
    '同名のフォルダーがあれば、メッセージを返す
    '
    '引数　：作成するフォルダ名
    '戻り値：
    '　　　　作成した場合、1
    '        同名のファイルが存在した場合は、2を返す
    
    
    Dim Filedir As String
    Filedir = Dir(FOLDERNAME, vbDirectory)
    
    If Len(Filedir) = 0 Then    '同名のフォルダーがない場合
'
'        MkDir FolderName
    
        FindDirectory = False
     
        Exit Function
    
    Else                        '同名のフォルダーがある場合
    
        FindDirectory = True
    
    End If
  
  
End Function


Public Function FindFile(filename As String) As Boolean

    '2020.4.15 作成
    'ファイルが存在しているか確認する
    'ファイルが存在していなければ、メッセージを表示する
    '
    '引数　：ファイル名
    '戻り値：
    '　　　　存在している場合、True
    '        存在していない場合は、False
    
    
    Dim Filedir As String
    Filedir = Dir(filename)
    
    If Len(Filedir) > 0 Then
    
        FindFile = True

    Else
    
        FindFile = False
    
    End If
  
End Function

Public Function LOTNUM(LotSt As String, Num As String) As String

    '2020.4.22 作成
    'LOT番号作成　LOT番号接頭字＋行番号（4桁）＋日付
    '引数　：Lot番号接頭字
    '　　　　行番号
    '戻り値：LOT番号
    
    
    'LotNum = LotSt & "_" & Format(Num, "0000") & "_" & Format(Date, "YYYY") & Format(Date, "MM") & Format(Date, "DD")

    LOTNUM = LotSt & "_" & Format(Num, "0000")

End Function

Public Function getLotNum(filename As String) As String

  '2020.4.22 作成
  '受領先フォルダ内のファイルとLOT番号の配列を作成
  
  '引数　：受領先フォルダ名
  '戻り値：受領先フォルダ内のファイルとLOT番号の配列

    Dim TargetBook As Workbook
    Set TargetBook = Workbooks.Open(filename, False, 0)

    With TargetBook.Sheets(1)
        getLotNum = .PageSetup.RightFooter
    End With

    TargetBook.Close SaveChanges:=False
    Set TargetBook = Nothing
    
    DoEvents

End Function



Public Sub InPutQuery(filename As String, SheetName As String, LOTNUM As String, Query As String, FindCol As String, ofset As Long)

'2020.5.11
'問い合わせ結果を進捗リストのLOTNUMの隣の列に追記していく

    Dim rng As Range

    '進捗リストファイルオープン
    Dim TargetBook As Workbook
    Set TargetBook = Workbooks.Open(WorkBookPath(ThisWorkbook.Path, "", filename), False, 0)
    
    '進捗リストファイル内部の作業に遷移
    
    With TargetBook.Sheets(SheetName)
    
        '進捗リストLOT番号追記列
        
        With .Columns(FindCol)
        
            Set rng = .Find(LOTNUM)
            ', xlValues, xlWhole, xlByColumns, xlNext
            If rng Is Nothing Then
            
                Exit Sub
            
            End If
            
'            .Cells(rng.Row, rng.Column).Value = Query
            Range(rng.Address).Offset(0, ofset).Value = Query
            
'            Debug.Print rng.Address()
            
        End With
        

        
    End With

    
    '進捗リストを閉じる。メモリー開放
    TargetBook.Close SaveChanges:=True
    Set TargetBook = Nothing
    
    DoEvents
    

End Sub
