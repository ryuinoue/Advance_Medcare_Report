Attribute VB_Name = "Tool3"
Option Explicit

Public Function WordTempMail(FldName As String, Filenme As String) As String()
  Const wdPrintView = 3
  Const wdTextRectangle = 0

'  Dim wd_app As Object  ' Word.Application
  Dim win As Object      ' Word.Window
  Dim pg As Object        ' Word.Page
  Dim rc As Object        ' Word.Rectangle
  Dim ln As Object        ' Word.Line
  Dim r As Long          ' 文字列を書き出すExcelの行番号

'  On Error GoTo ERR_HNDL
'
'  Set wd_app = GetObject(Class:="Word.Application")

Dim tmpfile As String
tmpfile = Filenme

'
' Dim wd_app As Workbook
 
 Dim arry(100) As String
 Dim i As Long
 
' Debug.Print ThisWorkbook.Path & "\" & tmpfile
'Application.EnableEvents = False

 
         Dim objWord As Word.Application
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = True
        
        Dim TargetDoc As Word.Document
        
        
        Set TargetDoc = objWord.Documents.Open(WorkBookPath(ThisWorkbook.Path, FldName, tmpfile), ReadOnly:=False)
 
 
  Set win = TargetDoc.ActiveWindow
'  win.View.Type = wdPrintView

'  With ThisWorkbook.Sheets("WtoE")
  r = 1
  For Each pg In win.ActivePane.Pages
    For Each rc In pg.Rectangles
      For Each ln In rc.Lines
          If rc.RectangleType = wdTextRectangle Then

'            .Cells(r, "A").Value = Replace(ln.Range.Text, "＿", Chr(-32448))
'            arry(i) = Replace(ln.Range.Text, "＿", Chr(-32448))
            arry(i) = ln.Range.Text


            r = r + 1
            i = i + 1
          End If
      Next ln
    Next rc
  Next pg

                 TargetDoc.Close
''
        Set TargetDoc = Nothing

         
         

r = 0
'Do Until i + 1 < r
'            .Cells(r + 1, "A").Value = arry(r)
''            r = r + 1
'Loop
  
'  End With

'  GoTo END_TASK

'ERR_HNDL:
'  MsgBox "エラーが発生したためマクロを終了します。"
'  Err.Clear
'  GoTo END_TASK
'
'END_TASK:
'  Set wd_app = Nothing
'
'Application.EnableEvents = True
'Application.Wait [Now()] + 1000 / 86400000

'
'Application.Wait [Now()] + 10000 / 86400000

        '一番最後タイミング
        
'        objWord.Quit

'処理が遅くなるので､開発環境で実行しない
        
If False Then
            objWord.Quit
End If


        Set objWord = Nothing

'         DoEvents
'MsgBox "終了"

'Exit Sub


  WordTempMail = arry
  
End Function


Public Function outPutText(s As String, filename As String)

'Dim s As String
Dim n As Integer

n = FreeFile()
Open ThisWorkbook.Path & "\" & filename & ".txt" For Append As #n

's = "Hello, world!5"
'Debug.Print s ' write to immediate
Print #n, s ' write to file
'Print #n, "\n" ' write to file

Close #n


End Function
