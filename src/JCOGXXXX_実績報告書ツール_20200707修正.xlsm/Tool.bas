Attribute VB_Name = "Tool"
Option Explicit

'���ѕ񍐏��̊g���q�@97-2003�̏ꍇ
Public Const Ext As String = ".xls"
Public Const xlsFmt As Integer = 56


'���ѕ񍐏��̊g���q�@2007�ȍ~�̏ꍇ
'Public Const Ext As String = ".xlsx"
'Public Const xlsFmt As Long = 51

'�����񍐏��̊g���q
Public Const wExt As String = ""
Public Const wrdFmt As Long = 0#


'�t�@�C���R�s�[�c�[��
'�萔:�t�@�C����
Public Const filename As String = "�t�@�C����"
'�萔:C
Public Const RETUC As String = "C"

'�萔:�t�H���_�[��
Public Const FOLDERNAME As String = "�t�H���_�[��"
'�萔:C
Public Const RETUB As String = "B"


Public Function WorkBookPath(FilePath As String, Filedir As String, filename As String) As String

    '2020.4.15 �쐬
    '�l�b�g���[�N��̃t���p�X�̃A�h���X���쐬����Ɓu\�v�ɂ���ĕ����񂪍쐬����Ȃ����ߋ��ʊ֐�
    '�t���p�X�̃t�@�C�������쐬
    
    '
    '�����@:�c�[���t�@�C���p�X�i�c�[���̏ꏊ�A��ʁj
    '�@�@�@:�c�[���t�@�C�����̃f�B���N�g��
    '      :�c�[���t�@�C�����@�܂��́A�쐬����񍐏��t�@�C��
    
    
    '�߂�l�F�t���p�X�t�@�C����
    
    
    If Len(Filedir) = 0 Then
      
      WorkBookPath = Trim(FilePath & "\" & filename)
    
    ElseIf Len(filename) = 0 Then
    
      WorkBookPath = Trim(FilePath & "\" & Filedir)
    
    Else
      
      WorkBookPath = Trim(FilePath & "\" & Filedir & "\" & filename)
    
    End If
  
End Function


Public Function FindDirectory(FOLDERNAME As String) As Boolean
  
    '2020.4.15 �쐬
    '�����̃t�H���_�[������΁A���b�Z�[�W��Ԃ�
    '
    '�����@�F�쐬����t�H���_��
    '�߂�l�F
    '�@�@�@�@�쐬�����ꍇ�A1
    '        �����̃t�@�C�������݂����ꍇ�́A2��Ԃ�
    
    
    Dim Filedir As String
    Filedir = Dir(FOLDERNAME, vbDirectory)
    
    If Len(Filedir) = 0 Then    '�����̃t�H���_�[���Ȃ��ꍇ
'
'        MkDir FolderName
    
        FindDirectory = False
     
        Exit Function
    
    Else                        '�����̃t�H���_�[������ꍇ
    
        FindDirectory = True
    
    End If
  
  
End Function


Public Function FindFile(filename As String) As Boolean

    '2020.4.15 �쐬
    '�t�@�C�������݂��Ă��邩�m�F����
    '�t�@�C�������݂��Ă��Ȃ���΁A���b�Z�[�W��\������
    '
    '�����@�F�t�@�C����
    '�߂�l�F
    '�@�@�@�@���݂��Ă���ꍇ�ATrue
    '        ���݂��Ă��Ȃ��ꍇ�́AFalse
    
    
    Dim Filedir As String
    Filedir = Dir(filename)
    
    If Len(Filedir) > 0 Then
    
        FindFile = True

    Else
    
        FindFile = False
    
    End If
  
End Function

Public Function LOTNUM(LotSt As String, Num As String) As String

    '2020.4.22 �쐬
    'LOT�ԍ��쐬�@LOT�ԍ��ړ����{�s�ԍ��i4���j�{���t
    '�����@�FLot�ԍ��ړ���
    '�@�@�@�@�s�ԍ�
    '�߂�l�FLOT�ԍ�
    
    
    'LotNum = LotSt & "_" & Format(Num, "0000") & "_" & Format(Date, "YYYY") & Format(Date, "MM") & Format(Date, "DD")

    LOTNUM = LotSt & "_" & Format(Num, "0000")

End Function

Public Function getLotNum(filename As String) As String

  '2020.4.22 �쐬
  '��̐�t�H���_���̃t�@�C����LOT�ԍ��̔z����쐬
  
  '�����@�F��̐�t�H���_��
  '�߂�l�F��̐�t�H���_���̃t�@�C����LOT�ԍ��̔z��

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
'�₢���킹���ʂ�i�����X�g��LOTNUM�ׂ̗̗�ɒǋL���Ă���

    Dim rng As Range

    '�i�����X�g�t�@�C���I�[�v��
    Dim TargetBook As Workbook
    Set TargetBook = Workbooks.Open(WorkBookPath(ThisWorkbook.Path, "", filename), False, 0)
    
    '�i�����X�g�t�@�C�������̍�ƂɑJ��
    
    With TargetBook.Sheets(SheetName)
    
        '�i�����X�gLOT�ԍ��ǋL��
        
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

    
    '�i�����X�g�����B�������[�J��
    TargetBook.Close SaveChanges:=True
    Set TargetBook = Nothing
    
    DoEvents
    

End Sub
