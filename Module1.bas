Attribute VB_Name = "Module1"
Option Explicit
'Connection�I�u�W�F�N�g
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
'�ڑ�������
Const dtSrc As String = "�iODBC�f�[�^�\�[�X���j"
Const myDb As String = "�i�f�[�^�x�[�X���j"
Const userId As String = "�i���[�U�h�c�j"
Const passWd As String = "�i�p�X���[�h�j"
Const conStr As String = "Driver={MySQL ODBC 8.0 Unicode Driver};" _
  & "Server=localhost; Data Source=" & dtSrc _
  & ";Database=" & myDb & "; Uid=" & userId & "; Pwd=" & passWd

Sub �e�[�u���ꗗ�쐬()
  On Error GoTo errHandler
  Dim row As Long: row = 3
    
  '�f�[�^�x�[�X�ɐڑ�����
  cn.ConnectionString = conStr
  cn.Open
    
  '�e�[�u���ꗗ���擾����
  Dim strSQL As String
  strSQL = "SELECT TABLE_NAME,TABLE_ROWS,TABLE_COMMENT FROM information_schema.TABLES WHERE TABLE_SCHEMA = '" & myDb & "'"
  If Sheets("�e�[�u���ꗗ").Cells(1, 2).Value <> "" Then
    strSQL = strSQL & " AND TABLE_NAME LIKE '" & Sheets("�e�[�u���ꗗ").Cells(1, 2).Value & "%'"
  End If

  Set rs = cn.Execute(strSQL)
  'RecordSet�̏I���܂�
  Do While rs.EOF = False
    If Cells(row, 2).Value <> rs(0).Value Then
      Cells(row, 1).Value = ""
      Cells(row, 2).Value = rs(0).Value
      Cells(row, 5).Value = ""
      With Cells(row, 2)
        .Hyperlinks.Delete
        .Font.Name = "Meiryo UI"
        .Font.Size = 11
        .Font.Underline = False
        .Font.ColorIndex = 1
      End With
    End If
    Cells(row, 4).Value = rs(1).Value
    Cells(row, 3).Value = rs(2).Value
    '���̃��R�[�h
    rs.MoveNext
    row = row + 1
  Loop
  '�����\�����N���A
  If Cells(row, 2) <> "" Then
    Range(Cells(row, 1), Cells(row, 4)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Clear
  End If
  Range("A2").Select
  '�ڑ���ؒf����
  cn.Close
  MsgBox "�ꗗ�쐬�I��"

procContinue:
  Set cn = Nothing
  Exit Sub

errHandler:
  '�G���[�����������ꍇ�ɃG���[���b�Z�[�W��\������
  MsgBox Err.Description
  Resume procContinue
End Sub

Sub �f�[�^�擾()
  Dim cnt As Long
  Dim row As Long: row = 3
  Dim TARGET As String
    
  Do While Cells(row, 2).Value <> ""
    'SIGN�`�F�b�N(�󔒈ȊO���Ώ�)
    If Cells(row, 1).Value <> "" Then
      TARGET = Cells(row, 2).Value
      '�V�[�g�`�F�b�N
      SheetSelect (TARGET)
      '�f�[�^�擾
      SelectTable (row)
      Sheets("�e�[�u���ꗗ").Select
      ActiveSheet.Hyperlinks.Add Anchor:=Cells(row, 2), _
        Address:="", SubAddress:=TARGET & "!A2"
      With Cells(row, 2)
        .Font.Name = "Meiryo UI"
        .Font.Size = 11
      End With
      cnt = cnt + 1
      Cells(row, 1).Select
    End If
    row = row + 1
  Loop
  If cnt = 0 Then
    MsgBox "A��ɃT�C��������܂���I"
  Else
    MsgBox "�f�[�^�擾�I��"
  End If
End Sub

Private Sub SelectTable(row As Long)
  On Error GoTo errHandler
  Dim col As Long
  Dim TARGET As String
    
  TARGET = Sheets("�e�[�u���ꗗ").Cells(row, 2).Value
  Cells.Select
  Selection.NumberFormatLocal = "@"
  Selection.Font.Name = "���S�V�b�N"
    
  '�f�[�^�x�[�X�ɐڑ�����
  cn.ConnectionString = conStr
  cn.Open
    
  '�J�������e���擾����
  Dim strSQL As String
  strSQL = "SELECT COLUMN_COMMENT,COLUMN_NAME,COLUMN_TYPE" & _
            " FROM information_schema.COLUMNS" & _
           " WHERE TABLE_SCHEMA = '" & myDb & "' AND TABLE_NAME = '" & TARGET & "'"
  Set rs = cn.Execute(strSQL)
  Do Until rs.EOF
    Cells(3, col + 1).Value = rs(0).Value
    Cells(4, col + 1).Value = rs(1).Value
    Cells(5, col + 1).Value = rs(2).Value
    rs.MoveNext
    col = col + 1
  Loop

  '�s���e���擾����
  strSQL = "SELECT * FROM " & TARGET
  If Sheets("�e�[�u���ꗗ").Cells(row, 5).Value <> "" Then
    strSQL = strSQL & " WHERE " & Sheets("�e�[�u���ꗗ").Cells(row, 5).Value
  End If
  Set rs = cn.Execute(strSQL)
  Cells(6, 1).CopyFromRecordset rs
   
  '�ڑ���ؒf����
  cn.Close

  '�r���`��
  Range("A4").Select
  Range(Selection, Selection.End(xlToRight)).Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Borders.LineStyle = xlContinuous
  Selection.EntireColumn.AutoFit

  '���o���r���`��
  Range("A5").Select
  Range(Selection, Selection.End(xlToRight)).Select
  Range(Selection, Selection.End(xlUp)).Select
  Selection.Borders.LineStyle = xlContinuous
  Selection.Interior.Color = 65535
  Range("A1").Select

  If Sheets("�e�[�u���ꗗ").Cells(row, 3).Value = "" Then
    Cells(1, 1).Value = TARGET
  Else
    Cells(1, 1).Value = TARGET & "(" & Sheets("�e�[�u���ꗗ").Cells(row, 3).Value & ")"
  End If
  Cells(2, 1).Value = Format(Now, "YYYY/MM/DD HH:MM")
  Cells(2, 3).Value = Sheets("�e�[�u���ꗗ").Cells(row, 5).Value

procContinue:
  Set cn = Nothing
  Exit Sub

errHandler:
  '�G���[�����������ꍇ�ɃG���[���b�Z�[�W��\������
  MsgBox Err.Description
  Resume procContinue
End Sub

'�V�[�g�L�����m�F���A�����ꍇ�͒ǉ�����
Private Sub SheetSelect(strSheetName As String)
  On Error GoTo NotExists

  Sheets(strSheetName).Select
  Cells.Clear
  Exit Sub

NotExists:
  Worksheets().Add After:=Worksheets(Worksheets.Count)   ' �����ɒǉ�
  ActiveSheet.Name = strSheetName
End Sub
