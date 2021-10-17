Option Explicit
'Connection�I�u�W�F�N�g
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
'�ڑ�������
Const schemaId As String = "dbo"
Const conStr As String = "Provider=MSOLEDBSQL; Server=(local); Database=Test; Integrated Security=SSPI;"

Sub �e�[�u���ꗗ�쐬()
  On Error GoTo errHandler
  Dim row As Long: row = 3
    
  '�f�[�^�x�[�X�ɐڑ�����
  cn.ConnectionString = conStr
  cn.Open
    
  '�e�[�u���ꗗ���擾����
  Dim strSQL As String
  strSQL = "SELECT t.name,i.rows,CAST(ep.value AS NVARCHAR(50))"
  strSQL = strSQL & " FROM sys.tables AS t,sys.extended_properties AS ep,sys.sysindexes AS i,sys.schemas AS s"
  strSQL = strSQL & " WHERE t.object_id = ep.major_id AND ep.minor_id = 0 AND t.object_id = i.id AND i.indid < 2"
  strSQL = strSQL & " AND t.schema_id = s.schema_id AND s.name = '" & schemaId & "' ORDER BY t.name"
  
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
  strSQL = "SELECT CAST(ep.value AS NVARCHAR(50)),c.name,type_name(user_type_id),max_length" & _
           " FROM sys.objects AS t" & _
           " INNER JOIN sys.columns AS c ON t.object_id = c.object_id" & _
           " LEFT JOIN sys.extended_properties AS ep ON t.object_id = ep.major_id" & _
           " AND c.column_id = ep.minor_id AND ep.name = 'MS_Description'" & _
           " WHERE t.type = 'U' AND t.name='" & TARGET & "' ORDER BY c.column_id"

  Set rs = cn.Execute(strSQL)
  Do Until rs.EOF
    Cells(3, col + 1).Value = rs(0).Value
    Cells(4, col + 1).Value = rs(1).Value
    Cells(5, col + 1).Value = rs(2).Value & "(" & rs(3).Value & ")"
    rs.MoveNext
    col = col + 1
  Loop

  '�s���e���擾����
  strSQL = "SELECT * FROM " & schemaId & "." & TARGET
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

'CSV�ɏo�͂���
Sub OutputCSV()

  Dim cnt As Long
  Dim TARGET As String
  Dim csvFile As String
  Dim strLine As String
  Dim R As Long
  Dim C As Long
  Dim MaxCol As Long
  Dim ST As ADODB.Stream
    
  Dim row As Long: row = 3
  Do While Cells(row, 2).Value <> ""
    'SIGN�`�F�b�N(�󔒈ȊO���Ώ�)
    If Cells(row, 1).Value <> "" Then
      TARGET = Cells(row, 2).Value
      '�V�[�g�`�F�b�N
      On Error GoTo NotExists
      Sheets(TARGET).Select
      On Error GoTo 0
      '�t�@�C�����ݒ�
      csvFile = ActiveWorkbook.Path & "\" & TARGET & ".csv"
      C = 1
      Do While Cells(4, C).Value <> ""
        MaxCol = C
        C = C + 1
      Loop
      'ADODB.Stream�I�u�W�F�N�g�𐶐�
      Dim adoSt As Object
      Set adoSt = CreateObject("ADODB.Stream")
      With adoSt
        .Charset = "UTF-8"
        .LineSeparator = adLF
        .Open
        R = 6
        Do While Cells(R, 1).Value <> ""
          strLine = Cells(R, 1).Value
          For C = 2 To MaxCol
            strLine = strLine & "," & Cells(R, C).Value
          Next C
          .WriteText strLine, adWriteLine
          R = R + 1
        Loop
        .Position = 0          '�X�g���[���̈ʒu��0�ɂ���
        .Type = adTypeBinary   '�f�[�^�̎�ނ��o�C�i���f�[�^�ɕύX
        .Position = 3          '�X�g���[���̈ʒu��3�ɂ���

        Dim byteData() As Byte '�ꎞ�i�[�p
        byteData = .Read       '�X�g���[���̓��e���ꎞ�i�[�p�ϐ��ɕۑ�
        .Close                 '��U�X�g���[�������i���Z�b�g�j

        .Open                  '�X�g���[�����J��
        .Write byteData        '�X�g���[���Ɉꎞ�i�[�����f�[�^�𗬂�����
        .SaveToFile csvFile, adSaveCreateOverWrite
        .Close
      End With
      cnt = cnt + 1
      Sheets("�e�[�u���ꗗ").Select
    End If
    row = row + 1
  Loop
  If cnt = 0 Then
    MsgBox "A��ɃT�C��������܂���I"
  Else
    MsgBox "CSV�o�͏I��"
  End If
  Exit Sub

NotExists:
  MsgBox "�f�[�^�V�[�g������܂���I"
End Sub
