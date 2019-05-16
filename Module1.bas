Attribute VB_Name = "Module1"
Option Explicit
'Connectionオブジェクト
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
'接続文字列
Const dtSrc As String = "（ODBCデータソース名）"
Const myDb As String = "（データベース名）"
Const userId As String = "（ユーザＩＤ）"
Const passWd As String = "（パスワード）"
Const conStr As String = "Driver={MySQL ODBC 8.0 Unicode Driver};" _
  & "Server=localhost; Data Source=" & dtSrc _
  & ";Database=" & myDb & "; Uid=" & userId & "; Pwd=" & passWd

Sub テーブル一覧作成()
  On Error GoTo errHandler
  Dim row As Long: row = 3
    
  'データベースに接続する
  cn.ConnectionString = conStr
  cn.Open
    
  'テーブル一覧を取得する
  Dim strSQL As String
  strSQL = "SELECT TABLE_NAME,TABLE_ROWS,TABLE_COMMENT FROM information_schema.TABLES WHERE TABLE_SCHEMA = '" & myDb & "'"
  If Sheets("テーブル一覧").Cells(1, 2).Value <> "" Then
    strSQL = strSQL & " AND TABLE_NAME LIKE '" & Sheets("テーブル一覧").Cells(1, 2).Value & "%'"
  End If

  Set rs = cn.Execute(strSQL)
  'RecordSetの終了まで
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
    '次のレコード
    rs.MoveNext
    row = row + 1
  Loop
  '既存表示をクリア
  If Cells(row, 2) <> "" Then
    Range(Cells(row, 1), Cells(row, 4)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Clear
  End If
  Range("A2").Select
  '接続を切断する
  cn.Close
  MsgBox "一覧作成終了"

procContinue:
  Set cn = Nothing
  Exit Sub

errHandler:
  'エラーが発生した場合にエラーメッセージを表示する
  MsgBox Err.Description
  Resume procContinue
End Sub

Sub データ取得()
  Dim cnt As Long
  Dim row As Long: row = 3
  Dim TARGET As String
    
  Do While Cells(row, 2).Value <> ""
    'SIGNチェック(空白以外が対象)
    If Cells(row, 1).Value <> "" Then
      TARGET = Cells(row, 2).Value
      'シートチェック
      SheetSelect (TARGET)
      'データ取得
      SelectTable (row)
      Sheets("テーブル一覧").Select
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
    MsgBox "A列にサインがありません！"
  Else
    MsgBox "データ取得終了"
  End If
End Sub

Private Sub SelectTable(row As Long)
  On Error GoTo errHandler
  Dim col As Long
  Dim TARGET As String
    
  TARGET = Sheets("テーブル一覧").Cells(row, 2).Value
  Cells.Select
  Selection.NumberFormatLocal = "@"
  Selection.Font.Name = "游ゴシック"
    
  'データベースに接続する
  cn.ConnectionString = conStr
  cn.Open
    
  'カラム内容を取得する
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

  '行内容を取得する
  strSQL = "SELECT * FROM " & TARGET
  If Sheets("テーブル一覧").Cells(row, 5).Value <> "" Then
    strSQL = strSQL & " WHERE " & Sheets("テーブル一覧").Cells(row, 5).Value
  End If
  Set rs = cn.Execute(strSQL)
  Cells(6, 1).CopyFromRecordset rs
   
  '接続を切断する
  cn.Close

  '罫線描画
  Range("A4").Select
  Range(Selection, Selection.End(xlToRight)).Select
  Range(Selection, Selection.End(xlDown)).Select
  Selection.Borders.LineStyle = xlContinuous
  Selection.EntireColumn.AutoFit

  '見出し罫線描画
  Range("A5").Select
  Range(Selection, Selection.End(xlToRight)).Select
  Range(Selection, Selection.End(xlUp)).Select
  Selection.Borders.LineStyle = xlContinuous
  Selection.Interior.Color = 65535
  Range("A1").Select

  If Sheets("テーブル一覧").Cells(row, 3).Value = "" Then
    Cells(1, 1).Value = TARGET
  Else
    Cells(1, 1).Value = TARGET & "(" & Sheets("テーブル一覧").Cells(row, 3).Value & ")"
  End If
  Cells(2, 1).Value = Format(Now, "YYYY/MM/DD HH:MM")
  Cells(2, 3).Value = Sheets("テーブル一覧").Cells(row, 5).Value

procContinue:
  Set cn = Nothing
  Exit Sub

errHandler:
  'エラーが発生した場合にエラーメッセージを表示する
  MsgBox Err.Description
  Resume procContinue
End Sub

'シート有無を確認し、無い場合は追加する
Private Sub SheetSelect(strSheetName As String)
  On Error GoTo NotExists

  Sheets(strSheetName).Select
  Cells.Clear
  Exit Sub

NotExists:
  Worksheets().Add After:=Worksheets(Worksheets.Count)   ' 末尾に追加
  ActiveSheet.Name = strSheetName
End Sub
