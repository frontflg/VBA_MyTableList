Option Explicit
' ツール>参照設定>Microsoft ActiveX Data Objects X.X Library
   Dim OraCn  As ADODB.Connection
   Dim OraRs  As ADODB.Recordset
   Dim vSQL   As Variant
Public stPass As String
Private Function WaitForIE(ByVal MinSec As Integer, ByVal MaxSec As Integer)
Dim WaitCnt As Integer
   For i = 0 To MaxSec * 1000
      Sleep 1
      If i = Int(i / 1000) * 1000 Then
         Application.StatusBar = "WAIT．．．．(" & i & "回目)"
         DoEvents
      End If
   Next i
   Application.StatusBar = False
End Function
' ADO : ORACLE DB 接続
Sub DB_CONNECT()
   On Err GoTo Err_Han

   Sheets("TABLE_LIST").Select
' パスワードの設定
   If Cells(4, 3).Value <> "" Then
      stPass = Cells(4, 3).Value
   Else
      stPass = InputBox("USER:" & stUsr & "のPassWord？")
      If stPass = "" Then
         MsgBox "パスワードの入力がありませんでした!"
         Exit Sub
      End If
      Cells(4, 3).Value = stPass
   End If
' ORACLE接続  HOST:C1 , SID:C2 , USER:C3 , PASSWORD:C4
   Set OraCn = CreateObject("ADODB.Connection")
   OraCn.Open "Provider=MSDAORA;Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=" & _
              "(PROTOCOL=TCP)(HOST=" & Cells(1, 3).Value & ")(PORT = 1521)))" & _
              "(CONNECT_DATA=(SID=" & Cells(2, 3).Value & ")));", Cells(3, 3).Value, stPass
   Exit Sub
Err_Han:
   MsgBox (Err.Description)
   OraCn.Close
   Set OraCn = Nothing
End Sub
' ADO : テーブル一覧作成
Sub TABLE_LISTING()
   On Err GoTo Err_Han
   Dim iRow   As Integer
   Dim stFilt As String

   Call DB_CONNECT
   If stPass = "" Then Exit Sub

   If Sheets("TABLE_LIST").Cells(9, 1) = 1 Then
      If MsgBox("既にあるマスター一覧を作り直しますか？", vbYesNo + vbExclamation, "一覧クリアの確認") = vbNo Then
         GoTo Kensu_Get
      End If
   End If
' テーブル情報取得
   Cells(6, 3) = UCase(Cells(6, 3))
   vSQL = "SELECT TABLE_NAME,COMMENTS FROM USER_TAB_COMMENTS"
   If Cells(6, 3) <> "" Then
      vSQL = vSQL & " WHERE TABLE_TYPE = 'TABLE' AND TABLE_NAME LIKE '" & Cells(6, 3)
      vSQL = vSQL & "%'"
   End If
   If Cells(7, 3) <> "" And Cells(7, 4) <> "" Then
      stFilt = "('" & Replace(Cells(7, 4), ",", "','") & "')"
      If Cells(6, 3) <> "" Then
         vSQL = vSQL & " AND"
      Else
         vSQL = vSQL & " WHERE"
      End If
      vSQL = vSQL & " SUBSTR(TABLE_NAME,1," & Cells(7, 3) & ") IN " & stFilt
   End If
   vSQL = vSQL & " ORDER BY TABLE_NAME"
   Set OraRs = OraCn.Execute(vSQL)

   iRow = 9
   Do While Not OraRs.EOF
      Cells(iRow, 1) = iRow - 8
      Cells(iRow, 2) = OraRs.Fields(0).Value
      Range(Cells(iRow, 2), Cells(iRow, 3)).MergeCells = True
      Cells(iRow, 4) = OraRs.Fields(1).Value
      Range(Cells(iRow, 4), Cells(iRow, 5)).MergeCells = True
      Cells(iRow, 1).Select
      iRow = iRow + 1
      OraRs.MoveNext
   Loop
   OraRs.Close
   Set OraRs = Nothing
' 余分行クリア
   Range(Cells(iRow, 1), Cells(iRow, 6)).Select
   Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
   Do While Cells(iRow, 1) <> ""
      Range(Cells(iRow, 1), Cells(iRow, 9)).Select
      Selection.ClearContents
      Selection.Borders.LineStyle = xlNone
      iRow = iRow + 1
   Loop
Kensu_Get:
' 件数取得
   iRow = 9
   Do While Cells(iRow, 2).Value <> ""
      vSQL = "SELECT COUNT(*) FROM " & Cells(iRow, 2)
      Set OraRs = OraCn.Execute(vSQL)
      If Not OraRs.EOF Then
         Cells(iRow, 6) = OraRs.Fields(0).Value
         Cells(iRow, 6).Select
         OraRs.Close
         Set OraRs = Nothing
      Else
         Cells(iRow, 6) = 0
      End If
      Cells(iRow, 7) = ""
      Cells(iRow, 8) = ""
      iRow = iRow + 1
   Loop

' 罫線セット
   Range(Cells(9, 1), Cells(iRow - 1, 6)).Borders.LineStyle = xlContinuous

   MsgBox "テーブル一覧作成終了！"
   Range("A6").Select
Obj_Rls:
   OraCn.Close
   Set OraCn = Nothing
   Exit Sub
Err_Han:
   MsgBox (Err.Description)
   GoTo Obj_Rls
End Sub
' ADO : テーブル展開作成
Sub COLUMN_LISTING()
   On Err GoTo Err_Han

   Dim stShn  As String
   Dim iRow   As Integer
   Call DB_CONNECT
   If stPass = "" Then Exit Sub
' 検索・展開
   iRow = 9
   Do While Sheets("TABLE_LIST").Cells(iRow, 2).Value <> ""
      stShn = Sheets("TABLE_LIST").Cells(iRow, 2).Value
      Call InitWorkSheet(stShn, iRow) 'ワークシート初期化
      Call SetItemName(iRow)          '検索結果展開
      Application.ScreenUpdating = False
      Sheets("TABLE_LIST").Select
    ' ハイパーリンクセット
      Cells(iRow, 2).Select
      ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:=stShn & "!A1", TextToDisplay:=stShn
      With Selection.Font
         .Name = "HGゴシックM"
         .Underline = xlUnderlineStyleNone
         .ColorIndex = xlAutomatic
      End With
      Application.ScreenUpdating = True
      iRow = iRow + 1                 '次の行へ
   Loop

   MsgBox "テーブル展開終了！"
   Range("A6").Select
Obj_Rls:
   OraCn.Close
   Set OraCn = Nothing
   Exit Sub
Err_Han:
   MsgBox (Err.Description)
   GoTo Obj_Rls
End Sub
'ワークシートクリア
Sub InitWorkSheet(pShn As String, pRow As Integer)

   Application.ScreenUpdating = False

   If ExistSheet(pShn) Then
      Sheets(pShn).Select
      Range("A11").Select
      ActiveCell.CurrentRegion.Select
      Selection.Borders.LineStyle = xlNone
      Selection.ClearContents
   Else
      Dim NewWorkSheet As Worksheet
      Set NewWorkSheet = Worksheets.Add(after:=Worksheets(Worksheets.Count)) '-- シートが無ければ追加する
      NewWorkSheet.Name = pShn
      Cells.Select
      Selection.Font.Name = "HGゴシックM"
   End If
   Cells(1, 1) = "【TABLE-ID】"
   Cells(1, 2) = pShn
   Cells(2, 1) = "【テーブル名】"
   Cells(2, 2) = Sheets("TABLE_LIST").Cells(pRow, 4).Value
   Cells(3, 1) = "【条件】"
   Range(Cells(1, 1), Cells(3, 1)).HorizontalAlignment = xlRight
   Range(Cells(1, 1), Cells(3, 1)).Columns.AutoFit
   Application.ScreenUpdating = True
End Sub
Sub SetItemName(pRow As Integer)
   On Err GoTo Err_Han

   Dim i      As Integer

   Sheets("TABLE_LIST").Cells(pRow, 7).Value = Time '開始時刻
   Application.ScreenUpdating = False
' 項目属性セット
   vSQL = "SELECT COLUMN_NAME,DATA_TYPE,DATA_LENGTH,DATA_PRECISION,DATA_SCALE FROM USER_TAB_COLS WHERE "
   vSQL = vSQL & "TABLE_NAME = '" & Sheets("TABLE_LIST").Cells(pRow, 2).Value
   vSQL = vSQL & "' ORDER BY COLUMN_ID"
   Set OraRs = OraCn.Execute(vSQL)
   If Not OraRs.EOF Then
      i = 1
      Do While Not OraRs.EOF
         Cells(6, i) = OraRs.Fields(0).Value                            ' 項目名
         Cells(7, i) = OraRs.Fields(1).Value                            ' データ型
         If Cells(7, i) = "NUMBER" Then
            Cells(8, i) = OraRs.Fields(3).Value & "," & OraRs.Fields(4) '
         ElseIf Cells(7, i) = "VARCHAR2" Or _
                Cells(7, i) = "CHAR" Then
            Cells(9, i).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.NumberFormatLocal = "@"
            Cells(8, i) = OraRs.Fields(2).Value                         ' データ長
         ElseIf Cells(7, i) = "DATE" Then
            Cells(9, i).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.NumberFormatLocal = "yyyy/mm/dd;@"
            Selection.HorizontalAlignment = xlCenter
            Selection.ShrinkToFit = True
         ElseIf Cells(7, i) = "BLOB" Then
            Cells(9, i).Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.NumberFormatLocal = "@"
            Selection.HorizontalAlignment = xlCenter
         Else
            Cells(8, i) = OraRs.Fields(2).Value                         ' データ長
         End If
         i = i + 1
         OraRs.MoveNext
      Loop
      OraRs.Close
      Set OraRs = Nothing
   End If
' 項目名(コメント名)セット
   i = 1
   Do While Cells(6, i) <> ""
      vSQL = "SELECT COMMENTS FROM USER_COL_COMMENTS WHERE TABLE_NAME = '"
      vSQL = vSQL & Cells(1, 2).Value & "' AND COLUMN_NAME = '"
      vSQL = vSQL & Cells(6, i).Value & "'"
      Set OraRs = OraCn.Execute(vSQL)
      If Not OraRs.EOF Then
         If IsNull(OraRs.Fields(0).Value) Then
         Else
            Cells(5, i) = OraRs.Fields(0).Value
         End If
         OraRs.Close
         Set OraRs = Nothing
      End If
      i = i + 1
   Loop
' 罫線セット
   Range("A5:A8").Select
   Range(Selection, Selection.End(xlToRight)).Select
   Selection.Borders.LineStyle = xlContinuous
   Selection.Interior.ColorIndex = 15
   Selection.Columns.AutoFit
   Selection.HorizontalAlignment = xlCenter
   Range(Cells(1, 1), Cells(6, 1)).Columns.AutoFit
   Range("A5").Select
   Range(Selection, Selection.End(xlToRight)).Select
   Selection.Interior.ColorIndex = 6
   Range("A4").Select
Obj_Rls:
   Sheets("TABLE_LIST").Cells(pRow, 8).Value = Time '終了時刻
   Application.ScreenUpdating = True
   Exit Sub
Err_Han:
   MsgBox (Err.Description)
   GoTo Obj_Rls
End Sub
' 引数 SheetName のシートが実際にあるかチェックする
Function ExistSheet(SheetName) As Boolean
   Dim i As Integer
   ExistSheet = False
   For i = 1 To Sheets.Count
      If StrConv(Sheets(i).Name, 1) = StrConv(SheetName, 1) Then
         ExistSheet = True
         Exit For
      End If
   Next i
End Function
' ADO : テーブルデータ展開
Sub ITEM_LISTING()
   On Err GoTo Err_Han
    
   Dim iRow   As Integer
   Call DB_CONNECT
   If stPass = "" Then Exit Sub
' 検索・展開
   iRow = 8
   Do While Sheets("TABLE_LIST").Cells(iRow, 2).Value <> ""
      If ExistSheet(Sheets("TABLE_LIST").Cells(iRow, 2).Value) Then
         Call SetItemData(iRow)       '検索結果展開
      End If
      iRow = iRow + 1                 '次の行へ
   Loop
   Sheets("TABLE_LIST").Select

   MsgBox "テーブル展開終了！"
   Range("A6").Select
Obj_Rls:
   OraCn.Close
   Set OraCn = Nothing
   Exit Sub
Err_Han:
   MsgBox (Err.Description)
   GoTo Obj_Rls
End Sub
Sub SetItemData(pRow As Integer)
   On Err GoTo Err_Han

   Dim i      As Long
   Dim j      As Integer
   Dim ColMax As Integer
   Dim RecMax As Long
   Dim lRcnt  As Long
   Dim BLOB_ON As Boolean
   If Sheets("TABLE_LIST").Cells(4, 9).Value <> "" Then
      RecMax = CLng(Sheets("TABLE_LIST").Cells(4, 9).Value)
   Else
      RecMax = Sheets("TABLE_LIST").Cells(4, 7).Value          'Sheet-Rows max(65527)
   End If

   Sheets("TABLE_LIST").Cells(pRow, 7).Value = Time '開始時刻
' 件数の取得
   Sheets(Sheets("TABLE_LIST").Cells(pRow, 2).Value).Select
   vSQL = "SELECT COUNT(*) FROM " & Cells(1, 2)
   If Cells(3, 2) <> "" Then
      vSQL = vSQL & " WHERE " & Cells(3, 2).Value
   End If
   Set OraRs = OraCn.Execute(vSQL)
   lRcnt = OraRs.Fields(0)
   If lRcnt > RecMax Then
      MsgBox "検索数が格納最大行を超えています:" & lRcnt & "件", vbOKOnly + vbInformation, "最大行数超過警告"
      lRcnt = RecMax
   End If

' 項目データ検索
   BLOB_ON = False
   vSQL = "SELECT " & Cells(6, 1)
   j = 2
   Do While Cells(6, j) <> ""
      If Cells(7, j) = "BLOB" Then
         vSQL = vSQL & ",DBMS_LOB.SUBSTR(" & Cells(6, j) & ",2000)"
         BLOB_ON = True
      Else
         vSQL = vSQL & "," & Cells(6, j)
      End If
      j = j + 1
   Loop
   If BLOB_ON Then
      vSQL = vSQL & " FROM " & Cells(1, 2)
   Else
      vSQL = "SELECT * FROM " & Cells(1, 2)
   End If
   If Cells(3, 2) <> "" Then
      vSQL = vSQL & " WHERE " & Cells(3, 2).Value
      If lRcnt = RecMax Then
         vSQL = vSQL & " AND ROWNUM <= " & lRcnt
      End If
   Else
      If lRcnt = RecMax Then
         vSQL = vSQL & " WHERE ROWNUM <= " & lRcnt
      End If
   End If
   vSQL = vSQL & " ORDER BY 1,2"
   Set OraRs = OraCn.Execute(vSQL)
   If lRcnt > 0 Then
      If BLOB_ON Then
         ReDim aryData(lRcnt - 1, OraRs.Fields.Count - 1) As Variant
         i = 0
         Do While Not OraRs.EOF And i < lRcnt
            For j = 0 To OraRs.Fields.Count - 1
               If IsNull(OraRs.Fields(j)) Then
                  aryData(i, j) = ""
               Else
                ' EXCEL では長文字列はエラーとなるため912文字でカット
                  aryData(i, j) = Left(OraRs.Fields(j).Value, 912)
               End If
            Next j
            i = i + 1
            OraRs.MoveNext
         Loop
         Range(Cells(9, 1), Cells(9 + i, j)) = aryData
      Else
         Range("A9").CopyFromRecordset OraRs
      End If
      OraRs.Close
      Set OraRs = Nothing
   End If
' 罫線セット
   If Cells(10, 1) <> "" Or Cells(11, 1) <> "" Then
      Range("A9").Select
      ActiveCell.CurrentRegion.Select
      Selection.Borders.LineStyle = xlContinuous
    ' Selection.Columns.AutoFit
   End If
   Range("A4").Select
Obj_Rls:
   Sheets("TABLE_LIST").Cells(pRow, 8).Value = Time '終了時刻
   Sheets("TABLE_LIST").Cells(8, 9).Value = Date    '実行日付
   Exit Sub
Err_Han:
   MsgBox (Err.Description)
   GoTo Obj_Rls
End Sub
