Attribute VB_Name = "Module1"
Sub lotscleandata()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim importantStartRow As Long
    Dim importantEndRow As Long
    Dim i As Long
    Dim lw As Long
    
    Dim cell As Range
    Dim translationDict As Object
    Dim chineseNames As Variant
    Dim englishNames As Variant
    Dim cellValue As String

    Dim csv_directory As String
    Dim csv_file As String
    Dim wb As Workbook
    Dim success As Boolean
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    Dim sql As String

    ' 設定 CSV 檔案所在的目錄
    csv_directory = "D:\taiwan_stock_DailyQuotes_20040211_20240322_cleandata"
    
    ' 檢查目錄是否存在
    If Dir(csv_directory, vbDirectory) = "" Then
        MsgBox "指定的目錄不存在！", vbExclamation
        Exit Sub
    End If
    



    ' 循環處理目錄中的每個 CSV 檔案
    csv_file = Dir(csv_directory & "\*.csv")
    Do While csv_file <> ""
        ' 打開 CSV 檔案
        Set wb = Workbooks.Open(Filename:=csv_directory & "\" & csv_file)
        Set ws = wb.Sheets(1)

        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        



        ' 先把上下不重要資料清除

        For i = 1 To lastRow
            If ws.Cells(i, 1).Value = "證券代號" Then
                importantStartRow = i
                Exit For
            End If
        Next i
    
        ws.Rows("1:" & importantStartRow - 1).EntireRow.Delete
    
        For i = importantStartRow To lastRow
            If ws.Cells(i, 1).Value = "備註:" Then
                importantEndRow = i - 1
                Exit For
            End If


        Next i
    
    
        ws.Rows(importantEndRow + 1 & ":" & lastRow).EntireRow.Delete


        ' 缺刪除第二欄
        ws.Columns(2).Delete



        ' 定義中文到英文的翻譯字典
        Set translationDict = CreateObject("Scripting.Dictionary")
        chineseNames = Array("證券代號", "證券名稱", "成交股數", "成交筆數", _
                             "成交金額", "開盤價", "最高價", "最低價", _
                             "收盤價", "漲跌(+/-)", "漲跌價差", "最後揭示買價", _
                             "最後揭示買量", "最後揭示賣價", "最後揭示賣量", "本益比")
        englishNames = Array("Security Code", "Security Name", "Volume Traded", "Number of Trades", _
                            "Transaction Amount", "Opening Price", "Highest Price", "Lowest Price", _
                            "Closing Price", "Change (+/-)", "Price Change", "Final Bid Price", _
                            "Final Bid Volume", "Final Ask Price", "Final Ask Volume", "Price-to-Earnings Ratio")
        For i = LBound(chineseNames) To UBound(chineseNames)
            translationDict.Add chineseNames(i), englishNames(i)
        Next i
    
        ' 將變數名稱從中文翻譯為英文
        For Each cell In ws.Rows(1).SpecialCells(xlCellTypeConstants)
            If translationDict.Exists(cell.Value) Then
                cell.Value = translationDict(cell.Value)
            End If
        Next cell


       ' 替代--
        Cells.Replace What:="--", Replacement:="", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False _
            , FormulaVersion:=xlReplaceFormula2
        


       lw = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row


       ' 清理股票名稱的資料
       ws.Range("B1:B" & lw).Select
       Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
       Range("A1").Select
       Selection.Copy
       Range("B1").Select
       ActiveSheet.Paste
       Range("B2").Select
       Application.CutCopyMode = False
       ActiveCell.FormulaR1C1 = "=""'""&RC[-1]"
       Range("B2").Select
       Selection.Copy
       ws.Range("B2:B" & lw).Select
       Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
           SkipBlanks:=False, Transpose:=False
       Application.CutCopyMode = False
       Selection.Copy
       Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
           :=False, Transpose:=False
       ws.Range("A1:A" & lw).Select
       Application.CutCopyMode = False
       Selection.Delete Shift:=xlToLeft


       
       ' 將變數格式統一
       ws.Range("A1:A" & lw).Select
       Selection.NumberFormatLocal = "@"
       ws.Range("B1:B" & lw).Select
       Selection.NumberFormatLocal = "0"
       ws.Range("C1:C" & lw).Select
       Selection.NumberFormatLocal = "0"
       ws.Range("D1:D" & lw).Select
       Selection.NumberFormatLocal = "0"
       ws.Range("E1:E" & lw).Select
       Selection.NumberFormatLocal = "0.00"
       ws.Range("F1:F" & lw).Select
       Selection.NumberFormatLocal = "0.00"
       ws.Range("G1:G" & lw).Select
       Selection.NumberFormatLocal = "0.00"
       ws.Range("H1:H" & lastRow).Select
       Selection.NumberFormatLocal = "0.00"
       ws.Range("I1:I" & lastRow).Select
       Selection.NumberFormatLocal = "@"
       ws.Range("J1:J" & lastRow).Select
       Selection.NumberFormatLocal = "0.00"
       ws.Range("K1:K" & lastRow).Select
       Selection.NumberFormatLocal = "0.00"
       ws.Range("L1:L" & lastRow).Select
       Selection.NumberFormatLocal = "0"
       ws.Range("M1:M" & lastRow).Select
       Selection.NumberFormatLocal = "0.00"
       ws.Range("N1:N" & lastRow).Select
       Selection.NumberFormatLocal = "0"
       ws.Range("O1:O" & lastRow).Select
       Selection.NumberFormatLocal = "0.00"
       
       
       ' 關閉 CSV 檔案
       wb.Close SaveChanges:=True
        
       ' 繼續處理下一個 CSV 檔案
       csv_file = Dir
    Loop

End Sub



