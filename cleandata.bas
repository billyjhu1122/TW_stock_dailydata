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

    ' �]�w CSV �ɮשҦb���ؿ�
    csv_directory = "D:\taiwan_stock_DailyQuotes_20040211_20240322_cleandata"
    
    ' �ˬd�ؿ��O�_�s�b
    If Dir(csv_directory, vbDirectory) = "" Then
        MsgBox "���w���ؿ����s�b�I", vbExclamation
        Exit Sub
    End If
    



    ' �`���B�z�ؿ������C�� CSV �ɮ�
    csv_file = Dir(csv_directory & "\*.csv")
    Do While csv_file <> ""
        ' ���} CSV �ɮ�
        Set wb = Workbooks.Open(Filename:=csv_directory & "\" & csv_file)
        Set ws = wb.Sheets(1)

        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        



        ' ����W�U�����n��ƲM��

        For i = 1 To lastRow
            If ws.Cells(i, 1).Value = "�Ҩ�N��" Then
                importantStartRow = i
                Exit For
            End If
        Next i
    
        ws.Rows("1:" & importantStartRow - 1).EntireRow.Delete
    
        For i = importantStartRow To lastRow
            If ws.Cells(i, 1).Value = "�Ƶ�:" Then
                importantEndRow = i - 1
                Exit For
            End If


        Next i
    
    
        ws.Rows(importantEndRow + 1 & ":" & lastRow).EntireRow.Delete


        ' �ʧR���ĤG��
        ws.Columns(2).Delete



        ' �w�q�����^�媺½Ķ�r��
        Set translationDict = CreateObject("Scripting.Dictionary")
        chineseNames = Array("�Ҩ�N��", "�Ҩ�W��", "����Ѽ�", "���浧��", _
                             "������B", "�}�L��", "�̰���", "�̧C��", _
                             "���L��", "���^(+/-)", "���^���t", "�̫ᴦ�ܶR��", _
                             "�̫ᴦ�ܶR�q", "�̫ᴦ�ܽ��", "�̫ᴦ�ܽ�q", "���q��")
        englishNames = Array("Security Code", "Security Name", "Volume Traded", "Number of Trades", _
                            "Transaction Amount", "Opening Price", "Highest Price", "Lowest Price", _
                            "Closing Price", "Change (+/-)", "Price Change", "Final Bid Price", _
                            "Final Bid Volume", "Final Ask Price", "Final Ask Volume", "Price-to-Earnings Ratio")
        For i = LBound(chineseNames) To UBound(chineseNames)
            translationDict.Add chineseNames(i), englishNames(i)
        Next i
    
        ' �N�ܼƦW�ٱq����½Ķ���^��
        For Each cell In ws.Rows(1).SpecialCells(xlCellTypeConstants)
            If translationDict.Exists(cell.Value) Then
                cell.Value = translationDict(cell.Value)
            End If
        Next cell


       ' ���N--
        Cells.Replace What:="--", Replacement:="", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False _
            , FormulaVersion:=xlReplaceFormula2
        


       lw = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row


       ' �M�z�Ѳ��W�٪����
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


       
       ' �N�ܼƮ榡�Τ@
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
       
       
       ' ���� CSV �ɮ�
       wb.Close SaveChanges:=True
        
       ' �~��B�z�U�@�� CSV �ɮ�
       csv_file = Dir
    Loop

End Sub



