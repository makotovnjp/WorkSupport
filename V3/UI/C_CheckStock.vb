Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel


Public Class C_CheckStock

    '在庫確認-------------------------
    Public Sub Stock_Display()

        '在庫ファイル名取得
        Dim dtToday As DateTime = DateTime.Today ' 現在の日付を取得する
        Dim filename = "C:\業務管理ソフトData\商品情報\" + dtToday.Year.ToString + "\"
        Dim str_yearmonth As String = dtToday.Year.ToString + Microsoft.VisualBasic.Right("0" & dtToday.Month.ToString, 2)
        filename += str_yearmonth + "\" + "在庫" + str_yearmonth + ".xlsx"

        '在庫ファイルから読み取り
        Dim product_code, product_name, product_slot, stock_number As String
        Dim app As Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet
        Dim row_no As Integer
        Dim strspc As String = ""
        Dim dgv As New DataGridView
        Dim sum As Integer = 0
        MainFunction.DataGridView1.Rows.Clear()
        dgv = MainFunction.DataGridView1

        app = CreateObject("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False

        If System.IO.File.Exists(filename) = False Then
            Exit Sub
        End If

        'File Open
        If IO.File.Exists(filename) Then 'Fileが存在する
            book = app.Workbooks.Open(filename)
        Else
            MsgBox("File:" + filename + "が存在しない")
            Exit Sub
        End If

        sheet = book.Worksheets(1)

        row_no = 4
        While sheet.Cells(row_no, 1).Value <> ""
            If sheet.Cells(row_no, 40).Value > 0 Then
                product_code = sheet.Cells(row_no, 1).Value.ToString
                product_name = sheet.Cells(row_no, 2).Value.ToString
                product_slot = sheet.Cells(row_no, 3).Value
                stock_number = sheet.Cells(row_no, 40).Value
                dgv.Rows.Add((row_no - 3).ToString, product_code, product_name, product_slot, stock_number)
                sum += 1
            End If
            row_no += 1
        End While
        MainFunction.Label16.Text = sum.ToString

        'File Close
        book.Close()

        app.Quit()

        ' オブジェクトを解放します。
        sheet = Nothing
        book = Nothing
        app = Nothing
    End Sub
End Class
