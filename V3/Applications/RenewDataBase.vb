Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

'********************************
'本クラスのコーディングルール
'********************************

'*************
'命名規則：キャメルケース
'*************

'PreFixのルール
'Public変数　 : ARRSCHD_
'Private変数  : l_

'*************
'暫定対応：Hoang(todo)
'*************

Public Class RenewDataBase

#Region "Public Method"
    '1月～9月なら01～09を返す。
    Public Function Reform_Month(ByVal mthstr As String) As String
        If Integer.Parse(mthstr) < 10 Then
            Return "0" & mthstr
        Else
            Return mthstr
        End If
    End Function

    '********************************************
    '今年や今月のフォルダがあるかどうか確認
    'なければフォルダ作成し、先月のデータをコピーする]
    '********************************************
    Public Sub ChangeDataForNewYearOrNewMonth()
        Dim dtToday As DateTime = DateTime.Today ' 現在の日付を取得する
        Dim lastmonth As DateTime = dtToday.AddMonths(-1)
        Dim linkyear, linkmonth As String
        linkyear = ""

        ' 今年のフォルダ が存在しているかどうか確認する
        linkyear = "C:\業務管理ソフトData\商品情報\" & dtToday.Year.ToString
        If System.IO.Directory.Exists(linkyear) = False Then
            System.IO.Directory.CreateDirectory(linkyear)
            ChangeDataForNewYear(linkyear)
        End If

        ' 今月のフォルダ が存在しているかどうか確認する
        linkmonth = linkyear + "\" + dtToday.Year.ToString + Reform_Month(dtToday.Month.ToString)
        If System.IO.Directory.Exists(linkmonth) = False Then
            System.IO.Directory.CreateDirectory(linkmonth)
            ChangeDataForNewMonth(linkmonth)
        End If

    End Sub

    Private Sub CopytemplateFile(ByVal fromlink As String, ByVal tolink As String, ByVal filename As String)
        'Templateファイルからコピーする
        System.IO.File.Copy(fromlink + "\" + filename, tolink + "\" + filename)
    End Sub

    '********************************************
    'イベント：
    '処理：入荷・出荷テンプレートcopy、在庫ファイルcopy
    '********************************************
    Private Sub ChangeDataForNewMonth(ByVal des_link As String)
        Dim dtToday As DateTime = DateTime.Today ' 現在の日付を取得する
        Dim fromlink As String = "C:\業務管理ソフトData\テンプレート情報\"
        Dim str_thismonth As String = dtToday.Year.ToString + Reform_Month(dtToday.Month.ToString)
        System.IO.File.Copy(fromlink + "Template_入荷.xlsx", des_link + "\" + "入荷" + str_thismonth + ".xlsx")
        System.IO.File.Copy(fromlink + "Template_出荷.xlsx", des_link + "\" + "出荷" + str_thismonth + ".xlsx")

        Dim lastmonth As DateTime = dtToday.AddMonths(-1)
        Dim str_lastmonth As String = lastmonth.Year.ToString + Reform_Month(lastmonth.Month.ToString)
        fromlink = "C:\業務管理ソフトData\商品情報\" + lastmonth.Year.ToString + "\" + str_lastmonth + "\在庫" + str_lastmonth + ".xlsx"
        System.IO.File.Copy(fromlink, des_link + "\" + "在庫" + str_thismonth + ".xlsx")

        '在庫ファイル内容初期化
        Dim writefile_name As String
        Dim app As Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet
        Dim row_no, col_no As Integer
        Dim strspc As String = ""
        app = CreateObject("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        writefile_name = des_link + "\" + "在庫" + str_thismonth + ".xlsx"
        'File Open
        book = app.Workbooks.Open(writefile_name)
        sheet = book.Worksheets(1)

        sheet.Cells(1, 2).Value = Today.Year.ToString + "年" + Reform_Month(Today.Month.ToString) + "月間販売実績表/在庫表"

        '初期化
        For row_no = 4 To 100
            sheet.Cells(row_no, 6).Value = sheet.Cells(row_no, 40).Value
            For col_no = 7 To 39
                sheet.Cells(row_no, col_no).Value = strspc
            Next
        Next
        'File Save
        book.Save()

        'File Close
        book.Close()

        app.Quit()

        ' オブジェクトを解放します。
        sheet = Nothing
        book = Nothing
        app = Nothing
    End Sub

    '********************************************
    'イベント：
    '処理：
    '********************************************
    Private Sub ChangeDataForNewYear(ByVal link As String)

    End Sub


#End Region
End Class
