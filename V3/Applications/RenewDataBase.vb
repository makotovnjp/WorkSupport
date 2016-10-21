Imports Microsoft.Office.Interop

'********************************
'本クラスのコーディングルール
'********************************

'*************
'命名規則：キャメルケース
'*************
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
        linkyear = DataPathDefinition.GetProductDataPath() + "\" & dtToday.Year.ToString
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
        If IO.File.Exists(fromlink + "\" + filename) Then 'Fileが存在する
            IO.File.Copy(fromlink + "\" + filename, tolink + "\" + filename)
        Else
            MsgBox("File:" + fromlink + "\" + filename + "が存在しない")
        End If
    End Sub

    '********************************************
    'イベント：
    '処理：入荷・出荷テンプレートcopy、在庫ファイルcopy
    '********************************************
    Private Sub ChangeDataForNewMonth(ByVal des_link As String)
        Dim dtToday As DateTime = DateTime.Today ' 現在の日付を取得する
        Dim fromlink As String = DataPathDefinition.GetTemplateDataPath() + "\"
        Dim str_thismonth As String = dtToday.Year.ToString + Reform_Month(dtToday.Month.ToString)

        'コピーする
        If IO.File.Exists(fromlink + "Template_入荷.xlsx") Then 'Fileが存在する
            IO.File.Copy(fromlink + "Template_入荷.xlsx", des_link + "\" + "入荷" + str_thismonth + ".xlsx")
        Else
            MsgBox("File:" + fromlink + "Template_入荷.xlsx" + "が存在しない")
        End If

        If IO.File.Exists(fromlink + "Template_出荷.xlsx") Then 'Fileが存在する
            IO.File.Copy(fromlink + "Template_出荷.xlsx", des_link + "\" + "出荷" + str_thismonth + ".xlsx")
        Else
            MsgBox("File:" + fromlink + "Template_出荷.xlsx" + "が存在しない")
        End If


        Dim lastmonth As DateTime = dtToday.AddMonths(-1)
        Dim str_lastmonth As String = lastmonth.Year.ToString + Reform_Month(lastmonth.Month.ToString)
        fromlink = DataPathDefinition.GetProductDataPath() + "\" + lastmonth.Year.ToString + "\" + str_lastmonth + "\在庫" + str_lastmonth + ".xlsx"

        'ファイルからコピーする
        If IO.File.Exists(fromlink) Then 'Fileが存在する
            IO.File.Copy(fromlink, des_link + "\" + "在庫" + str_thismonth + ".xlsx")
        Else
            MsgBox("File:" + fromlink + "が存在しない")
        End If

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
        If IO.File.Exists(writefile_name) Then 'Fileが存在する
            book = app.Workbooks.Open(writefile_name)
        Else
            MsgBox("File:" + writefile_name + "が存在しない")
            Exit Sub
        End If

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
