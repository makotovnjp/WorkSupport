Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Public Class Check_stock

#Region "Variable"
    Private hinban_Combobox As ComboBox
    Private hinmei_Combobox As ComboBox
    Private search_index As Integer
#End Region

#Region "Main Function"
    '1月～9月なら01～09を返す。
    Public Function Reform_Month(ByVal mthstr As String) As String
        If Integer.Parse(mthstr) < 10 Then
            Return "0" & mthstr
        Else
            Return mthstr
        End If
    End Function


    Private Sub Init_Display()
        Label1.Visible = True
        Label2.Visible = True
        Label9.Visible = True
        Label3.Visible = True
        Label4.Visible = True
        Label5.Visible = True
        Label10.Visible = True
        Label11.Visible = True
        Label3.Text = hinban_Combobox.Items.Item(search_index)
        Label4.Text = hinmei_Combobox.Items.Item(search_index)
        Label5.Text = ComboBox3.Items.Item(search_index)
        Label10.Text = ComboBox4.Items.Item(search_index)
    End Sub

    '決められた年月の入出荷情報習得
    Private Sub Get_info_of_the_month(ByVal filelink As String, ByVal intout As String)
        '在庫ファイルから読み取り
        Dim app As Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet
        Dim row_no As Integer
        Dim strspc As String = ""
        Dim dgv As New DataGridView
        Dim str_day, str_company, str_number As String

        dgv = NyukaKakutei_DataGridView

        app = CreateObject("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False

        ' ファイルが存在しているかどうか確認する
        If System.IO.File.Exists(filelink) = False Then
            Exit Sub
        End If

        'File Open
        book = app.Workbooks.Open(filelink)
        sheet = book.Worksheets(1)

        row_no = 2
        While sheet.Cells(row_no, 2).Value <> ""
            If sheet.Cells(row_no, 3).Value = Label3.Text Then
                '検索品番であれば：
                str_day = sheet.Cells(row_no, 1).Value '入出日付
                str_company = sheet.Cells(row_no, 2).Value '入先又は取り先
                str_number = sheet.Cells(row_no, 6).Value() '入出量
                dgv.Rows.Add(intout, str_day, str_company, str_number, "")
            End If
            row_no += 1
        End While

        'File Close
        book.Close()
        app.Quit()

        ' オブジェクトを解放します。
        sheet = Nothing
        book = Nothing
        app = Nothing
    End Sub

    '調べたい区間の入出荷情報表示
    Private Sub Display_StockInfo()
        Dim intyear As Integer = Integer.Parse(YearComboBox3.Text)
        Dim intmonth As Integer = Integer.Parse(MonthComboBox4.Text)
        Dim stryear, strmonth, folderlink As String

        NyukaKakutei_DataGridView.Rows.Clear()

        While True
            stryear = intyear.ToString
            folderlink = "C:\業務管理ソフトData\商品情報\" + stryear + "\"
            'フォルダ が存在しているかどうか確認する
            If System.IO.Directory.Exists(folderlink) Then
                strmonth = Reform_Month(intmonth.ToString)
                folderlink += stryear + strmonth + "\"
                'フォルダ が存在しているかどうか確認する
                If System.IO.Directory.Exists(folderlink) Then
                    Get_info_of_the_month(folderlink + "入荷" + stryear + strmonth + ".xlsx", "入荷")
                    Get_info_of_the_month(folderlink + "出荷" + stryear + strmonth + ".xlsx", "出荷")
                End If
            End If

            intmonth += 1
            If intmonth = 13 Then
                intmonth = 1
                intyear += 1
            End If
            If intyear.ToString = YearComboBox6.Text Then
                If intmonth > Integer.Parse(MonthComboBox5.Text) Then
                    Exit While
                End If
            End If
        End While

    End Sub

    '選択した時間の妥当性確認
    Private Function SelectedTimeOK() As Boolean
        Dim selectedtime As Integer
        selectedtime = (Integer.Parse(YearComboBox6.Text) - Integer.Parse(YearComboBox3.Text)) * 12 + Integer.Parse(MonthComboBox5.Text) - Integer.Parse(MonthComboBox4.Text)
        Return selectedtime >= 0
    End Function

    '確認ボタンんを押す時
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ComboBox2.Text = "" Then
            MessageBox.Show(Label6.Text, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        If SelectedTimeOK() = False Then
            MessageBox.Show("正しい時間を入力してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        '---------入力条件が正しい場合、検索結果を表示------
        search_index = ComboBox2.SelectedIndex
        Init_Display()
        '表を可視化
        NyukaKakutei_DataGridView.Visible = True
        '検索品番や品名に関する入荷と出荷情報を表示する
        Display_StockInfo()
        'WriteToFile_Button.Visible = True
    End Sub

    '確認用の選択した品番や品名リストの作成
    Private Sub Init_ComboboxOfPartName(ByVal choicekey As String)
        '在庫ファイル名取得
        Dim dtToday As DateTime = DateTime.Today ' 現在の日付を取得する
        Dim filename = "C:\業務管理ソフトData\商品情報\" + dtToday.Year.ToString + "\"
        Dim str_yearmonth As String = dtToday.Year.ToString + Reform_Month(dtToday.Month.ToString)
        filename += str_yearmonth + "\" + "在庫" + str_yearmonth + ".xlsx"

        '在庫ファイルから読み取り
        Dim app As Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet
        Dim row_no As Integer
        Dim strspc As String = ""
        Dim sum As Integer = 0

        If choicekey = "品番" Then
            Label6.Text = "品番を選択してください"
            hinban_Combobox = ComboBox2
            hinmei_Combobox = ComboBox1
        Else
            Label6.Text = "品名を選択してください"
            hinban_Combobox = ComboBox1
            hinmei_Combobox = ComboBox2
        End If

        app = CreateObject("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        'File Open
        book = app.Workbooks.Open(filename)
        sheet = book.Worksheets(1)

        row_no = 4
        hinban_Combobox.Items.Clear()
        hinmei_Combobox.Items.Clear()
        While sheet.Cells(row_no, 1).Value <> ""
            hinban_Combobox.Items.Add(sheet.Cells(row_no, 1).Value)
            hinmei_Combobox.Items.Add(sheet.Cells(row_no, 2).Value)
            ComboBox3.Items.Add(sheet.Cells(row_no, 3).Value)
            ComboBox4.Items.Add(sheet.Cells(row_no, 40).Value)
            row_no += 1
        End While

        'File Close
        book.Close()
        app.Quit()

        ' オブジェクトを解放します。
        sheet = Nothing
        book = Nothing
        app = Nothing
    End Sub

    '年Comboboxリスト作成
    Private Sub Init_ComboboxYear()
        Dim dtToday As DateTime = DateTime.Today ' 現在の日付を取得する
        Dim lastmonth As DateTime = dtToday.AddMonths(-1)
        Dim thisyear As Integer = Integer.Parse(Today.Year.ToString)
        Dim i As Integer
        YearComboBox3.Items.Clear()
        YearComboBox6.Items.Clear()
        For i = 2010 To thisyear
            YearComboBox3.Items.Add(i.ToString)
            YearComboBox6.Items.Add(i.ToString)
        Next
        YearComboBox3.Text = lastmonth.Year.ToString
        MonthComboBox4.Text = Reform_Month(lastmonth.Month.ToString)
        YearComboBox6.Text = dtToday.Year.ToString
        MonthComboBox5.Text = Reform_Month(dtToday.Month.ToString)
    End Sub

    Private Sub Check_stock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Partnumber_CheckBox.Checked = True
        NameOfProduct_CheckBox.Checked = False
        '確認用の選択した品番や品名リストの作成
        Init_ComboboxOfPartName("品番")
        Init_ComboboxYear()
    End Sub

    Private Sub Partnumber_CheckBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Partnumber_CheckBox.Click
        Partnumber_CheckBox.Checked = True
        NameOfProduct_CheckBox.Checked = False
        '確認用の選択した品番や品名リストの作成
        Init_ComboboxOfPartName("品番")
    End Sub

    Private Sub NameOfProduct_CheckBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NameOfProduct_CheckBox.Click
        NameOfProduct_CheckBox.Checked = True
        Partnumber_CheckBox.Checked = False
        '確認用の選択した品番や品名リストの作成
        Init_ComboboxOfPartName("品名")
    End Sub
#End Region


End Class