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
'Public変数　 : OUTGOODS_
'Private変数  : l_

'*************
'暫定対応：thanh(todo)
'*************

Public Class OutputGoods

#Region "データ部"
    Public Const OUTGOODS_OK As Integer = 1
    Public Const OUTGOODS_NG As Integer = (-1)

    '出荷ファイルの検索用
    Private Const OUTGOODS_START_ROW As Integer = 2
    Private Const OUTGOODS_MAX_ROW As Integer = 1000


    '出荷予定情報の列番号
    Private Const OUTGOODS_READFILE_DAY_COL_NO As Integer = 1     '出荷予定日
    Private Const OUTGOODS_READFILE_CUSTOMER_COL_NO As Integer = 2  '届先名
    Private Const OUTGOODS_READFILE_CODE_COL_NO As Integer = 3     '品番の列番号
    Private Const OUTGOODS_READFILE_NAME_COL_NO As Integer = 4     '品名の列番号
    Private Const OUTGOODS_READFILE_SLOT_COL_NO As Integer = 5     '出数の列番号
    Private Const OUTGOODS_READFILE_PRODUCT_OUTPUT_COL_NO As Integer = 6   '出庫数の列番号

    'DataGridViewの列番号
    Private Const OUTGOODS_DGV_CUSTOMER_NAME As Integer = 1     '出荷予定日
    Private Const OUTGOODS_DGV_DAY_COL_NO As Integer = 2     '出荷予定日
    Private Const OUTGOODS_DGV_CODE_COL_NO As Integer = 3     '品番の列番号
    Private Const OUTGOODS_DGV_NAME_COL_NO As Integer = 4    '品名の列番号
    Private Const OUTGOODS_DGV_SLOT_COL_NO As Integer = 5     '出数の列番号
    Private Const OUTGOODS_DGV_PRODUCT_OUT_COL_NO As Integer = 6   '出庫数の列番号

    '在庫表の列番号:InputGoodsのクラスの定義と同じ値にしないといけない
    Private Const INVENTORY_PRODUCTS_CODE_COL_NO As Integer = 1
    Private Const INVENTORY_PRODUCTS_NAME_COL_NO As Integer = 2
    Private Const INVENTORY_PRODUCTS_SLOT_COL_NO As Integer = 3
    Private Const INVENTORY_PRODUCTS_MONTH_BEGIN_COL_NO As Integer = 6      '月初在庫の列番号
    Private Const INVENTORY_PRODUCTS_MONTH_END_COL_NO As Integer = 40       '現在庫の列番号
    Private Const INVENTORY_PRODUCTS_STORAGE_COL_NO As Integer = 7          '今月出庫数合計
    Private Const INVENTORY_PRODUCTS_DATE_START_COL_NO As Integer = 7
    Private Const INVENTORY_START_ROW_NO As Integer = 4
    Private Const INVENTORY_MAX_ROW As Integer = 1000
    '出荷ファイルの列番号
    Private Const SYUKA_DAY_COL_NO As Integer = 1     '出荷予定日
    Private Const SYUKA_PRODUCTS_CUSTOMER_COL_NO As Integer = 2
    Private Const SYUKA_PRODUCTS_CODE_COL_NO As Integer = 3
    Private Const SYUKA_PRODUCTS_NAME_COL_NO As Integer = 4
    Private Const SYUKA_PRODUCTS_SLOT_COL_NO As Integer = 5
    Private Const SYUKA_PRODUCTS_SYUKO_COL_NO As Integer = 6

    Private Const OUTGOODS_TEMPLATE_FILENAME = "Template_出荷.xlsx"

#End Region

#Region "Public Method"
    ''' <summary>
    ''' 出荷確定の初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Init()
        'DataGridViewのデータを初期化
        InitDataGridView()

        '読み込んだデータをDataGridViewにセットする
        SetDataGridView()

        '未出荷項目数を表示する
        MainFunction.Num_No_Outputs.Text = (GetNotInputItems() - 1).ToString

    End Sub

    ''' <summary>
    ''' 出荷確定用のData Grid Viewのチェックボックスを設定する
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SetCheckBox(ByVal Status As Boolean)
        Dim i As Integer
        Dim RowCount As Integer
        Dim dgv As New DataGridView

        dgv = MainFunction.OutputKakutei_DataGridView

        RowCount = GetNotInputItems()

        If RowCount > 1 Then
            For i = 0 To RowCount - 2
                dgv(0, i).Value = Status
            Next
        End If

    End Sub

    ''' <summary>
    ''' 出荷ファイルの情報を更新する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function WriteOutputGoodsFile() As Integer
        Dim i As Integer
        Dim delete_no As Integer = 0    '削除した回数
        Dim Syuka_Number As Integer
        Dim dgv As New DataGridView

        Dim row_no As Integer = 2                                           '検索開始の行番号
        Dim day As Integer = 0                                              '日付の列番号
        Dim current_output_num As Integer = 0                              '同じ日にすでに出荷した数

        '商品の情報
        Dim customer_name As String = ""      '届先
        Dim product_day As String = ""      '出荷予定日
        Dim product_code As String = ""     '品番
        Dim product_name As String = ""     '品名
        Dim product_slot As String = ""     '出数
        Dim product_output As String = ""  '出庫数

        '保存ファイル名
        Dim save_filename As String = ""                '保存情報のファイル名
        Dim save_foldername As String = ""              '保存情報のフォルダ名
        Dim last_month_save_filename As String = ""     '先月の保存情報のファイル名

        Dim year_value As String = ""
        Dim month_value As String = ""

        'Excel処理ための変数
        Dim app As Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet

        dgv = MainFunction.OutputKakutei_DataGridView
        Syuka_Number = GetNotInputItems()

        app = CreateObject("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False

        If Syuka_Number > 1 Then

            For i = 0 To Syuka_Number - 2
                '出庫確定のチェックボックスがあるなら
                If dgv(0, i - delete_no).Value = True Then   '1行を削除した後に、行番号を1個増えるから
                    '出庫情報を設定する
                    customer_name = dgv(OUTGOODS_DGV_CUSTOMER_NAME, i - delete_no).Value.ToString
                    product_day = dgv(OUTGOODS_DGV_DAY_COL_NO, i - delete_no).Value.ToString
                    product_code = dgv(OUTGOODS_DGV_CODE_COL_NO, i - delete_no).Value.ToString
                    product_name = dgv(OUTGOODS_DGV_NAME_COL_NO, i - delete_no).Value.ToString
                    product_slot = dgv(OUTGOODS_DGV_SLOT_COL_NO, i - delete_no).Value.ToString
                    product_output = dgv(OUTGOODS_DGV_PRODUCT_OUT_COL_NO, i - delete_no).Value.ToString

                    '出庫情報を保存するファイルを探す
                    year_value = GetYearValue(product_day)
                    month_value = GetMonthValue(product_day)
                    save_filename = GetSyukaFileName(customer_name, year_value, month_value)

                    'File Open
                    If IO.File.Exists(save_filename) Then 'Fileが存在する
                        book = app.Workbooks.Open(save_filename)
                    Else
                        MsgBox("File:" + save_filename + "が存在しない")
                        Return OUTGOODS_NG
                    End If


                    sheet = book.Worksheets(1)


                    '出荷ファイルの挿出行の決定
                    For row_no = 2 To OUTGOODS_MAX_ROW
                        If Len(sheet.Cells(row_no, SYUKA_PRODUCTS_CUSTOMER_COL_NO).Value) = 0 Then
                            Exit For
                        Else
                            'row_no = row_no + 1

                        End If

                    Next

                    '出荷確定日付
                    sheet.Cells(row_no, SYUKA_DAY_COL_NO).Value = product_day

                    '届先
                    sheet.Cells(row_no, SYUKA_PRODUCTS_CUSTOMER_COL_NO).Value = customer_name

                    '品番
                    sheet.Cells(row_no, SYUKA_PRODUCTS_CODE_COL_NO).Value = product_code

                    '品名
                    sheet.Cells(row_no, SYUKA_PRODUCTS_NAME_COL_NO).Value = product_name

                    '出数
                    sheet.Cells(row_no, SYUKA_PRODUCTS_SLOT_COL_NO).Value = product_slot

                    '出庫
                    sheet.Cells(row_no, SYUKA_PRODUCTS_SYUKO_COL_NO).Value = product_output

                    row_no = row_no + 1

                    'File Save
                    book.Save()

                    'File Close
                    book.Close()

                End If
            Next

            app.Quit()

            ' オブジェクトを解放します。
            sheet = Nothing
            book = Nothing
            app = Nothing

        End If

        Return OUTGOODS_OK

    End Function

    ''' <summary>
    ''' 在庫情報のExcelを更新する
    ''' 更新の仕方は日付の列に出荷の数を出れる
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InputWareHouse() As Integer
        Dim i As Integer
        Dim delete_no As Integer = 0    '削除した回数
        Dim Syuka_Number As Integer
        Dim dgv As New DataGridView

        Dim COutputSchedule As New OutputGoodsSchedule
        Dim product_ino As OutputGoodsSchedule.ProductInfo

        Dim row_no As Integer = 4                                           '検索開始の行番号
        Dim col_no_data As Integer = INVENTORY_PRODUCTS_DATE_START_COL_NO   '日付開始の列番号
        Dim day As Integer = 0                                              '日付の列番号
        Dim current_output_num As Integer = 0                              '同じ日にすでに出荷した数

        '商品の情報
        Dim customer_name As String = ""      '取引先
        Dim product_day As String = ""      '出荷予定日
        Dim product_code As String = ""     '品番
        Dim product_name As String = ""     '品名
        Dim product_slot As String = ""     '出数
        Dim product_out As String = ""  '出庫数

        '保存ファイル名
        Dim save_filename As String = ""                '保存情報のファイル名
        Dim save_foldername As String = ""              '保存情報のフォルダ名
        Dim last_month_save_filename As String = ""     '先月の保存情報のファイル名

        Dim year_value As String = ""
        Dim month_value As String = ""
        Dim day_value As Integer

        'Excel処理ための変数
        Dim app As Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet

        dgv = MainFunction.OutputKakutei_DataGridView
        Syuka_Number = GetNotInputItems() ' 未出荷確定の商品数を取得する

        If Syuka_Number > 1 Then
            app = CreateObject("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False

            For i = 0 To Syuka_Number - 2
                '出庫確定
                If dgv(0, i - delete_no).Value = True Then   '1行を削除した後に、行番号を1個増えるから
                    '出庫情報を収得する
                    customer_name = dgv(OUTGOODS_DGV_CUSTOMER_NAME, i - delete_no).Value.ToString
                    product_day = dgv(OUTGOODS_DGV_DAY_COL_NO, i - delete_no).Value.ToString
                    product_code = dgv(OUTGOODS_DGV_CODE_COL_NO, i - delete_no).Value.ToString
                    product_name = dgv(OUTGOODS_DGV_NAME_COL_NO, i - delete_no).Value.ToString
                    product_slot = dgv(OUTGOODS_DGV_SLOT_COL_NO, i - delete_no).Value.ToString
                    product_out = dgv(OUTGOODS_DGV_PRODUCT_OUT_COL_NO, i - delete_no).Value.ToString

                    '出庫年月収得
                    year_value = GetYearValue(product_day)
                    month_value = GetMonthValue(product_day)
                    '今月の在庫ファイルのリンク収得
                    save_filename = GetFileNameSaveInputWare(customer_name, year_value, month_value)

                    '###############　この部分は要らないからとりあえずコメントアウト(Hoang)　　　########
                    'If System.IO.File.Exists(save_filename) Then    '在庫表のファイルが存在する
                    '    'Do Nothing
                    'Else    '在庫表のファイルが存在しない場合、新規ファイルを作成する
                    '    save_foldername = GetFolderNameSaveInputGoods(year_value, month_value)

                    '    'フォルダ作成
                    '    System.IO.Directory.CreateDirectory(save_foldername)

                    '    'Templateのファイルから新規の在庫表ファイルをコピーする
                    '    '本来なら、このルートには出らないはず
                    '    'Hoangの修正で、システム起動するときに、書き込み用のファイルは全部それっているはず？
                    '    System.IO.File.Copy(GetSaveTemplateFileName, save_filename)

                    '    last_month_save_filename = GetLastMonthSaveFileName(year_value, month_value, save_filename)
                    '    '先月のファイル存在確認
                    '    If last_month_save_filename <> "" Then
                    '        'ファイルが存在する場合
                    '        '先月の在庫情報からコピーする
                    '        CopyFromLastMonthInventory(last_month_save_filename, save_filename)
                    '    Else
                    '        'ファイルが存在しない場合: Do Nothing

                    '    End If

                    'End If
                    '##################################################################################

                    '在庫ファイルを開く
                    If IO.File.Exists(save_filename) Then 'Fileが存在する
                        book = app.Workbooks.Open(save_filename)
                    Else
                        MsgBox("File:" + save_filename + "が存在しない")
                        Return OUTGOODS_NG
                    End If
                    sheet = book.Worksheets(1)

                    '日付の列番号を決定
                    day_value = GetDayValue(product_day)
                    col_no_data = INVENTORY_PRODUCTS_DATE_START_COL_NO + day_value

                    '同じ品番の行番号を探す。同じ品番が無い場合、在庫ファイルに追加する
                    For row_no = INVENTORY_START_ROW_NO To INVENTORY_MAX_ROW
                        If sheet.Cells(row_no, INVENTORY_PRODUCTS_CODE_COL_NO).Value = product_code Then
                            '見つけたのでループから出る
                            Exit For
                        Else
                            If Len(sheet.Cells(row_no, INVENTORY_PRODUCTS_CODE_COL_NO).Value) = 0 Then
                                '在庫ファイルに存在しないため、在庫ファイルに追加
                                sheet.Cells(row_no, 1).Value = product_code
                                sheet.Cells(row_no, 2).Value = product_name
                                sheet.Cells(row_no, 3).Value = product_slot
                                sheet.Cells(row_no, 6).Value = "0"
                                sheet.Cells(row_no, 7).Value = "0"
                                sheet.Cells(row_no, 40).Value = "0"
                                Exit For
                            End If

                        End If

                    Next

                    '現在の出庫
                    current_output_num = sheet.Cells(row_no, col_no_data).Value
                    sheet.Cells(row_no, col_no_data).Value = product_out + current_output_num

                    'File Save
                    book.Save()

                    'File Close
                    book.Close()

                    '出荷予定ファイルの該当データを削除
                    product_ino.customer_name = customer_name
                    product_ino.product_day = product_day
                    product_ino.product_code = product_code
                    product_ino.product_name = product_name
                    product_ino.product_slot = product_slot
                    product_ino.product_out = product_out
                    COutputSchedule.DeleteScheduleData(product_ino)

                    dgv.Rows.RemoveAt(i - delete_no)
                    delete_no = delete_no + 1
                End If
            Next

            app.Quit()

            ' オブジェクトを解放します。
            sheet = Nothing
            book = Nothing
            app = Nothing

        End If

        '未出荷商品数を更新する
        MainFunction.Num_No_Outputs.Text = (GetNotInputItems() - 1).ToString

        If delete_no = 0 Then
            MsgBox("出荷確定商品を選択ください")
        Else
            MsgBox("出荷確定完了")
        End If

        Return OUTGOODS_OK

    End Function

#End Region

#Region "Private Method"
    ''' <summary>
    ''' DataGridViewの初期化
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitDataGridView()
        Dim i As Integer
        Dim RowCount As Integer
        Dim delete_no As Integer = 0    '削除した回数
        Dim dgv As New DataGridView

        dgv = MainFunction.OutputKakutei_DataGridView
        RowCount = GetNotInputItems()

        If RowCount > 1 Then
            For i = 0 To RowCount - 2
                dgv.Rows.RemoveAt(i - delete_no)    '1行を削除した後に、行番号を1個増えるから
                delete_no = delete_no + 1
            Next
        End If

    End Sub

    ''' <summary>
    ''' 未出荷確定の商品数を取得する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetNotInputItems() As Integer
        Dim items As Integer
        Dim dgv As New DataGridView

        dgv = MainFunction.OutputKakutei_DataGridView
        items = dgv.Rows.Count

        Return (items)

    End Function

    ''' <summary>
    ''' Data Grid Viewのデータを設定する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetDataGridView()
        Dim dgv As New DataGridView
        Dim goods_out_sch As New OutputGoodsSchedule
        Dim goods_out_sch_data_path As String = ""
        Dim strFile As String      '出荷予定ファイル 

        'Excelファイル関連
        Dim app As Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet
        Dim row_no As Integer

        '商品の情報
        Dim product_day As String = ""      '出荷予定日
        Dim product_customer_name As String = "" '届先情報
        Dim product_name As String = ""     '品名
        Dim product_code As String = ""     '品番
        Dim product_slot As String = ""     '出数
        Dim product_out As String = ""  '出庫数

        '初期化
        dgv = MainFunction.OutputKakutei_DataGridView
        goods_out_sch_data_path = goods_out_sch.WriteDataPath()
        strFile = goods_out_sch_data_path + "\出荷予定.xlsx"

        app = CreateObject("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False

        'File Open

        If IO.File.Exists(strFile) Then 'Fileが存在する
            book = app.Workbooks.Open(strFile)
        Else
            MsgBox("File:" + strFile + "が存在しない")
            Exit Sub
        End If


        sheet = book.Worksheets(1)

        For row_no = OUTGOODS_START_ROW To OUTGOODS_MAX_ROW
            If Len(sheet.Cells(row_no, OUTGOODS_READFILE_PRODUCT_OUTPUT_COL_NO).Value) > 0 Then
                '日付
                product_day = sheet.Cells(row_no, OUTGOODS_READFILE_DAY_COL_NO).Value

                '届先情報
                product_customer_name = sheet.Cells(row_no, OUTGOODS_READFILE_CUSTOMER_COL_NO).Value

                '品番
                product_code = sheet.Cells(row_no, OUTGOODS_READFILE_CODE_COL_NO).Value

                '品名
                product_name = sheet.Cells(row_no, OUTGOODS_READFILE_NAME_COL_NO).Value

                '出数
                product_slot = sheet.Cells(row_no, OUTGOODS_READFILE_SLOT_COL_NO).Value

                '出庫数
                product_out = sheet.Cells(row_no, OUTGOODS_READFILE_PRODUCT_OUTPUT_COL_NO).Value

                dgv.Rows.Add(False, product_customer_name, product_day, product_code, product_name, product_slot, product_out)

            Else
                Exit For

            End If

        Next

        'File Close
        book.Close()

        app.Quit()

        ' オブジェクトを解放します。
        sheet = Nothing
        book = Nothing
        app = Nothing
    End Sub

    ''' <summary>
    ''' 出力の日付からYearの値を取得する
    ''' </summary>
    ''' <param name="date_str">日付の形式は2014/09/15</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetYearValue(ByVal date_str As String) As String
        Dim date_split As String() = date_str.Split("/")

        Return date_split(0)

    End Function

    ''' <summary>
    ''' 出力の日付からMonthの値を取得する
    ''' </summary>
    ''' <param name="date_str">日付の形式は2014/09/15</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetMonthValue(ByVal date_str As String) As String
        Dim date_split As String() = date_str.Split("/")

        Return date_split(1)

    End Function

    Private Function GetSyukaFileName(ByVal customer_name As String, ByVal year As String, ByVal month As String) As String
        Dim file_path As String = ""    'ファイルの絶対パス
        Dim folder_path As String = ""  'フォルダパス
        Dim template_filename As String = ""

        folder_path = GetFolderNameSaveOutputGoods(year, month)

        file_path = folder_path + year + month + "\出荷" + year + month + ".xlsx"


        'ファイルの存在確認
        If System.IO.File.Exists(file_path) Then
            'ファイルがすでに存在する場合
            'Do Nothing
        Else
            '新規ファイルを作成する
            template_filename = DataPathDefinition.GetTemplateDataPath() + "\" + OUTGOODS_TEMPLATE_FILENAME

            If IO.File.Exists(template_filename) Then 'Fileが存在する
                IO.File.Copy(template_filename, file_path)
            Else
                MsgBox("File:" + template_filename + "が存在しない")
            End If

        End If

        Return file_path


    End Function

    ''' <summary>
    ''' 在庫情報のフォルダパス
    ''' </summary>
    ''' <param name="year"></param>
    ''' <param name="month"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetFolderNameSaveOutputGoods(ByVal year As String, ByVal month As String) As String
        Dim folder_path As String = ""

        folder_path = DataPathDefinition.GetProductDataPath()

        'yearの情報を加える
        folder_path = folder_path + "\" + year + "\"

        Return folder_path

    End Function

    ''' <summary>
    ''' 先月の在庫のファイルパス取得
    ''' </summary>
    ''' <param name="current_year">現在の年</param>
    ''' <param name="current_month">現在の月</param>
    ''' <param name="current_month_save_filename">現在の保存ファイル名</param>
    ''' <returns>先月の在庫のファイルパス取得</returns>
    ''' <remarks>ファイルが存在しない場合、""を返す</remarks>
    Private Function GetLastMonthSaveFileName(ByVal current_year As String, ByVal current_month As String, ByVal current_month_save_filename As String) As String
        Dim filepath As String = ""

        Dim current_folder_name As String = ""  '現在のフォルダ名
        Dim current_file_name As String = ""    '現在のファイル名(日付の部分のみ)

        Dim pre_year As String = ""             '先月のYear
        Dim pre_month As String = ""            '先月のMonth
        Dim pre_folder_name As String = ""      '先月のフォルダ名
        Dim pre_file_name As String = ""        '先月のファイル名(日付の部分のみ)

        '出荷日からYear, Monthの情報を取得
        current_folder_name = current_year + current_month
        current_file_name = current_year + current_month

        If current_month = "01" Then
            pre_year = Integer.Parse(current_year) - 1
            pre_year = pre_year.ToString
            pre_month = "12"

        Else
            pre_year = current_year
            pre_month = Integer.Parse(current_month) - 1
            If pre_month < 10 Then
                pre_month = pre_month.ToString
                pre_month = "0" + pre_month
            Else
                'Do Nothin

            End If

        End If

        '先月のファイル設定
        pre_folder_name = pre_year + pre_month
        pre_file_name = pre_year + pre_month

        filepath = current_month_save_filename
        filepath = filepath.Replace(current_folder_name, pre_folder_name)
        filepath = filepath.Replace(current_file_name, pre_file_name)

        'Fileの存在確認
        If System.IO.File.Exists(filepath) Then 'Fileが存在する
            'Do Nothing
        Else
            filepath = ""
        End If

        Return filepath

    End Function

    ''' <summary>
    ''' 先月の在庫ファイルから今月の在庫ファイルにデータをコピーする
    ''' コピーする項目としては、以下の通り
    ''' 品番、品名、月初在庫
    ''' </summary>
    ''' <param name="lastmonthfilepath">先月の在庫ファイル名</param>
    ''' <param name="currentmonthfilepath">今月の在庫ファイル名</param>
    ''' <remarks></remarks>
    Private Sub CopyFromLastMonthInventory(ByVal lastmonthfilepath As String, ByVal currentmonthfilepath As String)
        Dim app_last_month As Excel.Application
        Dim book_last_month As Excel.Workbook
        Dim sheet_last_month As Excel.Worksheet

        Dim app_current_month As Excel.Application
        Dim book_current_month As Excel.Workbook
        Dim sheet_current_month As Excel.Worksheet

        Dim row_no As Integer = 0

        'Fileの存在確認
        If System.IO.File.Exists(lastmonthfilepath) Then
            'Do Nothing
        Else
            MsgBox(lastmonthfilepath + "が存在しない")
            Exit Sub
        End If

        If System.IO.File.Exists(currentmonthfilepath) Then
            'Do Nothing
        Else
            MsgBox(currentmonthfilepath + "が存在しない")
            Exit Sub
        End If

        'Create
        app_last_month = CreateObject("Excel.Application")
        app_current_month = CreateObject("Excel.Application")

        'FileOpen
        If IO.File.Exists(lastmonthfilepath) Then 'Fileが存在する
            book_last_month = app_last_month.Workbooks.Open(lastmonthfilepath)
        Else
            MsgBox("File:" + lastmonthfilepath + "が存在しない")
            Exit Sub
        End If


        sheet_last_month = book_last_month.Worksheets(1)

        If IO.File.Exists(currentmonthfilepath) Then 'Fileが存在する
            book_last_month = app_last_month.Workbooks.Open(currentmonthfilepath)
        Else
            MsgBox("File:" + currentmonthfilepath + "が存在しない")
            Exit Sub
        End If

        sheet_current_month = book_current_month.Worksheets(1)

        '*****************************
        'データコピー処理開始
        '*****************************
        For row_no = OUTGOODS_START_ROW To OUTGOODS_MAX_ROW
            If Len(sheet_last_month.Cells(row_no, INVENTORY_PRODUCTS_CODE_COL_NO).Value) = 0 Then
                Exit For
            Else
                '品番のコピー
                sheet_current_month.Cells(row_no, INVENTORY_PRODUCTS_CODE_COL_NO).Value = sheet_last_month.Cells(row_no, INVENTORY_PRODUCTS_CODE_COL_NO).Value

                '品名のコピー
                sheet_current_month.Cells(row_no, INVENTORY_PRODUCTS_NAME_COL_NO).Value = sheet_last_month.Cells(row_no, INVENTORY_PRODUCTS_NAME_COL_NO).Value

                '月初在庫
                sheet_current_month.Cells(row_no, INVENTORY_PRODUCTS_MONTH_BEGIN_COL_NO).Value = sheet_last_month.Cells(row_no, INVENTORY_PRODUCTS_MONTH_END_COL_NO).Value
            End If
        Next

        'File 保存
        book_last_month.Save()
        book_current_month.Save()

        'File Close
        book_last_month.Close()
        book_current_month.Close()

        '開放
        app_last_month.Quit()
        app_current_month.Quit()

        ' オブジェクトを解放します。
        sheet_last_month = Nothing
        book_last_month = Nothing
        app_last_month = Nothing

        sheet_current_month = Nothing
        book_current_month = Nothing
        app_current_month = Nothing
    End Sub

    ''' <summary>
    ''' 出力の日付から日の値を取得する
    ''' </summary>
    ''' <param name="date_str">日付の形式は2014/09/15</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetDayValue(ByVal date_str As String) As Integer
        Dim date_split As String() = date_str.Split("/")

        Return date_split(2)

    End Function

    ''' <summary>
    ''' 在庫表のテンプレートファイルパス取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSaveTemplateFileName() As String
        Return "C:\業務管理ソフトData\Template情報\Template_在庫.xlsx"

    End Function

    ''' <summary>
    ''' 在庫情報を保存するファイル名を取得する
    ''' </summary>
    ''' <param name="customer_name">取引先</param>
    ''' <param name="year">在庫年</param>
    ''' <param name="month">在庫月</param>
    ''' <returns>在庫情報を保存するファイル名</returns>
    ''' <remarks></remarks>
    Private Function GetFileNameSaveInputWare(ByVal customer_name As String, ByVal year As String, ByVal month As String) As String
        Dim file_path As String = ""    'ファイルの絶対パス
        Dim folder_path As String = ""  'フォルダパス

        folder_path = GetFolderNameSaveOutputGoods(year, month)

        file_path = folder_path + "\" + year + month + "\" + "在庫" + year + month + ".xlsx"

        Return file_path


    End Function

#End Region
End Class
