Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

''' <summary>
''' 入荷機能を実現するクラス
''' </summary>
''' <remarks></remarks>
Public Class GoodsReceived
#Region "定数定義"

    '入荷入力のファイル名
    Private Const Nyuka_Input_FileName As String = "入荷入力.xls"

    'DataGridViewの列番号
    Private Const DGV_STORATE_SURE_COL_NO As Integer = 0        '入荷確定の列番号
    Private Const DGV_SUPPLIERNAME_COL_NO As Integer = 1        '仕入先名の列番号
    Private Const DGV_GOODSRECEIVED_DATE_COL_NO As Integer = 2  '入荷日の列番号
    Private Const DGV_PRODUCTS_CODE_COL_NO As Integer = 3       '品番の列番号
    Private Const DGV_PRODUCTS_NAME_COL_NO As Integer = 4       '品名の列番号
    Private Const DGV_PRODUCTS_SLOT_COL_NO As Integer = 5       '入数の列番号
    Private Const DGV_PRODUCTS_STORAGE_COL_NO As Integer = 6    '入庫数の列番号

    '在庫表の列番号
    Private Const INVENTORY_PRODUCTS_CODE_COL_NO As Integer = 1
    Private Const INVENTORY_PRODUCTS_NAME_COL_NO As Integer = 2
    Private Const INVENTORY_PRODUCTS_SLOT_COL_NO As Integer = 3
    Private Const INVENTORY_PRODUCTS_MONTH_BEGIN_COL_NO As Integer = 6      '月初在庫の列番号
    Private Const INVENTORY_PRODUCTS_MONTH_END_COL_NO As Integer = 40       '現在庫の列番号
    Private Const INVENTORY_PRODUCTS_STORAGE_COL_NO As Integer = 7          '今月入庫数合計
    Private Const INVENTORY_PRODUCTS_DATE_START_COL_NO As Integer = 7

    '在庫表の行番号
    Private Const INVENTORY_START_ROW_NO As Integer = 4
    Private Const MAX_ROW_NO As Integer = 1000

    '在庫情報のフォルダ名
    Private Const INVENTORY_FOLDER_NAME As String = "在庫情報"
    Private Const INVENTORY_TEMPLATE_FILE_NAME As String = "Template_在庫情報.xls"   '在庫表のTemplateファイル名

#End Region

#Region "変数宣言"

    Private l_Root_Data_Folder_Path As String = ""   'Excel ファイルのパス

    '入力データ
    Private l_category As New ArrayList             '分類
    Private l_supplier_name As New ArrayList        '仕入先名
    Private l_storage_sure As New ArrayList         '入荷確定
    Private l_goodsreceived_date As New ArrayList   '入荷日
    Private l_products_code As New ArrayList        '品番
    Private l_products_name As New ArrayList        '品名
    Private l_no_products_slot As New ArrayList     '入数
    Private l_no_products_storage As New ArrayList  '入荷
    Private l_price_without_tax As New ArrayList    '単価(税抜き)
    Private l_price_with_tax As New ArrayList       '単価(税抜き)
    Private l_rate_exchange As New ArrayList        '為替レート

#End Region

#Region "Public Method"
    'コンストラクタ
    Public Sub New()

    End Sub

    ''' <summary>
    ''' 在庫情報を更新
    ''' </summary>
    ''' <remarks>ここではExceクラスを利用せずに、Excelを操作する
    ''' 理由は現状のExcelクラスから２つのインスタンスを作成するとエラーになる  </remarks>
    Public Sub Update_Storage()
        'DataのFolderのPathを設定
        'thanh:暫定
        l_Root_Data_Folder_Path = "C:\業務管理ソフトData"

        '在庫更新
        Update_Storage_byDataGridView()

        '入力したデータを削除
        Delete_DataGridView()


    End Sub

#End Region

#Region "Private Method"

    ''' <summary>
    ''' Data Grid Viewに入力した値を消す
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Delete_DataGridView()
        Dim line_no As Integer = 0

    End Sub

    ''' <summary>
    ''' Data Grid Viewの入力かあ在庫更新　(Ver15以降の仕様)
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Update_Storage_byDataGridView()
        Dim i As Integer = 0
        Dim row_no As Integer = 4                                           '検索開始の行番号
        Dim col_no_data As Integer = INVENTORY_PRODUCTS_DATE_START_COL_NO   '日付開始の列番号
        Dim day As Integer = 0                                              '日付の列番号
        Dim current_storage_num As Integer = 0                              '同じ日にすでに入荷した数
        Dim last_month_inventoryfilepath As String = ""                     '先月の在庫表のファイルパス

        Dim inventoryfilepath As String = ""                                '在庫表ファイルのパス
        Dim inventoryfolderpath As String = ""                              '在庫表のフォルダパス

        Dim app As Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet

        'GridViewの入力よりDataListに保存する
        SetInputDataList_FromDGV()

        app = CreateObject("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False

        'Data書き込む: 品番を基準にして、データを書き込む
        For i = 0 To l_supplier_name.Count - 1

            '在庫表のファイルパスを取得
            inventoryfilepath = GetInventoryPath(l_supplier_name(i), l_goodsreceived_date(i))

            If inventoryfilepath <> "" Then '在庫表のファイルパス設定成功

                If System.IO.File.Exists(inventoryfilepath) Then    '在庫表のファイルが存在する
                    'Do Nothing
                Else    '在庫表のファイルが存在しない場合、新規ファイルを作成する
                    '在庫表のフォルダ作成をする
                    inventoryfolderpath = GetInventoryFolderPath(l_goodsreceived_date(i))
                    System.IO.Directory.CreateDirectory(inventoryfolderpath)

                    'Templateのファイルから新規の在庫表ファイルをコピーする
                    If IO.File.Exists(GetInventoryTemplatePath()) Then 'Fileが存在する
                        IO.File.Copy(GetInventoryTemplatePath(), inventoryfilepath)
                    Else
                        MsgBox("File:" + GetInventoryTemplatePath() + "が存在しない")
                    End If

                    '新規在庫表の場合、先月の在庫表から品番、品名、月初在庫を入力する
                    last_month_inventoryfilepath = GetLastMonthInventoryPath(l_goodsreceived_date(i), inventoryfilepath)

                    If last_month_inventoryfilepath <> "" Then
                        '先月の在庫表がある場合
                        '先月からのデータをコピーする
                        If IO.File.Exists(last_month_inventoryfilepath) Then 'Fileが存在する
                            CopyFromLastMonthInventory(last_month_inventoryfilepath, inventoryfilepath)
                        Else
                            MsgBox("File:" + last_month_inventoryfilepath + "が存在しない")
                        End If

                    End If

                End If

                'File Open
                If IO.File.Exists(inventoryfilepath) Then 'Fileが存在する
                    book = app.Workbooks.Open(inventoryfilepath)
                Else
                    MsgBox("File:" + inventoryfilepath + "が存在しない")
                    Exit Sub
                End If

                sheet = book.Worksheets(1)

                '日付の列番号を決定
                day = GetDayValue(l_goodsreceived_date(i))
                col_no_data = INVENTORY_PRODUCTS_DATE_START_COL_NO + day

                '同じ品番の行番号を探す。同じ品番が無い場合、新しい行を追加する
                For row_no = INVENTORY_START_ROW_NO To MAX_ROW_NO
                    If sheet.Cells(row_no, INVENTORY_PRODUCTS_CODE_COL_NO).Value = l_products_code(i) Then
                        Exit For
                    Else
                        If Len(sheet.Cells(row_no, INVENTORY_PRODUCTS_CODE_COL_NO).Value) = 0 Then
                            Exit For
                        End If

                    End If
                Next

                If l_storage_sure(i) <> "" Then
                    '入荷確定のデータを在庫表に書き込む
                    '品番
                    sheet.Cells(row_no, INVENTORY_PRODUCTS_CODE_COL_NO).Value = l_products_code(i)

                    '品名
                    sheet.Cells(row_no, INVENTORY_PRODUCTS_NAME_COL_NO).Value = l_products_name(i)

                    '入数
                    sheet.Cells(row_no, INVENTORY_PRODUCTS_SLOT_COL_NO).Value = l_no_products_slot(i)

                    '入庫
                    current_storage_num = sheet.Cells(row_no, col_no_data).Value
                    sheet.Cells(row_no, col_no_data) = current_storage_num + l_no_products_storage(i)

                    'File Save
                    book.Save()

                    'File Close
                    book.Close()

                End If ' If l_storage_sure(i) = Nyuka_KAKUTEI Then

            End If 'If inventoryfilepath <> "" Then

        Next

        app.Quit()

        ' オブジェクトを解放します。
        sheet = Nothing
        book = Nothing
        app = Nothing

        MsgBox("在庫表更新完了しました")

    End Sub

    ''' <summary>
    ''' DataGridViewより値を取得する
    ''' </summary>
    ''' <param name="Row">行番号:0からカウント</param>
    ''' <param name="Colown">列番号：1からカウント</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetCellValueFromDataGridView(ByVal Colown As Integer, ByVal Row As Integer) As String
        Dim ret As String = ""


        Return ret

    End Function

    Private Function GetLineNoFromDataGridView() As Integer
        Dim line_no As Integer = 0
        Dim i As Integer

        For i = 0 To 1000
            If GetCellValueFromDataGridView(DGV_SUPPLIERNAME_COL_NO, i) <> "" Then
                line_no = line_no + 1
            Else
                Exit For
            End If

        Next

        Return line_no
    End Function

    ''' <summary>
    ''' Data Grid Viewの入力よりData Listの保存する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetInputDataList_FromDGV()
        Dim data_no As Integer = 0
        Dim row_no As Integer = 0

        'Array Listの初期化
        l_category.Clear()
        l_supplier_name.Clear()
        l_storage_sure.Clear()
        l_goodsreceived_date.Clear()
        l_products_code.Clear()
        l_products_name.Clear()
        l_no_products_slot.Clear()
        l_no_products_storage.Clear()
        l_price_without_tax.Clear()
        l_price_with_tax.Clear()
        l_rate_exchange.Clear()

        data_no = GetLineNoFromDataGridView()

        If data_no > 0 Then
            For row_no = 0 To (data_no - 1)  '各行の処理
                'DataをListに保存する
                l_storage_sure.Add(GetCellValueFromDataGridView(DGV_STORATE_SURE_COL_NO, row_no))
                l_supplier_name.Add(GetCellValueFromDataGridView(DGV_SUPPLIERNAME_COL_NO, row_no))
                l_goodsreceived_date.Add(GetCellValueFromDataGridView(DGV_GOODSRECEIVED_DATE_COL_NO, row_no))
                l_products_code.Add(GetCellValueFromDataGridView(DGV_PRODUCTS_CODE_COL_NO, row_no))
                l_products_name.Add(GetCellValueFromDataGridView(DGV_PRODUCTS_NAME_COL_NO, row_no))
                l_no_products_slot.Add(GetCellValueFromDataGridView(DGV_PRODUCTS_SLOT_COL_NO, row_no))
                l_no_products_storage.Add(GetCellValueFromDataGridView(DGV_PRODUCTS_STORAGE_COL_NO, row_no))
            Next
        Else
            MsgBox("入力データが無い")

        End If
    End Sub


    ''' <summary>
    ''' 在庫表のフォルダパスを取得する
    ''' </summary>
    ''' <param name="goodsreceived_date"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetInventoryFolderPath(ByVal goodsreceived_date As String) As String
        Dim InventoryFolderPath As String = ""
        Dim Year As String = ""
        Dim Month As String = ""

        '入荷入力.xlsxの文字列を削除
        InventoryFolderPath = l_Root_Data_Folder_Path

        '在庫情報ファルダのパスを設定する
        InventoryFolderPath = InventoryFolderPath + INVENTORY_FOLDER_NAME

        '入荷日からYear, Monthの情報を取得
        Year = GetYearValue(goodsreceived_date)
        Month = GetMonthValue(goodsreceived_date)

        '日付対応のフォルダを設定する
        InventoryFolderPath = InventoryFolderPath + "\" + Year + "_" + Month

        Return InventoryFolderPath

    End Function

    ''' <summary>
    ''' 在庫表のパスを取得する
    ''' </summary>
    ''' <param name="supplier_name">仕入先名</param>
    ''' <param name="goodsreceived_date">入荷日</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' フォルダの構成として、次のようになる。この構成が変わると、ソースの変更が必要
    ''' 入荷入力.xlsと同じファルダに"在庫情報"のファルダがある
    ''' </remarks>
    Private Function GetInventoryPath(ByVal supplier_name As String, ByVal goodsreceived_date As String) As String
        Dim InventoryPath As String = ""
        Dim Year As String = ""
        Dim Month As String = ""

        InventoryPath = GetInventoryFolderPath(goodsreceived_date)

        '入荷日からYear, Monthの情報を取得
        Year = GetYearValue(goodsreceived_date)
        Month = GetMonthValue(goodsreceived_date)

        '入先からのパスを設定
        InventoryPath = InventoryPath + "\" + supplier_name + Year + "年" + Month + "月" + "_在庫表.xls"

        Return InventoryPath
    End Function

    ''' <summary>
    ''' 先月の在庫のファイルパス取得
    ''' </summary>
    ''' <param name="goodsreceived_date">入荷日</param>
    ''' <param name="current_month_inventoryfilepath">今月の在庫表のファイル</param>
    ''' <returns>先月の在庫のファイルパス取得</returns>
    ''' <remarks>ファイルが存在しない場合、""を返す</remarks>
    Private Function GetLastMonthInventoryPath(ByVal goodsreceived_date As String, ByVal current_month_inventoryfilepath As String) As String
        Dim filepath As String = ""

        Dim current_year As String = ""         '現在のYear
        Dim current_month As String = ""        '現在のMonth
        Dim current_folder_name As String = ""  '現在のフォルダ名
        Dim current_file_name As String = ""    '現在のファイル名(日付の部分のみ)

        Dim pre_year As String = ""             '先月のYear
        Dim pre_month As String = ""            '先月のMonth
        Dim pre_folder_name As String = ""      '先月のフォルダ名
        Dim pre_file_name As String = ""        '先月のファイル名(日付の部分のみ)

        '入荷日からYear, Monthの情報を取得
        current_year = GetYearValue(goodsreceived_date)
        current_month = GetMonthValue(goodsreceived_date)
        current_folder_name = current_year + "_" + current_month
        current_file_name = current_year + "年" + current_month + "月"

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
        pre_folder_name = pre_year + "_" + pre_month
        pre_file_name = pre_year + "年" + pre_month + "月"

        filepath = current_month_inventoryfilepath
        filepath = filepath.Replace(current_folder_name, pre_folder_name)
        filepath = filepath.Replace(current_file_name, pre_file_name)

        'Fileの存在確認
        If System.IO.File.Exists(filepath) Then 'Fileが存在する
            'Do Nothing
        Else
            MsgBox(filepath + "には先月のファイルが存在しない")
            filepath = ""
        End If

        Return filepath

    End Function

    ''' <summary>
    ''' 在庫表のテンプレートファイルパス取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetInventoryTemplatePath() As String
        Dim InventoryTemplatePath As String = ""

        '在庫情報ファルダのパスを設定する
        InventoryTemplatePath = l_Root_Data_Folder_Path + INVENTORY_FOLDER_NAME

        'Templateファイルのパスを設定
        InventoryTemplatePath = InventoryTemplatePath + "\" + INVENTORY_TEMPLATE_FILE_NAME

        'Fileの存在確認
        If System.IO.File.Exists(InventoryTemplatePath) Then 'Fileが存在
            'Do Nothing
        Else
            MsgBox(InventoryTemplatePath + "にテンプレートのファイルが存在しない")
        End If

        Return InventoryTemplatePath

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
            book_current_month = app_current_month.Workbooks.Open(currentmonthfilepath)
        Else
            MsgBox("File:" + currentmonthfilepath + "が存在しない")
            Exit Sub
        End If

        sheet_current_month = book_current_month.Worksheets(1)

        '*****************************
        'データコピー処理開始
        '*****************************
        For row_no = INVENTORY_START_ROW_NO To MAX_ROW_NO
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
    ''' 入力の日付からYearの値を取得する
    ''' </summary>
    ''' <param name="date_str">日付の形式は2014/09/15</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetYearValue(ByVal date_str As String) As String
        Dim date_split As String() = date_str.Split("/")

        Return date_split(0)

    End Function

    ''' <summary>
    ''' 入力の日付からMonthの値を取得する
    ''' </summary>
    ''' <param name="date_str">日付の形式は2014/09/15</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetMonthValue(ByVal date_str As String) As String
        Dim date_split As String() = date_str.Split("/")

        Return date_split(1)

    End Function

    ''' <summary>
    ''' 入力の日付から日の値を取得する
    ''' </summary>
    ''' <param name="date_str">日付の形式は2014/09/15</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetDayValue(ByVal date_str As String) As Integer
        Dim date_split As String() = date_str.Split("/")

        date_split(2) = date_split(2).Replace(" 0:00:00", "")

        Return date_split(2)

    End Function

    ''' <summary>
    ''' 入荷入力データの妥当性をチェックする
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CheckInputData()
        ' 仕入れ先名の入力チェック
        CheckInputSupplierName()

        ' 入荷日の入力チェック
        CheckInputDate()

        ' 品番の入力をチェックする
        CheckInputProductCode()

        ' 品名の入力をチェックする
        CheckInputProductName()

        ' 入数の入力をチェックする
        CheckInputSlot()

        ' 入庫数の入力をチェックする
        CheckInputStorage()

    End Sub

    ''' <summary>
    ''' 仕入れ先名の入力チェック
    ''' True : チェック成功
    ''' False:　チェックNG
    ''' </summary>
    ''' <remarks>暫定：チェック内容無し</remarks>
    Private Function CheckInputSupplierName() As Boolean
        Dim ret As Boolean = True

        Return ret

    End Function

    ''' <summary>
    ''' 入荷日の入力チェック
    ''' True : チェック成功
    ''' False: チェックNG
    ''' </summary>
    ''' <remarks></remarks>
    Private Function CheckInputDate() As Boolean
        Dim ret As Boolean = True

        Return ret
    End Function

    ''' <summary>
    ''' 品番の入力をチェックする
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CheckInputProductCode()

    End Sub

    ''' <summary>
    ''' 品名の入力をチェックする
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CheckInputProductName()

    End Sub

    ''' <summary>
    ''' 入数の入力をチェックする
    ''' 1. 数値であるかどうかをチェックする
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CheckInputSlot()

    End Sub

    ''' <summary>
    ''' 入庫数の入力をチェックする
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CheckInputStorage()

    End Sub


#End Region

End Class

