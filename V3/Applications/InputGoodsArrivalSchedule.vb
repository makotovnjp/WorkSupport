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
'暫定対応：thanh(todo)
'*************

Public Class InputGoodsArrivalSchedule

#Region "定数定義"
    Public Const ARRSCHD_OK As Integer = 1
    Public Const ARRSCHD_ERROR As Integer = -1

    '入荷入力情報の列番号
    Private Const PRODUCTS_CODE_COL_NO As Integer = 1     '品番の列番号
    Private Const PRODUCTS_NAME_COL_NO As Integer = 2     '品名の列番号
    Private Const PRODUCTS_SLOT_COL_NO As Integer = 3     '入数の列番号
    Private Const PRODUCTS_STORAGE_COL_NO As Integer = 4   '入庫数の列番号
    Private Const PRODUCTS_UNITPRICE_COL_NO As Integer = 5   '単価の列番号
    Private Const PRODUCTS_FX_COL_NO As Integer = 6   '為替の列番号

    Private Const ACCEPTABLE_INPUT_FILE_EXTENSION As String = "xls"
    Private Const PRODUCT_TEMPLATE_FILENAME = "Template_入荷予定.xlsx"

    Private Const INVENTORY_START_ROW_NO = 2        ' 開始の行
    Private Const MAX_ROW_NO As Integer = 1000      ' 最大の行数

    'Data Grid Viewの列番号
    'Private Const DGV_DAY_COL_NO As Integer = 0      '入荷予定日
    Private Const DGV_CODE_COL_NO As Integer = 0     '品番
    Private Const DGV_NAME_COL_NO As Integer = 1     '品名
    Private Const DGV_SLOT_COL_NO As Integer = 2     '入数
    Private Const DGV_STORAGE_COL_NO As Integer = 3  '入庫数
    Private Const DGV_UNITPRICE_COL_NO As Integer = 4  '単価
    Private Const DGV_FX_COL_NO As Integer = 5  '為替

    '入荷情報の列番号
    Private Const WRITEFILE_DAY_COL_NO As Integer = 1     '入荷予定日
    Private Const WRITEFILE_SHIIRE_COL_NO As Integer = 2     '仕入れ先名
    Private Const WRITEFILE_CODE_COL_NO As Integer = 3     '品番の列番号
    Private Const WRITEFILE_NAME_COL_NO As Integer = 4     '品名の列番号
    Private Const WRITEFILE_SLOT_COL_NO As Integer = 5     '入数の列番号
    Private Const WRITEFILE_STORAGE_COL_NO As Integer = 6   '入庫数の列番号
    Private Const WRITEFILE_UNITPRICE_COL_NO As Integer = 7   '単価の列番号
    Private Const WRITEFILE_FX_COL_NO As Integer = 8   '為替の列番号

#End Region

#Region "Private変数定義"


#End Region

#Region "構造体"
    Structure ProductInfo
        Dim client_name As String      '取引先
        Dim product_day As String      '入荷予定日
        Dim product_code As String     '品番
        Dim product_name As String     '品名
        Dim product_slot As String     '入数
        Dim product_storage As String  '入庫数

    End Structure
#End Region

#Region "Public Method"

    ''' <summary>
    ''' Cancelの処理
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Cancel()
        Init()
    End Sub

    ''' <summary>
    ''' 取引情報を表示する
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ShowClienList()
        Dim client_data_path As String = ""     '取引情報を格納するFolderのパス
        Dim strFiles() As String
        Dim strFile As String
        Dim FileName As String

        '取引先リストを初期化
        InitClientList()

        '取引情報を格納するFolderのパスを設定
        client_data_path = GetClientDataPath()

        'フォルダ存在の確認
        If System.IO.Directory.Exists(client_data_path) Then
            strFiles = System.IO.Directory.GetFiles(client_data_path, "*.xlsx")

            For Each strFile In strFiles
                FileName = System.IO.Path.GetFileNameWithoutExtension(strFile)
                MainFunction.NyukaYotei_ShiIreSaki_Combox.Items.Add(FileName)
            Next
        End If

    End Sub

    '1月～9月なら01～09を返す。
    Public Function Reform_Month(ByVal mthstr As String) As String
        If Integer.Parse(mthstr) < 10 Then
            Return "0" & mthstr
        Else
            Return mthstr
        End If
    End Function

    ''' <summary>
    ''' 手動入力の処理
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
    Public Sub ManualInputSchedule(ByRef sum As Integer)
        'Excelファイル処理変数
        Dim file_client_data As String = ""
        Dim app As Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet

        Dim row_no As Integer = 0   '行番号

        'Data格納変数
        Dim product_name As String = ""     '品名
        Dim product_code As String = ""     '品番
        Dim prodcut_slot As String = ""     '入数
        Dim prodcut_uprc As String = ""     '単価
        Dim prodcut_fxfx As String = ""     '為替
        'Dim product_numb As String = ""     '数
        Dim today As DateTime
        'Dim display_today As String         '日付の表示

        'Data Grid View用の変数
        Dim dgv As New DataGridView


        '初期設定
        sum = 0
        today = DateTime.Today
        'display_today = today.Year.ToString + "/" + Reform_Month(today.Month.ToString) + "/" + today.Day.ToString
        dgv = MainFunction.NyukaYoTei_DataGridView1

        app = CreateObject("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False

        '選択した取引に対応するファイル名を取得
        file_client_data = GetSelectedFileName()


        If System.IO.File.Exists(file_client_data) Then 'Exist File
            'ファイルOpen
            If IO.File.Exists(file_client_data) Then 'Fileが存在する
                book = app.Workbooks.Open(file_client_data)
            Else
                MsgBox("File:" + file_client_data + "が存在しない")
                Exit Sub
            End If

            sheet = book.Worksheets(1)


            For row_no = 2 To MAX_ROW_NO
                If sheet.Cells(row_no, PRODUCTS_CODE_COL_NO).Value <> "" Then
                    ''ファイルのデータ取得
                    product_code = sheet.Cells(row_no, PRODUCTS_CODE_COL_NO).Value.ToString
                    product_name = sheet.Cells(row_no, PRODUCTS_NAME_COL_NO).Value.ToString
                    prodcut_slot = sheet.Cells(row_no, PRODUCTS_SLOT_COL_NO).Value.ToString
                    'prodcut_uprc = sheet.Cells(row_no, PRODUCTS_UNITPRICE_COL_NO).Value.ToString
                    'prodcut_fxfx = sheet.Cells(row_no, PRODUCTS_FX_COL_NO).Value.ToString
                    'product_numb = sheet.Cells(row_no, 4).Value
                    'DataGridViewのデータを設定する
                    dgv.Rows.Add(product_code, product_name, prodcut_slot, "", "")
                    sum += 1
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

        End If

    End Sub


    ''' <summary>
    ''' 入荷予定ファイルをリードする
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReadFile() As Integer
        Dim OpenFileDialog As New OpenFileDialog()

        'Data格納変数
        Dim product_name As String = ""     '品名
        Dim product_code As String = ""     '品番
        Dim product_slot As String = ""     '入数
        Dim product_storage As String = ""  '入庫数
        Dim prodcut_uprc As String = ""     '単価
        Dim prodcut_fxfx As String = ""     '為替
        Dim today As DateTime
        Dim display_today As String         '日付の表示

        'Excelファイル関連
        Dim row_no As Integer
        Dim app As Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet

        'DataGridView関連
        Dim dgv As New DataGridView

        '初期値設定
        dgv = MainFunction.NyukaYoTei_DataGridView1
        today = DateTime.Today
        display_today = today.Year.ToString + "/" + Reform_Month(today.Month.ToString) + "/" + today.Day.ToString

        OpenFileDialog.Title = "ファイルを選択してください。"

        ' 初期表示するディレクトリを設定する
        OpenFileDialog.InitialDirectory = "C:\"

        ' ファイルのフィルタを設定する
        OpenFileDialog.Filter = "Excel ファイル|*.xlsx;*.xls|すべてのファイル|*.*"

        ' ファイルの種類 の初期設定を 2 番目に設定する (初期値 1)
        OpenFileDialog.FilterIndex = 1

        ' ダイアログボックスを閉じる前に現在のディレクトリを復元する (初期値 False)
        OpenFileDialog.RestoreDirectory = True

        ' 複数のファイルを選択可能にする (初期値 False)
        OpenFileDialog.Multiselect = False

        ' [ヘルプ] ボタンを表示する (初期値 False)
        OpenFileDialog.ShowHelp = True

        ' [読み取り専用] チェックボックスを表示する (初期値 False)
        OpenFileDialog.ShowReadOnly = True

        ' [読み取り専用] チェックボックスをオンにする (初期値 False)
        OpenFileDialog.ReadOnlyChecked = True

        ' 存在しないファイルを指定した場合は警告を表示する (初期値 True)
        OpenFileDialog.CheckFileExists = True

        ' 存在しないパスを指定した場合は警告を表示する (初期値 True)
        OpenFileDialog.CheckPathExists = True

        ' 拡張子を指定しない場合は自動的に拡張子を付加する (初期値 True)
        OpenFileDialog.AddExtension = True

        ' 有効な Win32 ファイル名だけを受け入れるようにする (初期値 True)
        OpenFileDialog.ValidateNames = True

        If OpenFileDialog.ShowDialog() = DialogResult.OK Then


            '入力ファイルの妥当性をチェックする
            If CheckFormatInputFile(OpenFileDialog.FileName) = ARRSCHD_OK Then
                app = CreateObject("Excel.Application")
                app.Visible = False
                app.DisplayAlerts = False

                'File Open
                If IO.File.Exists(OpenFileDialog.FileName) Then 'Fileが存在する
                    book = app.Workbooks.Open(OpenFileDialog.FileName)
                Else
                    MsgBox("File:" + OpenFileDialog.FileName + "が存在しない")
                    Return ARRSCHD_ERROR
                End If

                sheet = book.Worksheets(1)

                For row_no = INVENTORY_START_ROW_NO To MAX_ROW_NO
                    If Len(sheet.Cells(row_no, PRODUCTS_STORAGE_COL_NO).Value) > 0 Then
                        '在庫数がある場合

                        'Product情報を保持する
                        product_code = sheet.Cells(row_no, PRODUCTS_CODE_COL_NO).Value.ToString
                        product_name = sheet.Cells(row_no, PRODUCTS_NAME_COL_NO).Value.ToString
                        product_slot = sheet.Cells(row_no, PRODUCTS_SLOT_COL_NO).Value
                        product_storage = sheet.Cells(row_no, PRODUCTS_STORAGE_COL_NO).Value
                        prodcut_uprc = sheet.Cells(row_no, PRODUCTS_UNITPRICE_COL_NO).Value.ToString
                        prodcut_fxfx = sheet.Cells(row_no, PRODUCTS_FX_COL_NO).Value.ToString

                        'Data Grid Viewにデータを表示させる
                        dgv.Rows.Add(display_today, product_code, product_name, product_slot, product_storage, prodcut_uprc, prodcut_fxfx)

                    End If

                Next

                'File Close
                book.Close()

                app.Quit()

                ' オブジェクトを解放します。
                sheet = Nothing
                book = Nothing
                app = Nothing

                Return ARRSCHD_OK
            Else
                Return ARRSCHD_ERROR
            End If

        Else

            Return ARRSCHD_ERROR

        End If

    End Function

    Public Function WriteData(ByVal product_no As Integer) As Integer
        Dim writefile_name As String
        Dim dgv As New DataGridView

        Dim product_day As String = ""      '入荷予定日
        Dim product_clientname As String = ""       '取引先名
        Dim product_name As String = ""     '品名
        Dim product_code As String = ""     '品番
        Dim product_slot As String = ""     '入数
        Dim product_storage As String = ""  '入庫数
        Dim product_unitprice As String = ""  '単価
        Dim product_fx As String = ""  '為替
        'Dim product_no As Integer           '項目数
        Dim i As Integer

        Dim app As Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet
        Dim row_no As Integer

        '初期化
        dgv = MainFunction.NyukaYoTei_DataGridView1
        'product_no = dgv.Rows.Count

        '入荷予定情報のファイルを作成する
        writefile_name = MakeFileToWrite()

        '入荷予定情報のファイルにデータを書き込む
        If product_no > 0 Then
            app = CreateObject("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False

            'File Open
            If IO.File.Exists(writefile_name) Then 'Fileが存在する
                book = app.Workbooks.Open(writefile_name)
            Else
                MsgBox("File:" + writefile_name + "が存在しない")
                Return ARRSCHD_ERROR
            End If

            sheet = book.Worksheets(1)

            '書き込み開始の行番号を求める
            For row_no = INVENTORY_START_ROW_NO To MAX_ROW_NO
                If Len(sheet.Cells(row_no, WRITEFILE_DAY_COL_NO).Value) = 0 Then
                    Exit For
                Else
                    'Do Nothing

                End If
            Next

            product_day = Today.Year.ToString + "/" + Reform_Month(Today.Month.ToString) + "/" + Today.Day.ToString
            product_fx = MainFunction.TextBox1.Text.ToString
            For i = 0 To (product_no - 1)
                product_code = dgv(DGV_CODE_COL_NO, i).Value.ToString
                product_name = dgv(DGV_NAME_COL_NO, i).Value.ToString
                product_slot = dgv(DGV_SLOT_COL_NO, i).Value.ToString
                product_storage = dgv(DGV_STORAGE_COL_NO, i).Value.ToString
                product_unitprice = dgv(DGV_UNITPRICE_COL_NO, i).Value.ToString
                'product_fx = dgv(DGV_FX_COL_NO, i).Value.ToString


                product_clientname = MainFunction.NyukaYotei_ShiIreSaki_Combox.SelectedItem.ToString

                If Len(product_storage) > 0 Then

                    '日付
                    sheet.Cells(row_no, WRITEFILE_DAY_COL_NO).Value = product_day

                    '仕入れ先名
                    sheet.Cells(row_no, WRITEFILE_SHIIRE_COL_NO).Value = product_clientname

                    '品番
                    sheet.Cells(row_no, WRITEFILE_CODE_COL_NO).Value = product_code

                    '品名
                    sheet.Cells(row_no, WRITEFILE_NAME_COL_NO).Value = product_name

                    '入数
                    sheet.Cells(row_no, WRITEFILE_SLOT_COL_NO).Value = product_slot

                    '入庫数
                    sheet.Cells(row_no, WRITEFILE_STORAGE_COL_NO).Value = product_storage

                    '単価
                    sheet.Cells(row_no, WRITEFILE_UNITPRICE_COL_NO).Value = product_unitprice

                    '為替
                    sheet.Cells(row_no, WRITEFILE_FX_COL_NO).Value = product_fx


                    row_no = row_no + 1

                End If

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


        End If

        'Data書き込み完了後の処理
        Init()

        Return ARRSCHD_OK
    End Function

    ''' <summary>
    ''' 入荷予定情報のフォルダパス
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function WriteDataPath() As String
        Dim path As String = ""

        path = GetGoodsSchedulePath()

        Return path

    End Function

    ''' <summary>
    ''' 入荷予定情報を削除
    ''' </summary>
    ''' <param name="product_info">Product情報</param>
    ''' <remarks></remarks>
    Public Sub DeleteScheduleData(ByVal product_info As ProductInfo)
        Dim filename As String = ""
        Dim delete_row As Integer = 0 '削除行
        Dim delete_range As String    '削除範囲の設定文字列

        'Excelの処理変数
        Dim app As Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet

        filename = GetGoodsSchedulePath() + "\入荷予定.xlsx"

        'File存在確認
        If System.IO.File.Exists(filename) Then

            app = CreateObject("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False

            'File Open
            If IO.File.Exists(filename) Then 'Fileが存在する
                book = app.Workbooks.Open(filename)
            Else
                MsgBox("File:" + filename + "が存在しない")
                Exit Sub
            End If
            sheet = book.Worksheets(1)

            For delete_row = INVENTORY_START_ROW_NO To MAX_ROW_NO
                If (sheet.Cells(delete_row, WRITEFILE_DAY_COL_NO).Value = product_info.product_day) AndAlso
                    (sheet.Cells(delete_row, WRITEFILE_CODE_COL_NO).Value = product_info.product_code) AndAlso
                    (sheet.Cells(delete_row, WRITEFILE_NAME_COL_NO).Value = product_info.product_name) AndAlso
                    (sheet.Cells(delete_row, WRITEFILE_SLOT_COL_NO).Value = product_info.product_slot) AndAlso
                    (sheet.Cells(delete_row, WRITEFILE_STORAGE_COL_NO).Value = product_info.product_storage) Then

                    Exit For

                End If
            Next

            If delete_row = MAX_ROW_NO Then
                MsgBox("回答の入荷予定データが見つからない")
            End If

            '該当行を削除
            delete_range = "A" + delete_row.ToString + ":" + "A" + delete_row.ToString
            sheet.Range(delete_range).EntireRow.Delete()


            'File Save
            book.Save()

            'File Close
            book.Close()


            app.Quit()

            ' オブジェクトを解放します。
            sheet = Nothing
            book = Nothing
            app = Nothing

        End If




    End Sub

#End Region

#Region "Private Method"

    ''' <summary>
    ''' 入荷予定の機能の初期化処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Init()

        'Data Grid Viewの初期化
        InitDgv()

    End Sub

    ''' <summary>
    ''' Data Grid Viewのデータの初期化処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitDgv()
        Dim i As Integer
        Dim delete_no As Integer = 0    '削除した回数
        Dim RowCount As Integer
        Dim dgv As New DataGridView

        dgv = MainFunction.NyukaYoTei_DataGridView1
        RowCount = dgv.Rows.Count

        If RowCount > 1 Then
            For i = 0 To RowCount - 2
                dgv.Rows.RemoveAt(i - delete_no)
                delete_no = delete_no + 1
            Next
        End If

    End Sub

    Private Sub InitClientList()
        MainFunction.NyukaYotei_ShiIreSaki_Combox.DataSource = Nothing
        MainFunction.NyukaYotei_ShiIreSaki_Combox.Items.Clear()


    End Sub

    ''' <summary>
    ''' ClientDataを格納するパスの取得
    ''' </summary>
    ''' <remarks> 暫定：本来なら、パスを設定するタブを設けるべき </remarks>
    Private Function GetClientDataPath() As String

        Return DataPathDefinition.GetShiireDataPath

    End Function

    ''' <summary>
    ''' 取引数を取得する
    ''' </summary>
    ''' <param name="client_data_path">取引の情報を格納するパス</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetNumClient(ByVal client_data_path As String) As Integer
        Dim num_client As Integer = 0

        num_client = System.IO.Directory.GetFiles(client_data_path).Length()

        Return num_client

    End Function


    ''' <summary>
    ''' 入力ファイルのフォーマットチェックする
    ''' </summary>
    ''' <param name="FilePath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CheckFormatInputFile(ByVal FilePath As String) As Integer
        Dim ret As Integer = ARRSCHD_OK
        Dim file_ext As String = ""         'File拡張子

        'File 存在チェック
        If System.IO.File.Exists(FilePath) Then
            'Fileの拡張子を確認する
            file_ext = System.IO.Path.GetExtension(FilePath)

            'xlsの文字列を含むかどうかを検索する
            If 0 <= file_ext.IndexOf(ACCEPTABLE_INPUT_FILE_EXTENSION) Then
                ret = ARRSCHD_OK
            Else
                ret = ARRSCHD_ERROR
                MsgBox("このファイルタイプをサポートしていない")

            End If
        Else
            ret = ARRSCHD_ERROR
            MsgBox("選択したファイルが存在しない")
        End If

        Return ret

    End Function

    Private Function GetSelectedFileName() As String
        Dim filename As String = ""         '取引のファイル名
        Dim client_data_path As String = "" '取引情報を格納するフォルダのパス   
        Dim client_name As String = ""      '取引名

        If MainFunction.NyukaYotei_ShiIreSaki_Combox.SelectedIndex < 0 Then
            MessageBox.Show("仕入れ先を入力ください", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            client_name = MainFunction.NyukaYotei_ShiIreSaki_Combox.SelectedItem.ToString
            client_data_path = GetClientDataPath()
            filename = client_data_path + "\" + client_name + ".xlsx"
        End If

        Return filename
    End Function


    ''' <summary>
    ''' 入荷予定情報を格納するファイルを作成する
    ''' ※ファイルがすでに存在する場合、そのファイル名を返す
    ''' </summary>
    ''' <returns>入荷予定情報を格納するファイル名</returns>
    ''' <remarks></remarks>
    Private Function MakeFileToWrite() As String
        Dim clientname As String                '取引先名
        Dim goods_schedule_path As String       '取引情報を格納パス
        Dim template_filename As String = ""
        Dim filename As String = ""

        If MainFunction.NyukaYotei_ShiIreSaki_Combox.SelectedIndex < 0 Then
            MessageBox.Show("仕入れ先を入力ください", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            clientname = MainFunction.NyukaYotei_ShiIreSaki_Combox.SelectedItem.ToString
            goods_schedule_path = GetGoodsSchedulePath()
            filename = goods_schedule_path + "\入荷予定.xlsx"

            'ファイルの存在確認
            If System.IO.File.Exists(filename) Then
                'ファイルがすでに存在する場合
                'Do Nothing
            Else
                '新規ファイルを作成する
                template_filename = "C:\業務管理ソフトData\Template情報" + "\" + PRODUCT_TEMPLATE_FILENAME

                'Templateファイルからコピーする
                If IO.File.Exists(template_filename) Then 'Fileが存在する
                    IO.File.Copy(template_filename, filename)
                Else
                    MsgBox("File:" + template_filename + "が存在しない")
                End If

            End If

        End If

        Return filename

    End Function


    ''' <summary>
    ''' 入荷情報を格納するパスの取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>暫定：本来なら、パスを設定するタブを設けるべき</remarks>
    Private Function GetGoodsSchedulePath() As String

        Return DataPathDefinition.GetProductDataPath

    End Function

#End Region
End Class

