Imports Microsoft.Office.Interop

'入荷予定機能を実現するクラス

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

    '扱うFile Type (Extension)
    Private Const FILE_EXTENSION As String = "xlsx"

    '仕入れ先の情報を記載するシート番号
    Private Const CLIENT_FILE_SHEET_NO As Integer = 1

    '仕入れ先の情報のFormat
    Private Const CLIENT_FILE_FORMAT_DEF_ROW As Integer = 1 '各列の情報を定義する行
    Private Const CLIENT_FILE_CODE_COL_NAME As String = "品番"
    Private Const CLIENT_FILE_NAME_COL_NAME As String = "品名"
    Private Const CLIENT_FILE_SLOT_COL_NAME As String = "入数"


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
    ''' 取引の名前のリストを取得する
    ''' </summary>
    ''' <remarks></remarks>
    Public Function GetClienList() As List(Of String)
        Dim client_data_path As String = ""     '取引情報を格納するFolderのパス
        Dim strFiles() As String
        Dim strFile As String
        Dim FileName As String

        Dim clientList As New List(Of String)

        '取引先リストを初期化
        InitClientList()

        '取引情報を格納するFolderのパスを設定
        client_data_path = GetClientDataPath()

        'フォルダ存在の確認
        If IO.Directory.Exists(client_data_path) Then
            Try
                strFiles = IO.Directory.GetFiles(client_data_path, "*." + FILE_EXTENSION)

                If strFiles.Count > 0 Then
                    For Each strFile In strFiles
                        FileName = IO.Path.GetFileNameWithoutExtension(strFile)
                        clientList.Add(FileName)
                    Next
                Else
                    MsgBox("フォルダ:" + client_data_path + "の内に取引情報のファイルが存在しない")
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        Else
            MsgBox("フォルダ:" + client_data_path + "が存在しない")
        End If

        Return clientList

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
        Dim product_slot As String = ""     '入数
        Dim product_uprc As String = ""     '単価
        Dim product_fxfx As String = ""     '為替
        'Dim product_numb As String = ""     '数
        Dim today As DateTime
        'Dim display_today As String         '日付の表示

        'Data Grid View用の変数
        Dim dgv As New DataGridView

        '初期設定
        sum = 0
        today = DateTime.Today
        'display_today = today.Year.ToString + "/" + Microsoft.VisualBasic.Right("0" & Today.Month.ToString, 2) + "/" + today.Day.ToString
        dgv = MainFunction.NyukaYoTei_DataGridView1

        Try
            app = CreateObject("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try

        '選択した取引に対応するファイル名を取得
        file_client_data = GetSelectedFileName()

        If IO.File.Exists(file_client_data) Then 'Exist File
            'ファイルOpen
            Try
                book = app.Workbooks.Open(file_client_data)
            Catch ex As Exception
                MsgBox(ex.Message)
                Exit Sub
            End Try

            '有効なClient Fileである場合
            If IsValidClientFile(book) Then
                sheet = book.Worksheets(CLIENT_FILE_SHEET_NO)
                row_no = INVENTORY_START_ROW_NO
                Do While sheet.Cells(row_no, PRODUCTS_CODE_COL_NO).Value <> ""
                    ''ファイルのデータ取得
                    If Len(sheet.Cells(row_no, PRODUCTS_CODE_COL_NO).Value) = 0 Or
                        Len(sheet.Cells(row_no, PRODUCTS_NAME_COL_NO).Value) = 0 Or
                        Len(sheet.Cells(row_no, PRODUCTS_SLOT_COL_NO).Value) = 0 Then
                        MsgBox(file_client_data + "の行: " + row_no.ToString + "には不正なデータがあるため、システムが終了する")
                        Exit Do
                    End If

                    product_code = sheet.Cells(row_no, PRODUCTS_CODE_COL_NO).Value.ToString
                    product_name = sheet.Cells(row_no, PRODUCTS_NAME_COL_NO).Value.ToString
                    product_slot = sheet.Cells(row_no, PRODUCTS_SLOT_COL_NO).Value.ToString
                    'product_uprc = sheet.Cells(row_no, PRODUCTS_UNITPRICE_COL_NO).Value.ToString
                    'product_fxfx = sheet.Cells(row_no, PRODUCTS_FX_COL_NO).Value.ToString
                    'product_numb = sheet.Cells(row_no, 4).Value

                    'DataGridViewのデータを設定する
                    dgv.Rows.Add(product_code, product_name, product_slot, "", "")
                    sum += 1
                    row_no += 1
                Loop
            Else
                MsgBox(file_client_data + "は有効なファイルではない。ファイルのフォーマットを確認ください!")
            End If

            book.Close()
            app.Quit()

            ' オブジェクトを解放します。
            sheet = Nothing
            book = Nothing
            app = Nothing
        Else
            MsgBox("File:" + file_client_data + "が存在しない")
        End If

    End Sub

    ''' <summary>
    ''' 入荷予定ファイルをリードする
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReadFile() As Integer
        Dim fileOpenner As New FileOpener
        Dim openFileName As String = ""

        'Data格納変数
        Dim product_name As String = ""     '品名
        Dim product_code As String = ""     '品番
        Dim product_slot As String = ""     '入数
        Dim product_storage As String = ""  '入庫数
        Dim product_uprc As String = ""     '単価
        Dim product_fxfx As String = ""     '為替
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
        display_today = today.Year.ToString + "/" + Microsoft.VisualBasic.Right("0" & today.Month.ToString, 2) + "/" + today.Day.ToString

        openFileName = fileOpenner.GetFileFromDialog()

        If openFileName <> "" Then
            '入力ファイルの妥当性をチェックする
            If CheckFormatInputFile(openFileName) = ARRSCHD_OK Then

                Try
                    app = CreateObject("Excel.Application")
                    app.Visible = False
                    app.DisplayAlerts = False
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Return ARRSCHD_ERROR
                End Try

                'File Open
                If IO.File.Exists(openFileName) Then 'Fileが存在する
                    Try
                        book = app.Workbooks.Open(openFileName)
                    Catch ex As Exception
                        MsgBox(ex.Message)
                        Return ARRSCHD_ERROR
                    End Try
                Else
                    MsgBox("File:" + openFileName + "が存在しない")
                    Return ARRSCHD_ERROR
                End If

                sheet = book.Worksheets(CLIENT_FILE_SHEET_NO)

                '有効なBook
                If IsValidClientFile(book) Then
                    row_no = INVENTORY_START_ROW_NO

                    Do While Len(sheet.Cells(row_no, PRODUCTS_STORAGE_COL_NO).Value) > 0
                        'Product情報を保持する

                        'Sheetの値の妥当性を確認する
                        If Len(sheet.Cells(row_no, PRODUCTS_CODE_COL_NO).Value) = 0 Or
                           Len(sheet.Cells(row_no, PRODUCTS_NAME_COL_NO).Value) = 0 Or
                           Len(sheet.Cells(row_no, PRODUCTS_SLOT_COL_NO).Value) = 0 Or
                           Len(sheet.Cells(row_no, PRODUCTS_STORAGE_COL_NO).Value) = 0 Or
                           Len(sheet.Cells(row_no, PRODUCTS_UNITPRICE_COL_NO).Value) = 0 Then
                            MsgBox(openFileName + "の行: " + row_no.ToString + "には不正なデータがあるため、システムが終了する")
                            Exit Do
                        End If

                        product_code = sheet.Cells(row_no, PRODUCTS_CODE_COL_NO).Value.ToString
                        product_name = sheet.Cells(row_no, PRODUCTS_NAME_COL_NO).Value.ToString
                        product_slot = sheet.Cells(row_no, PRODUCTS_SLOT_COL_NO).Value
                        product_storage = sheet.Cells(row_no, PRODUCTS_STORAGE_COL_NO).Value
                        product_uprc = sheet.Cells(row_no, PRODUCTS_UNITPRICE_COL_NO).Value.ToString
                        'product_fxfx = sheet.Cells(row_no, PRODUCTS_FX_COL_NO).Value.ToString

                        'Data Grid Viewにデータを表示させる
                        dgv.Rows.Add(product_code, product_name, product_slot, product_storage, product_uprc)

                        row_no += 1
                    Loop

                End If

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

    ''' <summary>
    ''' 入力したデータを入荷ファイルにWriteする
    ''' </summary>
    ''' <param name="product_no"></param>
    ''' <returns></returns>
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
        Dim i As Integer

        Dim app As Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet
        Dim row_no As Integer

        '初期化
        dgv = MainFunction.NyukaYoTei_DataGridView1

        '入荷予定情報のファイルを作成する
        writefile_name = MakeFileToWrite()

        '入荷予定情報のファイルにデータを書き込む
        If product_no > 0 Then
            Try
                app = CreateObject("Excel.Application")
                app.Visible = False
                app.DisplayAlerts = False
            Catch ex As Exception
                MsgBox(ex.Message)
                Return ARRSCHD_ERROR
            End Try

            'File Open
            If IO.File.Exists(writefile_name) Then 'Fileが存在する
                Try
                    book = app.Workbooks.Open(writefile_name)
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Return ARRSCHD_ERROR
                End Try
            Else
                MsgBox("File:" + writefile_name + "が存在しない")
                Return ARRSCHD_ERROR
            End If

            sheet = book.Worksheets(1)

            '書き込み開始の行番号を求める
            row_no = INVENTORY_START_ROW_NO
            Do While Len(sheet.Cells(row_no, WRITEFILE_DAY_COL_NO).Value) > 0
                row_no = row_no + 1
            Loop

            product_day = Today.Year.ToString + "/" + Microsoft.VisualBasic.Right("0" & Today.Month.ToString, 2) + "/" + Today.Day.ToString
            product_fx = MainFunction.TextBox1.Text.ToString

            'データ書き込み部
            For i = 0 To (product_no - 1)
                product_code = dgv(DGV_CODE_COL_NO, i).Value.ToString
                product_name = dgv(DGV_NAME_COL_NO, i).Value.ToString
                product_slot = dgv(DGV_SLOT_COL_NO, i).Value.ToString
                product_storage = dgv(DGV_STORAGE_COL_NO, i).Value.ToString
                product_unitprice = dgv(DGV_UNITPRICE_COL_NO, i).Value.ToString

                product_clientname = MainFunction.NyukaYotei_ShiIreSaki_Combox.SelectedItem.ToString

                If Len(product_storage) And Len(product_unitprice) > 0 Then

                    If IsNumeric(product_storage) = False Then
                        MsgBox((i + 1).ToString & "行目の入庫数の値は数値ではない")
                        GoTo myError
                    End If

                    If IsNumeric(product_unitprice) = False Then
                        MsgBox((i + 1).ToString & "行目の単価の値は数値ではない")
                        GoTo myError
                    End If

                    If Len(product_fx) > 0 Then
                        If IsNumeric(product_fx) = False Then
                            MsgBox((i + 1).ToString & "為替の値は数値ではない")
                            GoTo myError
                        End If
                    End If


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

myError:
        'File Close
        book.Close()

        app.Quit()

        ' オブジェクトを解放します。
        sheet = Nothing
        book = Nothing
        app = Nothing
        Return ARRSCHD_ERROR

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
        If IO.File.Exists(FilePath) Then
            'Fileの拡張子を確認する
            file_ext = System.IO.Path.GetExtension(FilePath)

            'xlsxの文字列を含むかどうかを検索する
            If 0 <= file_ext.IndexOf(FILE_EXTENSION) Then
                ret = ARRSCHD_OK
            Else
                ret = ARRSCHD_ERROR
                MsgBox("選択したファイルはサポートしていないです。")
            End If
        Else
            ret = ARRSCHD_ERROR
            MsgBox("選択したファイルが存在しない")
        End If

        Return ret

    End Function

    ''' <summary>
    ''' 選択した取引の情報を保持するFile名を取得する
    ''' </summary>
    ''' <returns></returns>
    Private Function GetSelectedFileName() As String
        Dim filename As String = ""         '取引のファイル名
        Dim client_data_path As String = "" '取引情報を格納するフォルダのパス   
        Dim client_name As String = ""      '取引名

        If MainFunction.NyukaYotei_ShiIreSaki_Combox.SelectedIndex < 0 Then
            MessageBox.Show("仕入れ先を入力ください", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            client_name = MainFunction.NyukaYotei_ShiIreSaki_Combox.SelectedItem.ToString
            client_data_path = GetClientDataPath()
            filename = client_data_path + "\" + client_name + "." + FILE_EXTENSION
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
        Dim template_filename As String = ""
        Dim filename As String = ""

        If MainFunction.NyukaYotei_ShiIreSaki_Combox.SelectedIndex < 0 Then
            MessageBox.Show("仕入れ先を入力ください", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            filename = DataPathDefinition.GetNyukaYoteiFilePath()

            'ファイルの存在確認
            If IO.File.Exists(filename) Then
                'ファイルがすでに存在する場合
                'Do Nothing
            Else
                '新規ファイルを作成する
                template_filename = DataPathDefinition.GetTemplateNyukaPath()

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

    ''' <summary>
    ''' 有効なClient Fileなのかどうかをチェックする
    ''' </summary>
    ''' <param name="book"></param>
    ''' <returns></returns>
    Private Function IsValidClientFile(ByRef book As Excel.Workbook) As Boolean
        Dim result As Boolean = True
        Dim sheet As Excel.Worksheet

        sheet = book.Worksheets(CLIENT_FILE_SHEET_NO)

        'Client FileのFormatのチェック
        If sheet.Cells(CLIENT_FILE_FORMAT_DEF_ROW, PRODUCTS_CODE_COL_NO).Value <> CLIENT_FILE_CODE_COL_NAME Or
           sheet.Cells(CLIENT_FILE_FORMAT_DEF_ROW, PRODUCTS_NAME_COL_NO).Value <> CLIENT_FILE_NAME_COL_NAME Or
           sheet.Cells(CLIENT_FILE_FORMAT_DEF_ROW, PRODUCTS_SLOT_COL_NO).Value <> CLIENT_FILE_SLOT_COL_NAME Then
            result = False
        End If

        Return result
    End Function

#End Region
End Class

