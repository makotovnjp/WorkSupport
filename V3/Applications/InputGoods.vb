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
'Public変数　 : INGOODS_
'Private変数  : l_

'*************
'暫定対応：thanh(todo)
'*************

Public Class InputGoods

#Region "データ部"
    Public Const INGOODS_OK As Integer = 1
    Public Const INGOODS_NG As Integer = (-1)

    Private Const START_ROW As Integer = 2
    Private Const MAX_ROW As Integer = 1000

    '入荷情報の列番号
    Private Const READFILE_DAY_COL_NO As Integer = 1     '入荷予定日
    Private Const READFILE_SHIIRE_COL_NO As Integer = 2  '仕入れ先名
    Private Const READFILE_CODE_COL_NO As Integer = 3     '品番の列番号
    Private Const READFILE_NAME_COL_NO As Integer = 4     '品名の列番号
    Private Const READFILE_SLOT_COL_NO As Integer = 5     '入数の列番号
    Private Const READFILE_STORAGE_COL_NO As Integer = 6   '入庫数の列番号
    Private Const READFILE_UNITPRICE_COL_NO As Integer = 7   '単価の列番号
    Private Const READFILE_FX_COL_NO As Integer = 8   '為替の列番号

    'DataGridViewの列番号
    Private Const DGV_CLIENT_NAME As Integer = 1     '入荷予定日
    Private Const DGV_DAY_COL_NO As Integer = 2     '入荷予定日
    Private Const DGV_CODE_COL_NO As Integer = 3     '品番の列番号
    Private Const DGV_NAME_COL_NO As Integer = 4    '品名の列番号
    Private Const DGV_SLOT_COL_NO As Integer = 5     '入数の列番号
    Private Const DGV_STORAGE_COL_NO As Integer = 6   '入庫数の列番号
    Private Const DGV_UNITPRICE_COL_NO As Integer = 7   '単価の列番号
    Private Const DGV_FX_COL_NO As Integer = 8   '為替の列番号

    '在庫表の列番号
    Private Const INVENTORY_PRODUCTS_CODE_COL_NO As Integer = 1
    Private Const INVENTORY_PRODUCTS_NAME_COL_NO As Integer = 2
    Private Const INVENTORY_PRODUCTS_SLOT_COL_NO As Integer = 3
    Private Const INVENTORY_PRODUCTS_MONTH_BEGIN_COL_NO As Integer = 6      '月初在庫の列番号
    Private Const INVENTORY_PRODUCTS_MONTH_END_COL_NO As Integer = 40       '現在庫の列番号
    Private Const INVENTORY_PRODUCTS_STORAGE_COL_NO As Integer = 7          '今月入庫数合計
    Private Const INVENTORY_PRODUCTS_DATE_START_COL_NO As Integer = 7
    Private Const INVENTORY_START_ROW_NO As Integer = 4

    '入荷ファイルの列番号
    Private Const NYUKA_DAY_COL_NO As Integer = 1     '入荷予定日
    Private Const NYUKA_PRODUCTS_SHIIRE_COL_NO As Integer = 2
    Private Const NYUKA_PRODUCTS_CODE_COL_NO As Integer = 3
    Private Const NYUKA_PRODUCTS_NAME_COL_NO As Integer = 4
    Private Const NYUKA_PRODUCTS_SLOT_COL_NO As Integer = 5
    Private Const NYUKA_PRODUCTS_NYUKO_COL_NO As Integer = 6
    Private Const NYUKA_PRODUCTS_UNITPRICE_COL_NO As Integer = 7
    Private Const NYUKA_PRODUCTS_FX_COL_NO As Integer = 8

#End Region

#Region "プロパティ部"


#End Region


#Region "Public Method"

    ''' <summary>
    ''' 入荷確定の初期化
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Init()
        'DataGridViewのデータを初期化
        InitDataGridView()

        '読み込んだデータをDataGridViewにセットする
        SetDataGridView()

        '未入荷項目数を表示する
        MainFunction.Num_NotInputs.Text = (GetNotInputItems() - 1).ToString

    End Sub

    ''' <summary>
    ''' Data Grid Viewのチェックボックスを設定する
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SetCheckBox(ByVal Status As Boolean)
        Dim i As Integer
        Dim RowCount As Integer
        Dim dgv As New DataGridView

        dgv = MainFunction.NyukaKakutei_DataGridView

        RowCount = GetNotInputItems()

        If RowCount > 1 Then
            For i = 0 To RowCount - 2
                dgv(0, i).Value = Status
            Next
        End If

    End Sub

    '入荷情報を更新する
    Public Function InputNyukaFile() As Integer
        Dim i As Integer
        Dim delete_no As Integer = 0    '削除した回数
        Dim RowCount As Integer
        Dim dgv As New DataGridView

        Dim CInputSchedule As New InputGoodsArrivalSchedule

        Dim row_no As Integer = 2                                           '検索開始の行番号
        Dim day As Integer = 0                                              '日付の列番号
        Dim current_storage_num As Integer = 0                              '同じ日にすでに入荷した数

        '商品の情報
        Dim client_name As String = ""      '取引先
        Dim product_day As String = ""      '入荷予定日
        Dim product_code As String = ""     '品番
        Dim product_name As String = ""     '品名
        Dim product_slot As String = ""     '入数
        Dim product_storage As String = ""  '入庫数
        Dim product_unitprice As String = ""  '単価
        Dim product_fx As String = ""  '為替

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
        'Dim book_zaikin As Excel.Workbook
        'Dim sheet_zaikin As Excel.Worksheet

        dgv = MainFunction.NyukaKakutei_DataGridView
        RowCount = GetNotInputItems()

        If RowCount > 1 Then
            Try
                app = CreateObject("Excel.Application")
                app.Visible = False
                app.DisplayAlerts = False
            Catch ex As Exception
                Return INGOODS_NG
            End Try

            For i = 0 To RowCount - 2
                '入庫確定
                If dgv(0, i - delete_no).Value = True Then   '1行を削除した後に、行番号を1個増えるから
                    '入庫情報を設定する
                    If Len(dgv(DGV_CLIENT_NAME, i - delete_no).Value) > 0 Then
                        client_name = dgv(DGV_CLIENT_NAME, i - delete_no).Value.ToString
                    End If

                    If Len(dgv(DGV_DAY_COL_NO, i - delete_no).Value) > 0 Then
                        product_day = dgv(DGV_DAY_COL_NO, i - delete_no).Value.ToString
                    End If

                    If Len(dgv(DGV_CODE_COL_NO, i - delete_no).Value) > 0 Then
                        product_code = dgv(DGV_CODE_COL_NO, i - delete_no).Value.ToString
                    End If

                    If Len(dgv(DGV_NAME_COL_NO, i - delete_no).Value) > 0 Then
                        product_name = dgv(DGV_NAME_COL_NO, i - delete_no).Value.ToString
                    End If

                    If Len(dgv(DGV_SLOT_COL_NO, i - delete_no).Value) > 0 Then
                        product_slot = dgv(DGV_SLOT_COL_NO, i - delete_no).Value.ToString
                    End If

                    If Len(dgv(DGV_STORAGE_COL_NO, i - delete_no).Value) > 0 Then
                        product_storage = dgv(DGV_STORAGE_COL_NO, i - delete_no).Value.ToString
                    End If

                    If Len(dgv(DGV_UNITPRICE_COL_NO, i - delete_no).Value) > 0 Then
                        product_unitprice = dgv(DGV_UNITPRICE_COL_NO, i - delete_no).Value.ToString
                    End If

                    If (Len(dgv(DGV_FX_COL_NO, i - delete_no).Value) > 0) Then
                        product_fx = dgv(DGV_FX_COL_NO, i - delete_no).Value.ToString
                    End If

                    '入庫情報を保存するファイルを探す
                    year_value = GetYearValue(product_day)
                    month_value = GetMonthValue(product_day)
                        save_filename = GetNyukaFileName(client_name, year_value, month_value)

                        If IO.File.Exists(save_filename) Then    '在庫表のファイルが存在する
                            'Do Nothing
                        Else    '在庫表のファイルが存在しない場合、新規ファイルを作成する
                            save_foldername = GetFolderNameSaveInputGoods(year_value, month_value)

                            'フォルダ作成
                            IO.Directory.CreateDirectory(save_foldername)

                            'Templateのファイルから新規の在庫表ファイルをコピーする
                            '本来なら、このルートには入らないはずですが、バカ避けの意味で処理を入れている
                            If IO.File.Exists(GetSaveTemplateFileName) Then 'Fileが存在する
                                IO.File.Copy(GetSaveTemplateFileName, save_filename)
                            Else
                                MsgBox("File:" + GetSaveTemplateFileName() + "が存在しない")
                            End If

                            last_month_save_filename = GetLastMonthSaveFileName(year_value, month_value, save_filename)
                            '先月のファイル存在確認
                            If last_month_save_filename <> "" Then
                                'ファイルが存在する場合
                                '先月の在庫情報からコピーする
                                CopyFromLastMonthInventory(last_month_save_filename, save_filename)
                            Else
                                'ファイルが存在しない場合: Do Nothing

                            End If

                        End If

                        'File Open
                        If IO.File.Exists(save_filename) Then 'Fileが存在する
                            book = app.Workbooks.Open(save_filename)
                        Else
                            MsgBox("File:" + save_filename + "が存在しない")
                            Return INGOODS_NG
                        End If

                        sheet = book.Worksheets(1)

                        '日付の列番号を決定
                        day_value = GetDayValue(product_day)

                        '同じ品番の行番号を探す。同じ品番が無い場合、新しい行を追加する

                        row_no = START_ROW

                        Do While Len(sheet.Cells(row_no, NYUKA_PRODUCTS_SHIIRE_COL_NO).Value) > 0
                            row_no = row_no + 1
                        Loop

                        '入荷確定のデータを在庫表に書き込む
                        '入荷確定日付
                        sheet.Cells(row_no, NYUKA_DAY_COL_NO).Value = product_day

                        '仕入れ先
                        sheet.Cells(row_no, NYUKA_PRODUCTS_SHIIRE_COL_NO).Value = client_name

                        '品番
                        sheet.Cells(row_no, NYUKA_PRODUCTS_CODE_COL_NO).Value = product_code

                        '品名
                        sheet.Cells(row_no, NYUKA_PRODUCTS_NAME_COL_NO).Value = product_name

                        '入数
                        sheet.Cells(row_no, NYUKA_PRODUCTS_SLOT_COL_NO).Value = product_slot

                        '入庫
                        sheet.Cells(row_no, NYUKA_PRODUCTS_NYUKO_COL_NO).Value = product_storage

                        '単価
                        sheet.Cells(row_no, NYUKA_PRODUCTS_UNITPRICE_COL_NO).Value = product_unitprice

                        '為替
                        sheet.Cells(row_no, NYUKA_PRODUCTS_FX_COL_NO).Value = product_fx

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

        Return INGOODS_OK

    End Function


    '入荷金額情報を更新する
    Public Function ZaikoInfo() As Integer
        Dim i As Integer
        Dim delete_no As Integer = 0    '削除した回数
        Dim RowCount As Integer
        Dim dgv As New DataGridView

        Dim CInputSchedule As New InputGoodsArrivalSchedule

        Dim row_no As Integer = 2                                           '検索開始の行番号
        Dim day As Integer = 0                                              '日付の列番号
        Dim current_storage_num As Integer = 0                              '同じ日にすでに入荷した数

        '商品の情報
        Dim client_name As String = ""      '取引先
        Dim product_day As String = ""      '入荷予定日
        Dim product_code As String = ""     '品番
        Dim product_name As String = ""     '品名
        Dim product_slot As String = ""     '入数
        Dim product_storage As String = ""  '入庫数
        Dim product_unitprice As String = ""  '単価
        Dim product_fx As String = ""  '為替

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
        Dim book_zaikin As Excel.Workbook
        Dim sheet_zaikin As Excel.Worksheet

        dgv = MainFunction.NyukaKakutei_DataGridView
        RowCount = GetNotInputItems()

        If RowCount > 1 Then
            Try
                app = CreateObject("Excel.Application")
                app.Visible = False
                app.DisplayAlerts = False
            Catch ex As Exception
                MsgBox(ex.Message)
                Return INGOODS_NG
            End Try

            For i = 0 To RowCount - 2
                '入庫確定
                If dgv(0, i - delete_no).Value = True Then   '1行を削除した後に、行番号を1個増えるから
                    '入庫情報を設定する
                    client_name = dgv(DGV_CLIENT_NAME, i - delete_no).Value.ToString
                    product_day = dgv(DGV_DAY_COL_NO, i - delete_no).Value.ToString
                    product_code = dgv(DGV_CODE_COL_NO, i - delete_no).Value.ToString
                    product_name = dgv(DGV_NAME_COL_NO, i - delete_no).Value.ToString
                    product_slot = dgv(DGV_SLOT_COL_NO, i - delete_no).Value.ToString
                    product_storage = dgv(DGV_STORAGE_COL_NO, i - delete_no).Value.ToString
                    product_unitprice = dgv(DGV_UNITPRICE_COL_NO, i - delete_no).Value.ToString
                    product_fx = dgv(DGV_FX_COL_NO, i - delete_no).Value.ToString

                    '入庫情報を保存するファイルを探す
                    year_value = GetYearValue(product_day)
                    month_value = GetMonthValue(product_day)
                    save_filename = GetNyukaFileName(client_name, year_value, month_value)

                    '//在庫金額
                    If IO.File.Exists(DataPathDefinition.GetProductDataPath + "\" + year_value + "\" & year_value + month_value + "\在庫金額" + year_value & month_value + ".xlsx") Then 'Fileが存在する
                        book_zaikin = app.Workbooks.Open(DataPathDefinition.GetProductDataPath + "\" + year_value + "\" & year_value + month_value + "\在庫金額" + year_value & month_value + ".xlsx")
                    Else
                        MsgBox("File:" + DataPathDefinition.GetProductDataPath + "\" + year_value + "\" & year_value + month_value + "\在庫金額" + year_value & month_value + ".xlsx" + "が存在しない")
                        Return INGOODS_NG
                    End If

                    sheet_zaikin = book_zaikin.Worksheets(1)
                    Dim num As Integer = Integer.Parse(sheet_zaikin.Cells(1, 1).Value)
                    Dim ni, nj As Integer
                    For ni = 0 To num - 1
                        nj = 1 + ni * 3
                        If sheet_zaikin.Cells(nj + 2, 1).Value = product_code Then
                            Exit For
                        End If
                    Next
                    If ni = num Then
                        nj = 1 + ni * 3
                        sheet_zaikin.Cells(nj + 1, 1).Value = product_code
                        sheet_zaikin.Cells(nj + 1, 2).Value = product_name
                        sheet_zaikin.Cells(nj + 1, 3).Value = product_slot
                        sheet_zaikin.Cells(nj + 1, 4).Value = "0" '在庫金額
                    End If
                    Dim ionum As Integer = Integer.Parse(sheet_zaikin.Cells(nj + 2, 1).Value) + 1
                    sheet_zaikin.Cells(nj + 2, 1).Value = ionum.ToString
                    sheet_zaikin.Cells(nj + 2, ionum + 1).Value = "I," + product_storage + "," + product_unitprice + "," + product_fx
                    sheet_zaikin.Cells(nj + 3, 1).Value = "0"
                    'File Save
                    book_zaikin.Save()
                    'File Close
                    book_zaikin.Close()
                End If
            Next

            app.Quit()

            ' オブジェクトを解放します。
            sheet = Nothing
            book = Nothing
            app = Nothing

        End If

        Return INGOODS_OK

    End Function


    ''' <summary>
    ''' 在庫情報のExcelを更新する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function InputWareHouse() As Integer
        Dim i As Integer
        Dim delete_no As Integer = 0    '削除した回数
        Dim RowCount As Integer
        Dim dgv As New DataGridView

        Dim CInputSchedule As New InputGoodsArrivalSchedule
        Dim product_ino As InputGoodsArrivalSchedule.ProductInfo

        Dim row_no As Integer = 4                                           '検索開始の行番号
        Dim col_no_data As Integer = INVENTORY_PRODUCTS_DATE_START_COL_NO   '日付開始の列番号
        Dim day As Integer = 0                                              '日付の列番号
        Dim current_storage_num As Integer = 0                              '同じ日にすでに入荷した数

        '商品の情報
        Dim client_name As String = ""      '取引先
        Dim product_day As String = ""      '入荷予定日
        Dim product_code As String = ""     '品番
        Dim product_name As String = ""     '品名
        Dim product_slot As String = ""     '入数
        Dim product_storage As String = ""  '入庫数

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

        dgv = MainFunction.NyukaKakutei_DataGridView
        RowCount = GetNotInputItems() ' 未入荷確定の商品数を取得する

        If RowCount > 1 Then
            app = CreateObject("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False

            For i = 0 To RowCount - 2
                '入庫確定
                If dgv(0, i - delete_no).Value = True Then   '1行を削除した後に、行番号を1個増えるから
                    '入庫情報を収得する
                    client_name = dgv(DGV_CLIENT_NAME, i - delete_no).Value.ToString
                    product_day = dgv(DGV_DAY_COL_NO, i - delete_no).Value.ToString
                    product_code = dgv(DGV_CODE_COL_NO, i - delete_no).Value.ToString
                    product_name = dgv(DGV_NAME_COL_NO, i - delete_no).Value.ToString
                    product_slot = dgv(DGV_SLOT_COL_NO, i - delete_no).Value.ToString
                    product_storage = dgv(DGV_STORAGE_COL_NO, i - delete_no).Value.ToString

                    '入庫年月収得
                    year_value = GetYearValue(product_day)
                    month_value = GetMonthValue(product_day)
                    '今月の在庫ファイルのリンク収得
                    save_filename = GetFileNameSaveInputGoods(client_name, year_value, month_value)

                    '###############　この部分は要らないからとりあえずコメントアウト(Hoang)　　　########
                    'If System.IO.File.Exists(save_filename) Then    '在庫表のファイルが存在する
                    '    'Do Nothing
                    'Else    '在庫表のファイルが存在しない場合、新規ファイルを作成する
                    '    save_foldername = GetFolderNameSaveInputGoods(year_value, month_value)

                    '    'フォルダ作成
                    '    System.IO.Directory.CreateDirectory(save_foldername)

                    '    'Templateのファイルから新規の在庫表ファイルをコピーする
                    '    '本来なら、このルートには入らないはず
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
                        Return INGOODS_NG
                    End If

                    sheet = book.Worksheets(1)

                    '日付の列番号を決定
                    day_value = GetDayValue(product_day)
                    col_no_data = INVENTORY_PRODUCTS_DATE_START_COL_NO + day_value

                    '同じ品番の行番号を探す。同じ品番が無い場合、在庫ファイルに追加する
                    For row_no = INVENTORY_START_ROW_NO To MAX_ROW
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

                    '入荷確定のデータを在庫表に書き込む
                    '品番
                    'sheet.Cells(row_no, INVENTORY_PRODUCTS_CODE_COL_NO).Value = product_code

                    '品名
                    'sheet.Cells(row_no, INVENTORY_PRODUCTS_NAME_COL_NO).Value = product_name

                    '入数
                    'sheet.Cells(row_no, INVENTORY_PRODUCTS_SLOT_COL_NO).Value = product_slot

                    '入庫
                    'current_storage_num = sheet.Cells(row_no, col_no_data).Value
                    sheet.Cells(row_no, INVENTORY_PRODUCTS_STORAGE_COL_NO).Value = sheet.Cells(row_no, INVENTORY_PRODUCTS_STORAGE_COL_NO).Value + product_storage
                    'sheet.Cells(row_no, col_no_data).Value = sheet.Cells(row_no, col_no_data).Value + product_storage

                    'File Save
                    book.Save()

                    'File Close
                    book.Close()

                    '入荷予定ファイルの該当データを削除
                    product_ino.client_name = client_name
                    product_ino.product_day = product_day
                    product_ino.product_code = product_code
                    product_ino.product_name = product_name
                    product_ino.product_slot = product_slot
                    product_ino.product_storage = product_storage
                    CInputSchedule.DeleteScheduleData(product_ino)

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

        '未入荷商品数を更新する
        MainFunction.Num_NotInputs.Text = (GetNotInputItems() - 1).ToString

        If delete_no = 0 Then
            MsgBox("入荷確定商品を選択ください")
        Else
            MsgBox("入荷確定完了")
        End If

        Return INGOODS_OK

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

        dgv = MainFunction.NyukaKakutei_DataGridView
        RowCount = GetNotInputItems()

        If RowCount > 1 Then
            For i = 0 To RowCount - 2
                dgv.Rows.RemoveAt(i - delete_no)    '1行を削除した後に、行番号を1個増えるから
                delete_no = delete_no + 1
            Next
        End If

    End Sub

    ''' <summary>
    ''' Data Grid Viewのデータを設定する
    ''' 'thanh: 暫定に適当な数字を入れる
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetDataGridView()
        Dim dgv As New DataGridView
        Dim goods_arv_sch As New InputGoodsArrivalSchedule
        Dim strFile As String      '入荷予定ファイル 

        'Excelファイル関連
        Dim app As Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet
        Dim row_no As Integer

        '商品の情報
        Dim product_day As String = ""      '入荷予定日
        Dim product_client_name As String = "" '仕入れ先情報
        Dim product_name As String = ""     '品名
        Dim product_code As String = ""     '品番
        Dim product_slot As String = ""     '入数
        Dim product_storage As String = ""  '入庫数
        Dim product_unitprice As String = ""  '単価
        Dim product_fx As String = ""  '為替

        '初期化
        dgv = MainFunction.NyukaKakutei_DataGridView
        strFile = DataPathDefinition.GetNyukaYoteiFilePath()

        Try
            app = CreateObject("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
        Catch ex As Exception
            MsgBox(ex.Message)
            Exit Sub
        End Try

        'File Open
        If IO.File.Exists(strFile) Then 'Fileが存在する
            Try
                book = app.Workbooks.Open(strFile)
            Catch ex As Exception
                MsgBox(ex.Message)
                Exit Sub
            End Try
        Else
            MsgBox("File:" + strFile + "が存在しない")
            Exit Sub
        End If

        sheet = book.Worksheets(1)

        row_no = START_ROW

        Do While Len(sheet.Cells(row_no, READFILE_STORAGE_COL_NO).Value) > 0
            '日付
            product_day = sheet.Cells(row_no, READFILE_DAY_COL_NO).Value

            '仕入れ先情報
            product_client_name = sheet.Cells(row_no, READFILE_SHIIRE_COL_NO).Value


            '品番
            product_code = sheet.Cells(row_no, READFILE_CODE_COL_NO).Value

            '品名
            product_name = sheet.Cells(row_no, READFILE_NAME_COL_NO).Value

            '入数
            product_slot = sheet.Cells(row_no, READFILE_SLOT_COL_NO).Value

            '入庫数
            product_storage = sheet.Cells(row_no, READFILE_STORAGE_COL_NO).Value

            '単価
            product_unitprice = sheet.Cells(row_no, READFILE_UNITPRICE_COL_NO).Value

            '為替
            product_fx = sheet.Cells(row_no, READFILE_FX_COL_NO).Value

            dgv.Rows.Add(False, product_client_name, product_day, product_code, product_name, product_slot, product_storage, product_unitprice, product_fx)

            row_no = row_no + 1
        Loop

        'File Close
        book.Close()

        app.Quit()

        ' オブジェクトを解放します。
        sheet = Nothing
        book = Nothing
        app = Nothing

    End Sub

    ''' <summary>
    ''' 未入荷確定の商品数を取得する
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetNotInputItems() As Integer
        Dim items As Integer
        Dim dgv As New DataGridView

        dgv = MainFunction.NyukaKakutei_DataGridView
        items = dgv.Rows.Count

        Return (items)

    End Function

    ''' <summary>
    ''' 入庫情報を保存するファイル名を取得する
    ''' </summary>
    ''' <param name="client_name">取引先</param>
    ''' <param name="year">入庫年</param>
    ''' <param name="month">入庫月</param>
    ''' <returns>入庫情報を保存するファイル名</returns>
    ''' <remarks></remarks>
    Private Function GetFileNameSaveInputGoods(ByVal client_name As String, ByVal year As String, ByVal month As String) As String
        Dim file_path As String = ""    'ファイルの絶対パス
        Dim folder_path As String = ""  'フォルダパス

        folder_path = GetFolderNameSaveInputGoods(year, month)

        file_path = folder_path + "\" + year + month + "\" + "在庫" + year + month + ".xlsx"

        Return file_path


    End Function


    Private Function GetNyukaFileName(ByVal client_name As String, ByVal year As String, ByVal month As String) As String
        Dim file_path As String = ""    'ファイルの絶対パス
        Dim folder_path As String = ""  'フォルダパス

        folder_path = GetFolderNameSaveInputGoods(year, month)

        file_path = folder_path + "\" + year + month + "\" + "入荷" + year + month + ".xlsx"

        Return file_path


    End Function

    ''' <summary>
    ''' 在庫情報のフォルダパス
    ''' </summary>
    ''' <param name="year"></param>
    ''' <param name="month"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetFolderNameSaveInputGoods(ByVal year As String, ByVal month As String) As String
        Dim folder_path As String = ""

        folder_path = DataPathDefinition.GetProductDataPath()

        'yearの情報を加える
        folder_path = folder_path + "\" + year

        Return folder_path

    End Function


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

        Return date_split(2)

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

        '入荷日からYear, Monthの情報を取得
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
    ''' 在庫表のテンプレートファイルパス取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetSaveTemplateFileName() As String
        Return DataPathDefinition.GetTemplateZaikoPath()

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
        For row_no = INVENTORY_START_ROW_NO To MAX_ROW
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




#End Region


End Class
