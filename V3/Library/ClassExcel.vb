Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Public Class ClassExcel

#Region "Public Method"
    ''' <summary>
    ''' ExcelファイルOpen
    ''' </summary>
    ''' <param name="FilePath"></param>
    ''' <param name="SheetName"></param>
    ''' <param name="Visible">Excel表示するかどうかの設定</param>
    ''' <remarks></remarks>
    Public Sub OpenExcelFile(ByVal FilePath As String, ByVal SheetName As String, ByVal Visible As Boolean)
        ExcelOpen(FilePath, SheetName, Visible)
    End Sub

    Public Sub CloseExcelFile(ByVal FilePath As String)
        ExcelClose(FilePath, xlClose)
    End Sub

    Public Sub MRComObject(Of T As Class)(ByRef objCom As T, Optional ByVal force As Boolean = False)
        Dim IDEEnvironment As Boolean = False  'メッセージボックスを表示させたい場合は、True に設定
        If objCom Is Nothing Then
            If IDEEnvironment = True Then
                'テスト環境の場合は下記を実施し、後は、コメントにしておいて下さい。
                MessageBox.Show(Me, "Nothing です。")
            End If
            Return
        End If
        Try
            If System.Runtime.InteropServices.Marshal.IsComObject(objCom) Then
                Dim count As Integer
                If force Then
                    count = System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objCom)
                Else
                    count = System.Runtime.InteropServices.Marshal.ReleaseComObject(objCom)
                End If
                If IDEEnvironment = True AndAlso count <> 0 Then
                    Try
                        'テスト環境の場合は下記を実施し、後は、コメントにしておいて下さい。
                        MessageBox.Show(Me, TypeName(objCom) & " 要調査！ デクリメントされていません。")
                    Catch ex As Exception
                        MessageBox.Show(Me, " 要調査！ デクリメントされていません。")
                    End Try
                End If
            Else
                If IDEEnvironment = True Then
                    'テスト環境の場合は下記を実施し、後は、コメントにしておいて下さい。
                    MessageBox.Show(Me, "ComObject ではありませんので、解放処理の必要はありません。")
                End If
            End If
        Finally
            objCom = Nothing
        End Try
    End Sub

    Public Function InputLineData() As Integer
        Dim line_no As Integer
        Dim i As Integer

        For i = 2 To 1000
            xlRange = DirectCast(xlSheet.Cells(i, 1), Excel.Range)

            If xlRange.Value <> "" Then
                line_no = line_no + 1
            Else
                Exit For
            End If

        Next

        Return line_no

    End Function

    Public Function GetCellValue(ByVal Row As Integer, ByVal Column As Integer) As String

        xlRange = DirectCast(xlSheet.Cells(Row, Column), Excel.Range)

        If xlRange.Value <> vbNullString Then
            Return xlRange.Value.ToString
        Else
            Return 0
        End If

    End Function

    Public Sub SetCellValue(ByVal Row As Integer, ByVal Column As Integer, ByVal Value As String)
        xlSheet.Cells(Row, Column).Value = Value
    End Sub

    Public Sub DeleteRow(ByVal Row As Integer)
        xlSheet.Rows(Row).Delete()

    End Sub
#End Region

#Region "Private変数"

    '---------- Privateな変数の宣言 -----------------------------------
    Private xlApp As Excel.Application
    Private xlBooks As Excel.Workbooks
    Private xlBook As Excel.Workbook
    Private xlSheets As Excel.Sheets
    Private xlSheet As Excel.Worksheet
    Private xlRange As Excel.Range

    Private xlClose As Boolean    'ユーザが Excel を閉じようとしたかのフラグ

#End Region

#Region "Private Method"
    ''' <summary>
    ''' ExcelファイルOpenの処理
    ''' </summary>
    ''' <param name="FilePath"></param>
    ''' <param name="SheetName"></param>
    ''' <remarks></remarks>
    Private Sub ExcelOpen(ByVal FilePath As String, ByVal SheetName As String, ByVal Visible As Boolean)
        'Excel のオープン処理用プロシージャ
        xlClose = False   '起動中は、ユーザが Excel を閉じれないように

        xlApp = New Excel.Application
        xlApp.Visible = Visible

        'Excel の WorkbookBeforeClose イベントを取得
        '開くExcelファイルのXボタンをクリックするイベントとxlApp_WorkbookBeforeCloseのプロシジャを連動させる
        AddHandler xlApp.WorkbookBeforeClose, AddressOf xlApp_WorkbookBeforeClose

        xlBooks = xlApp.Workbooks
        If FilePath.Length = 0 Then
            '新規のファイルを開く場合
            xlBook = xlBooks.Add
            xlSheets = xlBook.Worksheets
            xlSheet = DirectCast(xlSheets.Item(1), Excel.Worksheet)
        Else
            '既存のファイルを開く場合
            xlBook = xlBooks.Open(FilePath)
            xlSheets = xlBook.Worksheets
            xlSheet = DirectCast(xlSheets(SheetName), Excel.Worksheet)
        End If
        xlApp.Visible = True
    End Sub

    Private Sub xlApp_WorkbookBeforeClose(ByVal Wb As Excel.Workbook, ByRef Cancel As Boolean)
        'VB2010 から Excel の WorkbookBeforeClose イベントを監視してユーザが Excel を閉じれないようにする
        If xlClose = False Then
            Cancel = True        'ユーザが Excel を閉じれないように
        Else
            Cancel = False    'プログラム上から指定の場合は、閉じる
        End If
    End Sub

    Public Sub ExcelClose(ByVal FilePath As String, Optional ByVal CancelSave As Boolean = True)
        'Excelファイルを上書き保存して終了処理用プロシージャ
        xlClose = True                'プログラムからExcel を閉じた時のフラグ
        xlApp.DisplayAlerts = False   '保存時の問合せのダイアログを非表示に設定
        If CancelSave Then
            Dim kts As String = System.IO.Path.GetExtension(FilePath).ToLower()
            Dim fm As Excel.XlFileFormat
            '拡張子に合せて保存形式を変更(使用する Excel のバージョンに注意)
            Select Case kts
                Case ".csv"    'CSV (カンマ区切り) 形式
                    fm = Excel.XlFileFormat.xlCSV
                Case ".xls"    'Excel 97～2003 ブック形式
                    fm = Excel.XlFileFormat.xlExcel8
                Case ".xlsx"   'Excel 2007～ブック形式
                    fm = Excel.XlFileFormat.xlOpenXMLWorkbook
                Case ".xlsm"   'Excel 2007～マクロ有効ブック形式
                    fm = Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled
                Case Else      '必要なものは、追加して下さい。
                    fm = Excel.XlFileFormat.xlWorkbookDefault
                    MessageBox.Show("ファイルの保存形式を確認して下さい。")
            End Select
            Try
                xlBook.SaveAs(Filename:=FilePath, FileFormat:=fm)    'ファイルに保存
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
        MRComObject(xlSheet)          'xlSheet の解放
        MRComObject(xlSheets)         'xlSheets の解放
        xlBook.Close()                'xlBook を閉じる
        MRComObject(xlBook)           'xlBook の解放
        MRComObject(xlBooks)          'xlBooks の解放
        xlApp.Quit()                  'Excelを閉じる
        MRComObject(xlApp)            'xlApp を解放
    End Sub

#End Region
End Class
