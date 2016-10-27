'File Openするクラス
'まずはOpenFileDialogを操作するクラスとして使う

Public Class FileOpener
    ''' <summary>
    ''' FileDialogで選択したFileのFile名
    ''' </summary>
    ''' <returns></returns>
    Public Function GetFileFromDialog() As String
        Dim filename As String = ""

        Dim OpenFileDialog As New OpenFileDialog()

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
            filename = OpenFileDialog.FileName
        End If

        Return filename
    End Function

End Class
