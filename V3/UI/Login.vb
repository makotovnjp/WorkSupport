Imports System.Runtime.InteropServices
Public Class Login

    Public Structure AuthorityStr
        Dim Stock As Boolean '入荷
        Dim Shipment As Boolean '出荷
        Dim Secondary_processing_domestic As Boolean '国内二次加工
        Dim Inventory_check As Boolean '在庫確認
        Dim Print_doccument As Boolean '資料発行
        Dim Set_user As Boolean 'ユーザーの権限設定
    End Structure

    Public Structure Member
        Dim Name As String
        Dim ID As String
        Dim Password As String
        '権限
        Dim Authority As AuthorityStr
        '履歴
        Dim Creation_date As String '作成日
        Dim Last_modification_date As String '最後に修正した日
    End Structure

    Public ID_list() As Member
    Public User_Number As Integer
    Public User_Index As Integer

    Dim password As String

    'ログイン
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Login.Click
        Dim i As Integer

        For i = 0 To User_Number - 1
            If tbx_ID.Text = ID_list(i).ID Then
                If tbx_Password.Text = ID_list(i).Password Then
                    User_Index = i
                    Exit For
                End If
            End If
        Next

        If i = User_Number Then
            tbx_Password.Text = ""
        Else
            Me.Hide()
            MainFunction.Show()
        End If
    End Sub




    'プログラム終了
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Exit.Click
        Me.Close()
    End Sub

    Private Sub Login_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim i As Integer
        Dim str_users As String
        'Dim File_path As String = System.IO.Directory.GetCurrentDirectory() + "\ID_pass.txt"
        ''ファイルを上書きし、Shift JISで書き込む 
        'Dim sw As New System.IO.StreamWriter(File_path, False, System.Text.Encoding.GetEncoding("shift_jis"))

        'sw.WriteLine(User_Number)
        ''1行ずつ書き込む 
        'For i = 0 To User_Number - 1
        '    sw.WriteLine("_NAME_ " & ID_list(i).Name & " _ID_ " & ID_list(i).ID & " _PASS_ " & ID_list(i).Password & "#")
        '    str = "_AUTHORITY_ "
        '    If ID_list(i).Authority.Stock = True Then
        '        str &= "Stock "
        '    End If
        '    If ID_list(i).Authority.Shipment = True Then
        '        str &= "Shipment "
        '    End If
        '    If ID_list(i).Authority.Secondary_processing_domestic = True Then
        '        str &= "Secondary_PD "
        '    End If
        '    If ID_list(i).Authority.Inventory_check = True Then
        '        str &= "Inventory_check "
        '    End If
        '    If ID_list(i).Authority.Print_doccument = True Then
        '        str &= "Print_doccument "
        '    End If
        '    If ID_list(i).Authority.Set_user = True Then
        '        str &= "Set_user"
        '    End If
        '    sw.WriteLine(str & "#")
        '    sw.WriteLine("_HISTORY_ " & ID_list(i).Creation_date & " _REVISE_ " & ID_list(i).Last_modification_date & "#")
        'Next
        'sw.Close()

        'ユーザーの数取得
        str_users = User_Number.ToString & vbCrLf
        For i = 0 To User_Number - 1
            str_users &= "_USER_" & ID_list(i).Name & "#" & ID_list(i).ID & "#" & ID_list(i).Password
            str_users &= "_ATR_"
            If ID_list(i).Authority.Stock = True Then
                str_users &= "1#"
            End If
            If ID_list(i).Authority.Shipment = True Then
                str_users &= "2#"
            End If
            If ID_list(i).Authority.Secondary_processing_domestic = True Then
                str_users &= "3#"
            End If
            If ID_list(i).Authority.Inventory_check = True Then
                str_users &= "4#"
            End If
            If ID_list(i).Authority.Print_doccument = True Then
                str_users &= "5#"
            End If
            If ID_list(i).Authority.Set_user = True Then
                str_users &= "6#"
            End If
            str_users &= "_TIME_" & ID_list(i).Creation_date & "#" & ID_list(i).Last_modification_date & "#" & vbCrLf
        Next
        My.Settings.ID_Pass = str_users
    End Sub


    'ユーザー情報をファイルから取得
    Private Sub Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim File_path As String = System.IO.Directory.GetCurrentDirectory() + "\ID_pass.txt"
        ''ファイルをShift(-JISコードとして開く)
        'Dim sr As New System.IO.StreamReader(File_path, System.Text.Encoding.GetEncoding("shift_jis"))
        'Dim Data_line As String
        'Dim i, u, v, t As Integer

        'User_Number = Integer.Parse(sr.ReadLine())
        'ReDim ID_list(User_Number)

        'For i = 0 To User_Number - 1
        '    '名前　ID　Password　取得
        '    Data_line = sr.ReadLine()
        '    u = Data_line.IndexOf(" _ID_")
        '    v = Data_line.IndexOf(" _PASS_")
        '    t = Data_line.IndexOf("#")
        '    ID_list(i).Name = Data_line.Substring(7, u - 7) '名前取得
        '    ID_list(i).ID = Data_line.Substring(u + 6, v - u - 6) 'ID取得
        '    ID_list(i).Password = Data_line.Substring(v + 8, t - v - 8) 'Password取得

        '    '権限情報　取得
        '    Data_line = sr.ReadLine()
        '    ID_list(i).Authority.Stock = Data_line.IndexOf("Stock") > 0
        '    ID_list(i).Authority.Shipment = Data_line.IndexOf("Shipment") > 0
        '    ID_list(i).Authority.Secondary_processing_domestic = Data_line.IndexOf("Secondary_PD") > 0
        '    ID_list(i).Authority.Inventory_check = Data_line.IndexOf("Inventory_check") > 0
        '    ID_list(i).Authority.Print_doccument = Data_line.IndexOf("Print_doccument") > 0
        '    ID_list(i).Authority.Set_user = Data_line.IndexOf("Set_user") > 0

        '    '履歴　取得
        '    Data_line = sr.ReadLine()
        '    u = Data_line.IndexOf(" _REVISE_")
        '    v = Data_line.IndexOf("#")
        '    ID_list(i).Creation_date = Data_line.Substring(10, u - 10)
        '    ID_list(i).Last_modification_date = Data_line.Substring(u + 10, v - u - 10)
        'Next
        Get_ID_Pass_Info()

    End Sub


    '********************************************
    'イベント：
    '処理：システムリーソースから情報取得
    '********************************************
    Private Sub Get_ID_Pass_Info()
        Dim str_sub As String
        Dim str_users As String = My.Settings.ID_Pass
        Dim t, i, r As Integer

        t = str_users.IndexOf(vbCrLf)
        User_Number = str_users.Substring(0, t)
        ReDim ID_list(User_Number)

        For i = 0 To User_Number - 1
            '名前取得
            t = str_users.IndexOf("_USER_", t) + 6
            r = str_users.IndexOf("#", t)
            ID_list(i).Name = str_users.Substring(t, r - t)
            'ID取得
            t = r + 1
            r = str_users.IndexOf("#", t)
            ID_list(i).ID = str_users.Substring(t, r - t)
            'Password取得
            t = r + 1
            r = str_users.IndexOf("_ATR_", t)
            ID_list(i).Password = str_users.Substring(t, r - t)

            '権限取得
            t = r + 5
            r = str_users.IndexOf("_TIME_", t)
            str_sub = str_users.Substring(t, r - t)
            ID_list(i).Authority.Stock = str_sub.IndexOf("1#") >= 0
            ID_list(i).Authority.Shipment = str_sub.IndexOf("2#") >= 0
            ID_list(i).Authority.Secondary_processing_domestic = str_sub.IndexOf("3#") >= 0
            ID_list(i).Authority.Inventory_check = str_sub.IndexOf("4#") >= 0
            ID_list(i).Authority.Print_doccument = str_sub.IndexOf("5#") >= 0
            ID_list(i).Authority.Set_user = str_sub.IndexOf("6#") >= 0

            '時間取得
            t = r + 6
            r = str_users.IndexOf("#", t)
            ID_list(i).Creation_date = str_users.Substring(t, r - t)
            t = r + 1
            r = str_users.IndexOf("#", t)
            ID_list(i).Last_modification_date = str_users.Substring(t, r - t)
        Next
    End Sub


    '********************************************
    'イベント：ID入力
    '処理：アルファベット、数字のみ入力
    '********************************************
    Private Sub tbx_ID_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbx_ID.KeyPress
        If e.KeyChar < "0"c OrElse "9"c < e.KeyChar Then
            '押されたキーが 0～9でない場合は
            If Not ((e.KeyChar >= "a"c And e.KeyChar <= "z") Or (e.KeyChar >= "A"c And e.KeyChar <= "Z")) Then
                If e.KeyChar <> ControlChars.Back Then
                    e.Handled = True
                End If
            End If
        End If
    End Sub

    '********************************************
    'イベント：Password入力
    '処理：アルファベット、数字のみ入力
    '      パースワード入力途中にエンターキーを押せばログイン
    '********************************************
    Private Sub tbx_Password_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles tbx_Password.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(13) Then
            Button_Login.PerformClick()
        Else
            If e.KeyChar < "0"c OrElse "9"c < e.KeyChar Then
                '押されたキーが 0～9でない場合は
                If Not ((e.KeyChar >= "a"c And e.KeyChar <= "z") Or (e.KeyChar >= "A"c And e.KeyChar <= "Z")) Then
                    If e.KeyChar <> ControlChars.Back Then
                        e.Handled = True
                    End If
                End If
            End If
        End If
    End Sub
End Class
