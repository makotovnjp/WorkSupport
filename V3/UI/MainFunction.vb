Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports System.Runtime.InteropServices

Public Class MainFunction

#Region "定数の定義"


#End Region

#Region "変数定義"

    'データ更新関連
    Private renew_data As New RenewDataBase()

    '入荷機能関連の変数
    Private arr_schd As New InputGoodsArrivalSchedule()
    Private input_goods As New InputGoods()
    Private input_sum As Integer

    '出荷機能関連の変数
    'Dim shipment_instance As New Shipment()
    'Private shipment_open_file_flag As Boolean
    Private output_schd As New OutputGoodsSchedule()
    Private output_goods As New OutputGoods()
    Private output_sum As Integer  '出荷予定の最大数


    '在庫確認関連
    Private checkstk As New C_CheckStock
#End Region

#Region "MainFormの処理"



    '********************************************
    'イベント：MainFunctionを閉じる
    '処理：
    '********************************************
    Private Sub MainFunction_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If MessageBox.Show("終了してもいいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            e.Cancel = True
        Else
            Login.tbx_ID.Text = ""
            Login.tbx_Password.Text = ""
            Login.Show()
        End If

    End Sub

    '********************************************
    'イベント：
    '処理：ユーザーが実行可能な機能を表示
    '********************************************
    Private Sub MainFunction_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim clientList As New List(Of String)

        '月や年が始まる時にデータ更新
        renew_data.ChangeDataForNewYearOrNewMonth()

        '入荷機能
        If Login.ID_list(Login.User_Index).Authority.Stock = False Then
            'TabControl1.TabPages.Remove(TabPage7)
            OutputYotei_FileInput_Bt.TabPages.Remove(InputGoods)
            OutputYotei_FileInput_Bt.TabPages.Remove(NyukaYotei)
        End If

        '出荷機能
        If Login.ID_list(Login.User_Index).Authority.Shipment = False Then
            OutputYotei_FileInput_Bt.TabPages.Remove(OutputGoodsYotei)
            OutputYotei_FileInput_Bt.TabPages.Remove(OutputGoods)
        End If

        '国内二次加工機能
        'If Login.ID_list(Login.User_Index).Authority.Secondary_processing_domestic = False Then
        OutputYotei_FileInput_Bt.TabPages.Remove(TabPage3)
        'End If
        '在庫確認機能
        If Login.ID_list(Login.User_Index).Authority.Inventory_check = False Then
            OutputYotei_FileInput_Bt.TabPages.Remove(TabPage4)
        End If
        '資料発行機能
        'If Login.ID_list(Login.User_Index).Authority.Print_doccument = False Then
        OutputYotei_FileInput_Bt.TabPages.Remove(TabPage5)
        'End If
        'ユーザー設定機能
        If Login.ID_list(Login.User_Index).Authority.Set_user = False Then
            OutputYotei_FileInput_Bt.TabPages.Remove(TabPage6)
        Else
            Get_Users_Information()
            If User_List.Items.Count >= 0 Then
                User_List.SelectedIndex = 0 '一人のユーザーの情報表示
            End If
        End If

        'Loadする時に、入荷予定のタブが表示されるため、入荷予定のClientListを表示する必要がある
        clientList = arr_schd.GetClienList()

        For Each client As String In clientList
            Me.NyukaYotei_ShiIreSaki_Combox.Items.Add(client)
        Next

    End Sub

#End Region
#Region "入荷予定"


    Private Sub OutputYotei_FileInput_Bt_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles OutputYotei_FileInput_Bt.SelectedIndexChanged
        If OutputYotei_FileInput_Bt.SelectedTab.Text = "在庫確認" Then
            checkstk.Stock_Display()
        ElseIf OutputYotei_FileInput_Bt.SelectedTab.Text = "入荷予定" Then
            NyukaYotei_ShiIreSaki_Combox.Text = ""
        End If
    End Sub
    ''' <summary>
    ''' イベント:入荷予定タブをクリックする時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TabNyukaYoteiSelecting(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles OutputYotei_FileInput_Bt.Selecting
        Dim clientList As New List(Of String)

        If OutputYotei_FileInput_Bt.SelectedTab.Name = "NyukaYotei" Then
            clientList = arr_schd.GetClienList()

            For Each client As String In clientList
                Me.NyukaYotei_ShiIreSaki_Combox.Items.Add(client)
            Next
        End If
    End Sub

    '*****************************************
    'GroupBox1
    '↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
    '*****************************************
    ' イベント：入荷予定入力をクリック
    Private Sub NyukaYotei_Input_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If NyukaYotei_ShiIreSaki_Combox.Text = "" Then
            MessageBox.Show("仕入れ先を入力ください", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            '入荷予定入力ボタンの色を変更して、編集できないようにする
            NyukaYotei_FileInput_Bt.Enabled = True

            '入荷予定入力のフォームを表示する
            NyukaYotei_OK_Bt.Enabled = False
        End If


    End Sub

    '*****************************************
    'GroupBox2
    '↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
    '*****************************************
    ''' <summary>
    ''' イベント：手動入力ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub NyukaYotei_ManualInput_Bt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NyukaYotei_ManualInput_Bt.Click
        If NyukaYotei_ShiIreSaki_Combox.Text = "" Then
            MessageBox.Show("仕入れ先を入力ください", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            '手動入力とファイル入力ボタンを非表示する
            NyukaYotei_ManualInput_Bt.Visible = False
            NyukaYotei_FileInput_Bt.Visible = False

            '入荷予定実行と選択やり直しボタンを表示する
            NyukaYotei_OK_Bt.Visible = True
            NyukaYoTei_Cancel.Visible = True

            '入力ファームを表示する
            NyukaYoTei_DataGridView1.Visible = True

            '為替入力表示
            Label14.Visible = True
            TextBox1.Visible = True

            '手動入力の処理
            arr_schd.ManualInputSchedule(input_sum)
        End If

    End Sub


    ''' <summary>
    ''' イベント：ファイル入力ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub NyukaYotei_FileInput_Bt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NyukaYotei_FileInput_Bt.Click
        If NyukaYotei_ShiIreSaki_Combox.Text = "" Then
            MessageBox.Show("仕入れ先を入力ください", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            If arr_schd.ReadFile() = InputGoodsArrivalSchedule.ARRSCHD_OK Then
                '手動入力とファイル入力ボタンを非表示する
                NyukaYotei_ManualInput_Bt.Visible = False
                NyukaYotei_FileInput_Bt.Visible = False

                '入荷予定保存ボタンと選択やり直しボタンを表示する
                NyukaYotei_OK_Bt.Visible = True
                NyukaYoTei_Cancel.Visible = True

                '為替入力表示
                Label14.Visible = True
                TextBox1.Visible = True

                'Data Grid Viewを表示する
                NyukaYoTei_DataGridView1.Visible = True

            End If
        End If
    End Sub

    ''' <summary>
    ''' イベント：入荷予定OKボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub NyukaYotei_OK_Bt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NyukaYotei_OK_Bt.Click
        If TextBox1.Text = "" Then
            MessageBox.Show("為替比を入力ください", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim result As DialogResult = MessageBox.Show("入荷予定を入れますか？", "確認", MessageBoxButtons.YesNo)

        If result = Windows.Forms.DialogResult.Yes Then
            If arr_schd.WriteData(input_sum) = InputGoodsArrivalSchedule.ARRSCHD_OK Then

                MsgBox("入荷予定商品の処理は完了しました")
                NyukaYotei_Form_Ini()
            Else

            End If

        End If

    End Sub

    ''' <summary>
    ''' 入荷予定のファームの初期化処理
    ''' </summary>
    ''' <remarks></remarks>
    ''' 
    Private Sub NyukaYotei_Form_Ini()
        '手動入力ボタン初期化
        NyukaYotei_ManualInput_Bt.BackColor = Color.Transparent
        NyukaYotei_ManualInput_Bt.Visible = True
        '自動入力ボタン初期化
        NyukaYotei_FileInput_Bt.BackColor = Color.Transparent
        NyukaYotei_FileInput_Bt.Visible = True
        '入荷予定ボタン初期化
        NyukaYotei_OK_Bt.BackColor = Color.Transparent
        NyukaYotei_OK_Bt.Visible = False
        'やり直しボタン初期化
        NyukaYoTei_Cancel.BackColor = Color.Transparent
        NyukaYoTei_Cancel.Visible = False

        'DataGridView
        NyukaYoTei_DataGridView1.Visible = False

        '為替入力非表示
        Label14.Visible = False
        TextBox1.Visible = False

        'thanh:todo Data GridViewのデータを削除する
        '仕入れ先の表示を初期化する
        NyukaYotei_ShiIreSaki_Combox.Text = ""

    End Sub

    ''' <summary>
    ''' イベント：Cancelボタンをクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub NyukaYoTei_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NyukaYoTei_Cancel.Click
        Dim result As DialogResult = MessageBox.Show("キャンセルしますか？", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        If result = Windows.Forms.DialogResult.Yes Then
            NyukaYotei_Form_Ini()
            arr_schd.Cancel()
        End If

    End Sub

#End Region

#Region "入荷確定機能"

    ''' <summary>
    ''' イベント:入荷確定タブをクリックする時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TabControl1_Selecting(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles OutputYotei_FileInput_Bt.Selecting
        If OutputYotei_FileInput_Bt.SelectedTab.Name = "InputGoods" Then
            NyukaKakutei_OK.Visible = True
            NyukaKakutei_CheckBoxAll.Visible = True
            NyukaKakutei_DataGridView.Visible = True

            '入荷確定の初期化処理
            input_goods.Init()

        End If
    End Sub

    ''' <summary>
    ''' イベント：入荷確定ボタンをクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub NyukaKakutei_OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NyukaKakutei_OK.Click
        Dim result As DialogResult = MessageBox.Show("入荷を確定して良いですか？", "確認", MessageBoxButtons.YesNo)

        If result = Windows.Forms.DialogResult.Yes Then
            '今月の入荷ファイルに追加
            input_goods.InputNyukaFile()
            '在庫情報のエクセルファイル更新
            input_goods.InputWareHouse()

            '在庫金額情報のエクセルファイル更新
            input_goods.ZaikoInfo()
            NyukaKakutei_CheckBoxAll.Text = "全てを選択"
        End If

    End Sub

    ''' <summary>
    ''' 全てをチェックのボタンをクリックイベントの処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub NyukaKakutei_CheckBoxAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NyukaKakutei_CheckBoxAll.Click
        If NyukaKakutei_CheckBoxAll.Text = "全てを選択" Then
            NyukaKakutei_CheckBoxAll.Text = "全てを外す"
            input_goods.SetCheckBox(True)

        Else
            NyukaKakutei_CheckBoxAll.Text = "全てを選択"
            input_goods.SetCheckBox(False)

        End If

    End Sub


#End Region

#Region "出荷予定機能"
    ''' <summary>
    ''' イベント:出荷予定タブをクリックする時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TabOutputYoteiSelecting(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles OutputYotei_FileInput_Bt.Selecting
        If OutputYotei_FileInput_Bt.SelectedTab.Name = "OutputGoodsYotei" Then
            output_schd.ShowCustomersList()

        End If
    End Sub

    ' イベント：入荷予定入力をクリック
    Private Sub OutputYotei_Input_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OutputYotei_InputInfo_Button.Click
        If OutputYotei_CustomerList_Combox.Text = "" Then
            MessageBox.Show("届け先を入力ください", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            '出荷予定入力ボタンの色を変更して、編集できないようにする
            OutputYotei_InputInfo_Button.Enabled = False
            OutputYotei_FileInput_Bt.Enabled = True

            '入荷予定入力のフォームを表示する
            OutputYotei_GroupBox2.Visible = True
            OutputYotei_OK_Button.Enabled = False
        End If

    End Sub


    ''' <summary>
    ''' イベント：出荷予定機能の手動入力のボタンをクリックする
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub OutputYotei_ManualInput_Bt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OutputYotei_ManualInput_Bt.Click

        '手動入力ボタンの色を変更する
        OutputYotei_ManualInput_Bt.BackColor = Color.Yellow 'ボタンの背景を黄色にする

        'ファイル入力ボタン, 手動入力ボタンをクリックできないようにする
        OutputYotei_FileInput_Bt2.Visible = False
        OutputYotei_ManualInput_Bt.Enabled = False

        '出荷予定を入れるボタンをEnableする
        OutputYotei_OK_Button.Enabled = True

        '入力ファームを表示する
        OutputYoTei_DataGridView1.Visible = True

        '手動入力の処理
        output_schd.ManualInputSchedule(output_sum)

    End Sub

    ''' <summary>
    ''' イベント：出荷予定機能のファイル入力ボタンクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub OutputYotei_FileInput_Bt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OutputYotei_FileInput_Bt2.Click

        If output_schd.ReadFile() = OutputGoodsSchedule.OUTSCHD_OK Then

            'ファイル入力ボタンの色を変更する
            OutputYotei_FileInput_Bt2.BackColor = Color.Yellow 'ボタンの背景を黄色にする

            'ファイル入力ボタン, 手動入力ボタンをクリックできないようにする
            OutputYotei_FileInput_Bt2.Enabled = False
            OutputYotei_ManualInput_Bt.Visible = False

            '出荷予定を入れるボタンをEnableする
            OutputYotei_OK_Button.Enabled = True

            'Data Grid Viewを表示する
            OutputYoTei_DataGridView1.Visible = True

        End If

    End Sub

    Private Sub OutputYotei_OK_Bt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OutputYotei_OK_Button.Click
        Dim result As DialogResult = MessageBox.Show("出荷予定を入れますか？", "確認", MessageBoxButtons.YesNo)

        If result = Windows.Forms.DialogResult.Yes Then
            output_schd.WriteData(output_sum)

            MsgBox("出荷予定商品の処理は完了しました")

            OutputYotei_Form_Ini()
        End If

    End Sub


    ''' <summary>
    ''' 出荷予定のファームを初期化する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub OutputYotei_Form_Ini()
        'Groupbox1
        OutputYotei_InputInfo_Button.BackColor = Color.Transparent
        OutputYotei_InputInfo_Button.Enabled = True

        'Groupbox2
        OutputYotei_GroupBox2.Visible = False


        OutputYotei_ManualInput_Bt.BackColor = Color.Transparent
        OutputYotei_ManualInput_Bt.Visible = True
        OutputYotei_ManualInput_Bt.Enabled = True

        OutputYotei_FileInput_Bt2.BackColor = Color.Transparent
        OutputYotei_FileInput_Bt2.Visible = True
        OutputYotei_FileInput_Bt2.Enabled = False

        OutputYotei_OK_Button.BackColor = Color.Transparent
        OutputYotei_OK_Button.Enabled = False

        OutputYotei_Cancel.BackColor = Color.Transparent
        OutputYotei_Cancel.Enabled = True

        'DataGridView
        OutputYoTei_DataGridView1.Visible = False

        '届先の表示を初期化する
        OutputYotei_CustomerList_Combox.Text = ""

    End Sub

    ''' <summary>
    ''' イベント：出荷予定機能のCancelボタンをクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub OutputYoTei_Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OutputYotei_Cancel.Click
        Dim result As DialogResult = MessageBox.Show("キャンセルしますか？", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        If result = Windows.Forms.DialogResult.Yes Then
            OutputYotei_Form_Ini()

            output_schd.Cancel()

        End If

    End Sub


#End Region

#Region "出荷確定機能"
    ''' <summary>
    ''' イベント:出荷確定タブをクリックする時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub OutputKakutei_TabControl_Selecting(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles OutputYotei_FileInput_Bt.Selecting
        If OutputYotei_FileInput_Bt.SelectedTab.Name = "OutputGoods" Then
            OutputKakutei_OK.Visible = True
            OutputKakutei_CheckBoxAll.Visible = True
            OutputKakutei_DataGridView.Visible = True

            '入荷確定の初期化処理
            output_goods.Init()

        End If
    End Sub

    ''' <summary>
    ''' 出荷確定機能の全てをチェックのボタンをクリックイベントの処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub OutputKakutei_CheckBoxAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OutputKakutei_CheckBoxAll.Click
        If OutputKakutei_CheckBoxAll.Text = "全てを選択" Then
            OutputKakutei_CheckBoxAll.Text = "全てを外す"
            output_goods.SetCheckBox(True)

        Else
            OutputKakutei_CheckBoxAll.Text = "全てを選択"
            output_goods.SetCheckBox(False)

        End If

    End Sub

    ''' <summary>
    ''' イベント：出荷確定ボタンをクリック
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub OutputKakutei_OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OutputKakutei_OK.Click
        Dim result As DialogResult = MessageBox.Show("出荷を確定して良いですか？", "確認", MessageBoxButtons.YesNo)

        If result = Windows.Forms.DialogResult.Yes Then
            '今月の出荷ファイルに追加
            output_goods.WriteOutputGoodsFile()

            '在庫情報のエクセルファイル更新
            output_goods.InputWareHouse()
            OutputKakutei_CheckBoxAll.Text = "全てを選択"
        End If

    End Sub

#End Region

#Region "Event処理"
    '********************************************
    'イベント：
    '処理：名簿表示更新
    '********************************************
    Private Sub Get_Users_Information()
        Dim i As Integer

        User_List.Items.Clear()
        ID_TextBox.Text = ""
        Password_TextBox.Text = ""
        For i = 0 To Login.User_Number - 1
            User_List.Items.Add(Login.ID_list(i).Name)
        Next
    End Sub



    '********************************************
    'イベント：「名簿」から選択
    '処理：選択したユーザーの権限を表示する
    '********************************************
    Private Sub User_List_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles User_List.SelectedIndexChanged
        Dim sltindex As Integer = User_List.SelectedIndex

        ID_TextBox.Text = Login.ID_list(User_List.SelectedIndex).ID
        Password_TextBox.Text = Login.ID_list(User_List.SelectedIndex).Password
        View_Authority_Check(sltindex)
    End Sub


    '********************************************
    'イベント：
    '処理：選択したユーザーの権限を表示する
    '********************************************
    Private Sub View_Authority_Check(ByVal sltindex As Integer)
        CheckBox1.Checked = Login.ID_list(sltindex).Authority.Stock
        CheckBox2.Checked = Login.ID_list(sltindex).Authority.Shipment
        CheckBox3.Checked = Login.ID_list(sltindex).Authority.Secondary_processing_domestic
        CheckBox4.Checked = Login.ID_list(sltindex).Authority.Inventory_check
        CheckBox5.Checked = Login.ID_list(sltindex).Authority.Print_doccument
        CheckBox6.Checked = Login.ID_list(sltindex).Authority.Set_user
        Label3.Text = Login.ID_list(sltindex).Creation_date
        Label4.Text = Login.ID_list(sltindex).Last_modification_date
    End Sub

    '********************************************
    'イベント：
    '処理：「新規」や「修正」の枠内の表示初期化
    '********************************************
    Private Sub Initialzing_Input_Frame()
        CheckBox7.Checked = False
        CheckBox8.Checked = False
        CheckBox9.Checked = False
        CheckBox10.Checked = False
        CheckBox11.Checked = False
        CheckBox12.Checked = False
    End Sub


    '********************************************
    'イベント：「新規」ボタンクリック
    '処理：新しいユーザー追加
    '********************************************
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        If GroupBox5.Visible = False Then
            TextBox9.Text = ""
            TextBox8.Text = ""
            TextBox7.Text = ""
            Button8.BackColor = Color.Yellow '「新規」ボタンの背景を黄色にする
            'Button10.Enabled = False '「修正」ボタンの使用禁止
            'Button_Delete.Enabled = False '「削除」ボタンの使用禁止
            Button10.Visible = False
            Button_Delete.Visible = False
            GroupBox5.Visible = True
            'Button12.Text = "OK"
            Initialzing_Input_Frame()
        End If
    End Sub


    '********************************************
    'イベント：「修正」ボタンクリック
    '処理：ユーザー情報修正
    '********************************************
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        If User_List.SelectedIndex <0 Then
            MessageBox.Show("ユーザーを選択してください。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            If GroupBox5.Visible = False Then
                TextBox9.Text = User_List.Text
                TextBox8.Text = ID_TextBox.Text
                TextBox7.Text = Password_TextBox.Text

                Button10.BackColor = Color.Yellow '「修正」ボタンの背景を黄色にする
                'Button8.Enabled = False '「新規」ボタンの使用禁止
                'Button_Delete.Enabled = False '「削除」ボタンの使用禁止
                Button8.Visible = False
                Button_Delete.Visible = False
                GroupBox5.Visible = True
                'Button12.Text = "OK"
                Initialzing_Input_Frame()
            End If
        End If

    End Sub


    '********************************************
    'イベント：新規又は修正処理の「キャンセル」ボタンクリック
    '処理：入力フレームを非表示
    '********************************************
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Button8.BackColor = Color.Transparent
        Button10.BackColor = Color.Transparent
        Button_Delete.BackColor = Color.Transparent
        'Button8.Enabled = True
        'Button10.Enabled = True
        'Button_Delete.Enabled = True
        Button8.Visible = True
        Button10.Visible = True
        Button_Delete.Visible = True
        GroupBox5.Text = ""
        GroupBox5.Visible = False

    End Sub


    '********************************************
    'イベント：入力枠内の「新規」や「修正」ボタンクリック
    '処理：新規処理又は修正処理を行う
    '********************************************
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim i As Integer

        '全て入力したかの確認
        If (TextBox7.Text = "") Or (TextBox8.Text = "") Or (TextBox9.Text = "") Then
            MessageBox.Show("ユーザーの名前、ID、パースワールドを全て入力してください。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            If Button12.Text = "新規" Then
                
                '新規するユーザーのIDが現存のユーザーのIDと同じか確認する
                For i = 0 To Login.User_Number - 1
                    If Login.ID_list(i).ID = TextBox8.Text Then
                        Exit For
                    End If
                Next

                If i = Login.User_Number Then
                    '現存するユーザーの同じIDがない場合
                    Add_New_User()
                    MessageBox.Show("新しいユーザーを追加できた", "お知らせ", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Get_Users_Information() '名簿更新
                    Button13.PerformClick()
                    User_List.Text = User_List.Items.Item(User_List.Items.Count - 1).ToString
                Else
                    MessageBox.Show("このIDを使用しています。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If

            ElseIf Button12.Text = "修正" Then
                '修正処理
                
                Dim sltIndex As Integer = User_List.SelectedIndex
                Renew_User_Information(sltIndex)
                User_List.Text = User_List.Items.Item(sltIndex).ToString
                Button13.PerformClick()
            End If
        End If
    End Sub


    '********************************************
    'イベント：
    '処理：選択ユーザーの情報をID_listで修正する
    '********************************************
    Private Sub Renew_User_Information(ByVal sltIndex As Integer)
        Dim NewInput As New Login.Member

        'ID_list上で修正
        Get_New_Input(NewInput)
        Login.ID_list(sltIndex) = NewInput
        '表示更新
        Get_Users_Information()
        View_Authority_Check(sltIndex)
    End Sub


    '********************************************
    'イベント：
    '処理：'新規したユーザーをID_listに追加する
    '********************************************
    Private Sub Add_New_User()
        Dim NewInput As New Login.Member

        ReDim Preserve Login.ID_list(Login.User_Number)
        Get_New_Input(NewInput) '入力した情報取得
        Login.ID_list(Login.User_Number) = NewInput
        Login.User_Number += 1
    End Sub

    '********************************************
    'イベント：
    '処理：「新規」や「修正」に関する入力した情報取得し、NewInputに入れる
    '********************************************
    Private Sub Get_New_Input(ByRef NewInput As Login.Member)
        Dim dtToday As DateTime = DateTime.Today

        '名前　ID　Password　取得
        NewInput.Name = TextBox9.Text '名前取得
        NewInput.ID = TextBox8.Text 'ID取得
        NewInput.Password = TextBox7.Text 'Password取得

        '権限情報　取得
        NewInput.Authority.Stock = CheckBox8.Checked
        NewInput.Authority.Shipment = CheckBox10.Checked
        NewInput.Authority.Secondary_processing_domestic = CheckBox7.Checked
        NewInput.Authority.Inventory_check = CheckBox12.Checked
        NewInput.Authority.Print_doccument = CheckBox11.Checked
        NewInput.Authority.Set_user = CheckBox9.Checked

        '履歴　取得
        NewInput.Creation_date = dtToday.ToString()
        NewInput.Last_modification_date = dtToday.ToString()
    End Sub

    '********************************************
    'イベント：「削除」ボタンクリック
    '処理：削除したいユーザーをID_Listから消す
    '********************************************
    Private Sub Button_Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Delete.Click
        Dim i, j As Integer
        Dim result As DialogResult = MessageBox.Show("このユーザーを削除ですか。", "確認", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning)

        If result = DialogResult.Yes Then
            '名簿において削除したいユーザーのインデックス
            j = User_List.SelectedIndex
            'データシフト
            For i = j To Login.User_Number - 2
                Login.ID_list(i) = Login.ID_list(i + 1)
            Next
            Login.User_Number -= 1
            '名簿の再更新
            Get_Users_Information()
            User_List.Text = "名前　/ ID　/ Password"
            MessageBox.Show("ユーザーを削除しました。", "お知らせ", MessageBoxButtons.OK, MessageBoxIcon.Information)
            If User_List.Items.Count >= 0 Then
                User_List.SelectedIndex = 0 '一人のユーザーの情報表示
            End If
        End If
    End Sub

    '********************************************
    'イベント：新しいID又はPassword入力
    '処理：アルファベット、数字のみ入力
    '********************************************
    Private Sub TextBox7_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox7.KeyPress, TextBox8.KeyPress
        If e.KeyChar < "0"c OrElse "9"c < e.KeyChar Then
            '押されたキーが 0～9でない場合は
            If Not ((e.KeyChar >= "a"c And e.KeyChar <= "z") Or (e.KeyChar >= "A"c And e.KeyChar <= "Z")) Then
                If e.KeyChar <> ControlChars.Back Then
                    e.Handled = True
                End If
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Check_stock.Show()
    End Sub


#End Region

   
End Class