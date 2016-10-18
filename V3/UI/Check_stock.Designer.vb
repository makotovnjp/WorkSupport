<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Check_stock
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle16 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle15 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.NyukaKakutei_DataGridView = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn7 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn13 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn15 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn16 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.単価 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.MonthComboBox5 = New System.Windows.Forms.ComboBox()
        Me.YearComboBox6 = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.MonthComboBox4 = New System.Windows.Forms.ComboBox()
        Me.YearComboBox3 = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.NameOfProduct_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Partnumber_CheckBox = New System.Windows.Forms.CheckBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.ComboBox3 = New System.Windows.Forms.ComboBox()
        Me.ComboBox4 = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        CType(Me.NyukaKakutei_DataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'NyukaKakutei_DataGridView
        '
        Me.NyukaKakutei_DataGridView.AllowUserToAddRows = False
        Me.NyukaKakutei_DataGridView.AllowUserToDeleteRows = False
        Me.NyukaKakutei_DataGridView.AllowUserToResizeColumns = False
        Me.NyukaKakutei_DataGridView.AllowUserToResizeRows = False
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.NyukaKakutei_DataGridView.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle9
        Me.NyukaKakutei_DataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.NyukaKakutei_DataGridView.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle10.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle10.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        DataGridViewCellStyle10.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle10.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.NyukaKakutei_DataGridView.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle10
        Me.NyukaKakutei_DataGridView.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn7, Me.DataGridViewTextBoxColumn13, Me.DataGridViewTextBoxColumn15, Me.DataGridViewTextBoxColumn16, Me.単価})
        Me.NyukaKakutei_DataGridView.Location = New System.Drawing.Point(18, 94)
        Me.NyukaKakutei_DataGridView.Name = "NyukaKakutei_DataGridView"
        Me.NyukaKakutei_DataGridView.ReadOnly = True
        DataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle16.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle16.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        DataGridViewCellStyle16.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle16.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle16.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle16.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.NyukaKakutei_DataGridView.RowHeadersDefaultCellStyle = DataGridViewCellStyle16
        Me.NyukaKakutei_DataGridView.RowTemplate.Height = 21
        Me.NyukaKakutei_DataGridView.Size = New System.Drawing.Size(770, 316)
        Me.NyukaKakutei_DataGridView.TabIndex = 13
        Me.NyukaKakutei_DataGridView.Visible = False
        '
        'DataGridViewTextBoxColumn7
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.DataGridViewTextBoxColumn7.DefaultCellStyle = DataGridViewCellStyle11
        Me.DataGridViewTextBoxColumn7.FillWeight = 42.08072!
        Me.DataGridViewTextBoxColumn7.HeaderText = "入・出"
        Me.DataGridViewTextBoxColumn7.Name = "DataGridViewTextBoxColumn7"
        Me.DataGridViewTextBoxColumn7.ReadOnly = True
        '
        'DataGridViewTextBoxColumn13
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.DataGridViewTextBoxColumn13.DefaultCellStyle = DataGridViewCellStyle12
        Me.DataGridViewTextBoxColumn13.FillWeight = 88.63597!
        Me.DataGridViewTextBoxColumn13.HeaderText = "日付"
        Me.DataGridViewTextBoxColumn13.Name = "DataGridViewTextBoxColumn13"
        Me.DataGridViewTextBoxColumn13.ReadOnly = True
        '
        'DataGridViewTextBoxColumn15
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.DataGridViewTextBoxColumn15.DefaultCellStyle = DataGridViewCellStyle13
        Me.DataGridViewTextBoxColumn15.FillWeight = 157.3604!
        Me.DataGridViewTextBoxColumn15.HeaderText = "取引名"
        Me.DataGridViewTextBoxColumn15.Name = "DataGridViewTextBoxColumn15"
        Me.DataGridViewTextBoxColumn15.ReadOnly = True
        '
        'DataGridViewTextBoxColumn16
        '
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.DataGridViewTextBoxColumn16.DefaultCellStyle = DataGridViewCellStyle14
        Me.DataGridViewTextBoxColumn16.FillWeight = 105.508!
        Me.DataGridViewTextBoxColumn16.HeaderText = "入・出庫数"
        Me.DataGridViewTextBoxColumn16.Name = "DataGridViewTextBoxColumn16"
        Me.DataGridViewTextBoxColumn16.ReadOnly = True
        '
        '単価
        '
        DataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.単価.DefaultCellStyle = DataGridViewCellStyle15
        Me.単価.FillWeight = 106.415!
        Me.単価.HeaderText = "メモ"
        Me.単価.Name = "単価"
        Me.単価.ReadOnly = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label1.Location = New System.Drawing.Point(41, 70)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 13)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "品番："
        Me.Label1.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label2.Location = New System.Drawing.Point(202, 70)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 13)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "品名："
        Me.Label2.Visible = False
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("MS UI Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Red
        Me.Label3.Location = New System.Drawing.Point(79, 68)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 15)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "###"
        Me.Label3.Visible = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("MS UI Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Red
        Me.Label4.Location = New System.Drawing.Point(244, 68)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(34, 15)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "###"
        Me.Label4.Visible = False
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("MS UI Gothic", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Red
        Me.Button1.Location = New System.Drawing.Point(632, 13)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(109, 42)
        Me.Button1.TabIndex = 18
        Me.Button1.Text = "確認"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.MonthComboBox5)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.YearComboBox6)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.MonthComboBox4)
        Me.GroupBox1.Controls.Add(Me.YearComboBox3)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.ComboBox2)
        Me.GroupBox1.Controls.Add(Me.NameOfProduct_CheckBox)
        Me.GroupBox1.Controls.Add(Me.Partnumber_CheckBox)
        Me.GroupBox1.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(18, 1)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(765, 61)
        Me.GroupBox1.TabIndex = 19
        Me.GroupBox1.TabStop = False
        '
        'MonthComboBox5
        '
        Me.MonthComboBox5.FormattingEnabled = True
        Me.MonthComboBox5.Items.AddRange(New Object() {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"})
        Me.MonthComboBox5.Location = New System.Drawing.Point(537, 32)
        Me.MonthComboBox5.Name = "MonthComboBox5"
        Me.MonthComboBox5.Size = New System.Drawing.Size(38, 21)
        Me.MonthComboBox5.TabIndex = 11
        Me.MonthComboBox5.Text = "12"
        '
        'YearComboBox6
        '
        Me.YearComboBox6.FormattingEnabled = True
        Me.YearComboBox6.Location = New System.Drawing.Point(480, 32)
        Me.YearComboBox6.Name = "YearComboBox6"
        Me.YearComboBox6.Size = New System.Drawing.Size(51, 21)
        Me.YearComboBox6.TabIndex = 10
        Me.YearComboBox6.Text = "2015"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(456, 35)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(20, 13)
        Me.Label8.TabIndex = 9
        Me.Label8.Text = "～"
        '
        'MonthComboBox4
        '
        Me.MonthComboBox4.FormattingEnabled = True
        Me.MonthComboBox4.Items.AddRange(New Object() {"01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"})
        Me.MonthComboBox4.Location = New System.Drawing.Point(412, 32)
        Me.MonthComboBox4.Name = "MonthComboBox4"
        Me.MonthComboBox4.Size = New System.Drawing.Size(38, 21)
        Me.MonthComboBox4.TabIndex = 8
        Me.MonthComboBox4.Text = "12"
        '
        'YearComboBox3
        '
        Me.YearComboBox3.FormattingEnabled = True
        Me.YearComboBox3.Location = New System.Drawing.Point(355, 32)
        Me.YearComboBox3.Name = "YearComboBox3"
        Me.YearComboBox3.Size = New System.Drawing.Size(51, 21)
        Me.YearComboBox3.TabIndex = 7
        Me.YearComboBox3.Text = "2015"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(353, 14)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(59, 13)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "確認期間"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(155, 14)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(125, 13)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "品番を選択してください"
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Items.AddRange(New Object() {"なし"})
        Me.ComboBox2.Location = New System.Drawing.Point(157, 31)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(148, 21)
        Me.ComboBox2.TabIndex = 4
        '
        'NameOfProduct_CheckBox
        '
        Me.NameOfProduct_CheckBox.AutoSize = True
        Me.NameOfProduct_CheckBox.Location = New System.Drawing.Point(8, 35)
        Me.NameOfProduct_CheckBox.Name = "NameOfProduct_CheckBox"
        Me.NameOfProduct_CheckBox.Size = New System.Drawing.Size(89, 17)
        Me.NameOfProduct_CheckBox.TabIndex = 1
        Me.NameOfProduct_CheckBox.Text = "品名で確認"
        Me.NameOfProduct_CheckBox.UseVisualStyleBackColor = True
        '
        'Partnumber_CheckBox
        '
        Me.Partnumber_CheckBox.AutoSize = True
        Me.Partnumber_CheckBox.Checked = True
        Me.Partnumber_CheckBox.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Partnumber_CheckBox.Location = New System.Drawing.Point(8, 13)
        Me.Partnumber_CheckBox.Name = "Partnumber_CheckBox"
        Me.Partnumber_CheckBox.Size = New System.Drawing.Size(89, 17)
        Me.Partnumber_CheckBox.TabIndex = 0
        Me.Partnumber_CheckBox.Text = "品番で確認"
        Me.Partnumber_CheckBox.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("MS UI Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.Red
        Me.Label5.Location = New System.Drawing.Point(399, 68)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(34, 15)
        Me.Label5.TabIndex = 21
        Me.Label5.Text = "###"
        Me.Label5.Visible = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label9.Location = New System.Drawing.Point(361, 70)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(40, 13)
        Me.Label9.TabIndex = 20
        Me.Label9.Text = "入数："
        Me.Label9.Visible = False
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"なし"})
        Me.ComboBox1.Location = New System.Drawing.Point(681, 68)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(30, 20)
        Me.ComboBox1.TabIndex = 22
        Me.ComboBox1.Visible = False
        '
        'ComboBox3
        '
        Me.ComboBox3.FormattingEnabled = True
        Me.ComboBox3.Items.AddRange(New Object() {"なし"})
        Me.ComboBox3.Location = New System.Drawing.Point(717, 68)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(30, 20)
        Me.ComboBox3.TabIndex = 23
        Me.ComboBox3.Visible = False
        '
        'ComboBox4
        '
        Me.ComboBox4.FormattingEnabled = True
        Me.ComboBox4.Items.AddRange(New Object() {"なし"})
        Me.ComboBox4.Location = New System.Drawing.Point(753, 68)
        Me.ComboBox4.Name = "ComboBox4"
        Me.ComboBox4.Size = New System.Drawing.Size(30, 20)
        Me.ComboBox4.TabIndex = 24
        Me.ComboBox4.Visible = False
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("MS UI Gothic", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Red
        Me.Label10.Location = New System.Drawing.Point(541, 68)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(34, 15)
        Me.Label10.TabIndex = 26
        Me.Label10.Text = "###"
        Me.Label10.Visible = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("MS UI Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.Label11.Location = New System.Drawing.Point(503, 70)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(40, 13)
        Me.Label11.TabIndex = 25
        Me.Label11.Text = "在庫："
        Me.Label11.Visible = False
        '
        'Check_stock
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 451)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.ComboBox4)
        Me.Controls.Add(Me.ComboBox3)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.NyukaKakutei_DataGridView)
        Me.MaximizeBox = False
        Me.Name = "Check_stock"
        Me.Text = "詳細確認"
        CType(Me.NyukaKakutei_DataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents NyukaKakutei_DataGridView As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents NameOfProduct_CheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents Partnumber_CheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents MonthComboBox5 As System.Windows.Forms.ComboBox
    Friend WithEvents YearComboBox6 As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents MonthComboBox4 As System.Windows.Forms.ComboBox
    Friend WithEvents YearComboBox3 As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox3 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox4 As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents DataGridViewTextBoxColumn7 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn13 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn15 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn16 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents 単価 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
