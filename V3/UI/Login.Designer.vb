<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Login
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
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Button_Exit = New System.Windows.Forms.Button()
        Me.Button_Login = New System.Windows.Forms.Button()
        Me.tbx_Password = New System.Windows.Forms.TextBox()
        Me.tbx_ID = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(82, 95)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(191, 12)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "（半角アルファベット又は数字のみ入力）"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(82, 47)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(191, 12)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "（半角アルファベット又は数字のみ入力）"
        '
        'Button_Exit
        '
        Me.Button_Exit.Location = New System.Drawing.Point(198, 115)
        Me.Button_Exit.Name = "Button_Exit"
        Me.Button_Exit.Size = New System.Drawing.Size(75, 23)
        Me.Button_Exit.TabIndex = 19
        Me.Button_Exit.Text = "終了"
        Me.Button_Exit.UseVisualStyleBackColor = True
        '
        'Button_Login
        '
        Me.Button_Login.Location = New System.Drawing.Point(84, 115)
        Me.Button_Login.Name = "Button_Login"
        Me.Button_Login.Size = New System.Drawing.Size(75, 23)
        Me.Button_Login.TabIndex = 18
        Me.Button_Login.Text = "ログイン"
        Me.Button_Login.UseVisualStyleBackColor = True
        '
        'tbx_Password
        '
        Me.tbx_Password.Location = New System.Drawing.Point(84, 69)
        Me.tbx_Password.Name = "tbx_Password"
        Me.tbx_Password.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.tbx_Password.Size = New System.Drawing.Size(189, 19)
        Me.tbx_Password.TabIndex = 17
        '
        'tbx_ID
        '
        Me.tbx_ID.Location = New System.Drawing.Point(84, 23)
        Me.tbx_ID.Name = "tbx_ID"
        Me.tbx_ID.Size = New System.Drawing.Size(189, 19)
        Me.tbx_ID.TabIndex = 16
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(21, 72)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(54, 12)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Password"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(21, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(16, 12)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "ID"
        '
        'Login
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(312, 162)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button_Exit)
        Me.Controls.Add(Me.Button_Login)
        Me.Controls.Add(Me.tbx_Password)
        Me.Controls.Add(Me.tbx_ID)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "Login"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ID入力"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Button_Exit As System.Windows.Forms.Button
    Friend WithEvents Button_Login As System.Windows.Forms.Button
    Friend WithEvents tbx_Password As System.Windows.Forms.TextBox
    Friend WithEvents tbx_ID As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label

End Class
