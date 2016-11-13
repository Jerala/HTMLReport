<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Events
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ButtonForQueries = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBoxForQueries = New System.Windows.Forms.TextBox()
        Me.ButtonForChangeDataBase = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(150, 167)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(136, 84)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "LoadToHTML"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ButtonForQueries)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.TextBoxForQueries)
        Me.GroupBox1.Location = New System.Drawing.Point(13, 13)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(284, 116)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Add SQL query"
        '
        'ButtonForQueries
        '
        Me.ButtonForQueries.Location = New System.Drawing.Point(137, 78)
        Me.ButtonForQueries.Name = "ButtonForQueries"
        Me.ButtonForQueries.Size = New System.Drawing.Size(75, 23)
        Me.ButtonForQueries.TabIndex = 3
        Me.ButtonForQueries.Text = "Add query"
        Me.ButtonForQueries.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(119, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Enter your SQL queries:"
        '
        'TextBoxForQueries
        '
        Me.TextBoxForQueries.Location = New System.Drawing.Point(9, 53)
        Me.TextBoxForQueries.Name = "TextBoxForQueries"
        Me.TextBoxForQueries.Size = New System.Drawing.Size(203, 20)
        Me.TextBoxForQueries.TabIndex = 0
        Me.TextBoxForQueries.Text = "SELECT * FROM ForGraphic"
        '
        'ButtonForChangeDataBase
        '
        Me.ButtonForChangeDataBase.Location = New System.Drawing.Point(13, 167)
        Me.ButtonForChangeDataBase.Name = "ButtonForChangeDataBase"
        Me.ButtonForChangeDataBase.Size = New System.Drawing.Size(100, 45)
        Me.ButtonForChangeDataBase.TabIndex = 10
        Me.ButtonForChangeDataBase.Text = "Change DataBase"
        Me.ButtonForChangeDataBase.UseVisualStyleBackColor = True
        '
        'Events
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(336, 277)
        Me.Controls.Add(Me.ButtonForChangeDataBase)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Events"
        Me.Text = "FromDataBaseToHTML"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents TextBoxForQueries As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents ButtonForQueries As Button
    Friend WithEvents ButtonForChangeDataBase As Button
End Class
