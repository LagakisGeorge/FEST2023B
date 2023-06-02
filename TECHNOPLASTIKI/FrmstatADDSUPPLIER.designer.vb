<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Airport
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Airport))
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.DTCheckin = New System.Windows.Forms.DateTimePicker
        Me.DTCheckout = New System.Windows.Forms.DateTimePicker
        Me.DtAirAfixi = New System.Windows.Forms.DateTimePicker
        Me.Label9 = New System.Windows.Forms.Label
        Me.dtAirAnax = New System.Windows.Forms.DateTimePicker
        Me.Label10 = New System.Windows.Forms.Label
        Me.send = New System.Windows.Forms.Button
        Me.txtMessage = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.txtTo = New System.Windows.Forms.TextBox
        Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox
        Me.CheckedListBox2 = New System.Windows.Forms.CheckedListBox
        Me.cmdsave = New System.Windows.Forms.Button
        Me.cmdcancel = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.PTHSHAN = New System.Windows.Forms.TextBox
        Me.PTHSHAF = New System.Windows.Forms.TextBox
        Me.cAirport = New System.Windows.Forms.ComboBox
        Me.toexcel = New System.Windows.Forms.Button
        Me.Subject = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.SQLBuild = New System.Windows.Forms.Button
        Me.sqlText = New System.Windows.Forms.TextBox
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.mAttachment = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "+.ico")
        Me.ImageList1.Images.SetKeyName(1, "-.ico")
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(219, 20)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(70, 13)
        Me.Label5.TabIndex = 108
        Me.Label5.Text = "Αεροδρόμιο :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(3, 60)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(42, 13)
        Me.Label4.TabIndex = 109
        Me.Label4.Text = "Αφιξη :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(3, 103)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 13)
        Me.Label1.TabIndex = 110
        Me.Label1.Text = "Αναχώρηση :"
        '
        'DTCheckin
        '
        Me.DTCheckin.Checked = False
        Me.DTCheckin.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTCheckin.Location = New System.Drawing.Point(83, 59)
        Me.DTCheckin.Name = "DTCheckin"
        Me.DTCheckin.ShowCheckBox = True
        Me.DTCheckin.Size = New System.Drawing.Size(160, 20)
        Me.DTCheckin.TabIndex = 111
        '
        'DTCheckout
        '
        Me.DTCheckout.Checked = False
        Me.DTCheckout.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DTCheckout.Location = New System.Drawing.Point(84, 103)
        Me.DTCheckout.Name = "DTCheckout"
        Me.DTCheckout.ShowCheckBox = True
        Me.DTCheckout.Size = New System.Drawing.Size(160, 20)
        Me.DTCheckout.TabIndex = 112
        '
        'DtAirAfixi
        '
        Me.DtAirAfixi.Checked = False
        Me.DtAirAfixi.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DtAirAfixi.Location = New System.Drawing.Point(249, 60)
        Me.DtAirAfixi.Name = "DtAirAfixi"
        Me.DtAirAfixi.ShowCheckBox = True
        Me.DtAirAfixi.Size = New System.Drawing.Size(105, 20)
        Me.DtAirAfixi.TabIndex = 113
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(246, 44)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(108, 13)
        Me.Label9.TabIndex = 114
        Me.Label9.Text = "Αφιξη σε Αεροδρόμιο"
        '
        'dtAirAnax
        '
        Me.dtAirAnax.Checked = False
        Me.dtAirAnax.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtAirAnax.Location = New System.Drawing.Point(250, 103)
        Me.dtAirAnax.Name = "dtAirAnax"
        Me.dtAirAnax.ShowCheckBox = True
        Me.dtAirAnax.Size = New System.Drawing.Size(104, 20)
        Me.dtAirAnax.TabIndex = 115
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(246, 87)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(144, 13)
        Me.Label10.TabIndex = 116
        Me.Label10.Text = "Αναχώρηση από Αεροδρόμιο"
        '
        'send
        '
        Me.send.Location = New System.Drawing.Point(289, 443)
        Me.send.Name = "send"
        Me.send.Size = New System.Drawing.Size(122, 24)
        Me.send.TabIndex = 117
        Me.send.Text = "Αποστολή"
        Me.send.UseVisualStyleBackColor = True
        '
        'txtMessage
        '
        Me.txtMessage.Location = New System.Drawing.Point(94, 214)
        Me.txtMessage.Multiline = True
        Me.txtMessage.Name = "txtMessage"
        Me.txtMessage.Size = New System.Drawing.Size(317, 219)
        Me.txtMessage.TabIndex = 118
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(96, 652)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(62, 13)
        Me.Label11.TabIndex = 119
        Me.Label11.Text = "Πρός(email)"
        Me.Label11.Visible = False
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(9, 296)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(79, 13)
        Me.Label12.TabIndex = 120
        Me.Label12.Text = "Κείμενο (email)"
        '
        'txtTo
        '
        Me.txtTo.Location = New System.Drawing.Point(250, 649)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.Size = New System.Drawing.Size(176, 20)
        Me.txtTo.TabIndex = 121
        Me.txtTo.Visible = False
        '
        'CheckedListBox1
        '
        Me.CheckedListBox1.FormattingEnabled = True
        Me.CheckedListBox1.Items.AddRange(New Object() {"Εθνικό", "Διεθνές", "Εθνικό Σπουδαστικό", "Διεθνές Σπουδαστικό", "Short & Green", "KIDDO", "Animation"})
        Me.CheckedListBox1.Location = New System.Drawing.Point(494, 19)
        Me.CheckedListBox1.Name = "CheckedListBox1"
        Me.CheckedListBox1.Size = New System.Drawing.Size(194, 244)
        Me.CheckedListBox1.TabIndex = 123
        '
        'CheckedListBox2
        '
        Me.CheckedListBox2.FormattingEnabled = True
        Me.CheckedListBox2.Items.AddRange(New Object() {"Χορηγός", "Συνεργάτης", "Προσωπικό", "Σωματείο/Οργανισμός", "Masterclass/Panel", "Δημοσιογράφος", "VIP", "Pitching Lab", "Industry", "Κριτική Εθνικό", "Κριτική Διεθνές", "Κριτική Εθνικό Σπουδαστικό", "Κριτική Διεθνές Σπουδαστικό", "Κριτική Short & Green", "Κριτική KIDDO", "Κριτική Animation"})
        Me.CheckedListBox2.Location = New System.Drawing.Point(694, 19)
        Me.CheckedListBox2.Name = "CheckedListBox2"
        Me.CheckedListBox2.Size = New System.Drawing.Size(333, 244)
        Me.CheckedListBox2.TabIndex = 124
        '
        'cmdsave
        '
        Me.cmdsave.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdsave.Location = New System.Drawing.Point(494, 441)
        Me.cmdsave.Name = "cmdsave"
        Me.cmdsave.Size = New System.Drawing.Size(100, 26)
        Me.cmdsave.TabIndex = 8
        Me.cmdsave.Text = "Υπολογισμός"
        Me.cmdsave.UseVisualStyleBackColor = True
        '
        'cmdcancel
        '
        Me.cmdcancel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdcancel.Location = New System.Drawing.Point(924, 441)
        Me.cmdcancel.Name = "cmdcancel"
        Me.cmdcancel.Size = New System.Drawing.Size(103, 26)
        Me.cmdcancel.TabIndex = 9
        Me.cmdcancel.Text = "Ακυρο"
        Me.cmdcancel.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(94, 443)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(122, 24)
        Me.Button1.TabIndex = 125
        Me.Button1.Text = "Επιλογή Συνημμένου"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.mAttachment)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.PTHSHAN)
        Me.GroupBox1.Controls.Add(Me.PTHSHAF)
        Me.GroupBox1.Controls.Add(Me.cAirport)
        Me.GroupBox1.Controls.Add(Me.toexcel)
        Me.GroupBox1.Controls.Add(Me.Subject)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.SQLBuild)
        Me.GroupBox1.Controls.Add(Me.sqlText)
        Me.GroupBox1.Controls.Add(Me.DataGridView1)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.cmdcancel)
        Me.GroupBox1.Controls.Add(Me.cmdsave)
        Me.GroupBox1.Controls.Add(Me.CheckedListBox2)
        Me.GroupBox1.Controls.Add(Me.CheckedListBox1)
        Me.GroupBox1.Controls.Add(Me.txtTo)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.txtMessage)
        Me.GroupBox1.Controls.Add(Me.send)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.dtAirAnax)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.DtAirAfixi)
        Me.GroupBox1.Controls.Add(Me.DTCheckout)
        Me.GroupBox1.Controls.Add(Me.DTCheckin)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Location = New System.Drawing.Point(3, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1068, 722)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(385, 44)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 13)
        Me.Label2.TabIndex = 136
        Me.Label2.Text = "Αριθ,Πτήσης"
        '
        'PTHSHAN
        '
        Me.PTHSHAN.Location = New System.Drawing.Point(370, 103)
        Me.PTHSHAN.Name = "PTHSHAN"
        Me.PTHSHAN.Size = New System.Drawing.Size(100, 20)
        Me.PTHSHAN.TabIndex = 135
        '
        'PTHSHAF
        '
        Me.PTHSHAF.Location = New System.Drawing.Point(370, 57)
        Me.PTHSHAF.Name = "PTHSHAF"
        Me.PTHSHAF.Size = New System.Drawing.Size(100, 20)
        Me.PTHSHAF.TabIndex = 134
        '
        'cAirport
        '
        Me.cAirport.FormattingEnabled = True
        Me.cAirport.Items.AddRange(New Object() {"Θεσσαλονίκης", "Καβάλας", "Αλεξανδρούπολης", "Αμυγδαλεώνα"})
        Me.cAirport.Location = New System.Drawing.Point(319, 20)
        Me.cAirport.Name = "cAirport"
        Me.cAirport.Size = New System.Drawing.Size(151, 21)
        Me.cAirport.TabIndex = 133
        '
        'toexcel
        '
        Me.toexcel.Location = New System.Drawing.Point(607, 441)
        Me.toexcel.Name = "toexcel"
        Me.toexcel.Size = New System.Drawing.Size(81, 26)
        Me.toexcel.TabIndex = 132
        Me.toexcel.Text = "Σε Excel"
        Me.toexcel.UseVisualStyleBackColor = True
        '
        'Subject
        '
        Me.Subject.Location = New System.Drawing.Point(235, 155)
        Me.Subject.Name = "Subject"
        Me.Subject.Size = New System.Drawing.Size(176, 20)
        Me.Subject.TabIndex = 131
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(91, 158)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(36, 13)
        Me.Label3.TabIndex = 130
        Me.Label3.Text = "Θέμα:"
        '
        'SQLBuild
        '
        Me.SQLBuild.Location = New System.Drawing.Point(494, 392)
        Me.SQLBuild.Name = "SQLBuild"
        Me.SQLBuild.Size = New System.Drawing.Size(100, 26)
        Me.SQLBuild.TabIndex = 128
        Me.SQLBuild.Text = "Δόμηση SQL"
        Me.SQLBuild.UseVisualStyleBackColor = True
        '
        'sqlText
        '
        Me.sqlText.Location = New System.Drawing.Point(494, 307)
        Me.sqlText.Multiline = True
        Me.sqlText.Name = "sqlText"
        Me.sqlText.Size = New System.Drawing.Size(533, 79)
        Me.sqlText.TabIndex = 127
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(94, 490)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(933, 109)
        Me.DataGridView1.TabIndex = 126
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'mAttachment
        '
        Me.mAttachment.AutoSize = True
        Me.mAttachment.Location = New System.Drawing.Point(91, 470)
        Me.mAttachment.Name = "mAttachment"
        Me.mAttachment.Size = New System.Drawing.Size(16, 13)
        Me.mAttachment.TabIndex = 137
        Me.mAttachment.Text = "..."
        '
        'Airport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ClientSize = New System.Drawing.Size(1438, 851)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Airport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Προσκεκλημένος"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DTCheckin As System.Windows.Forms.DateTimePicker
    Friend WithEvents DTCheckout As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtAirAfixi As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents dtAirAnax As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents send As System.Windows.Forms.Button
    Friend WithEvents txtMessage As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents txtTo As System.Windows.Forms.TextBox
    Friend WithEvents CheckedListBox1 As System.Windows.Forms.CheckedListBox
    Friend WithEvents CheckedListBox2 As System.Windows.Forms.CheckedListBox
    Friend WithEvents cmdsave As System.Windows.Forms.Button
    Friend WithEvents cmdcancel As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents sqlText As System.Windows.Forms.TextBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents SQLBuild As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Subject As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents toexcel As System.Windows.Forms.Button
    Friend WithEvents cAirport As System.Windows.Forms.ComboBox
    Friend WithEvents PTHSHAF As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents PTHSHAN As System.Windows.Forms.TextBox
    Friend WithEvents mAttachment As System.Windows.Forms.Label
End Class
