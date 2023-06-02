<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form2))
        Me.cFind = New System.Windows.Forms.TextBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.ImageList3 = New System.Windows.Forms.ImageList(Me.components)
        Me.DEL_TIMAGOR = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.delete_label = New System.Windows.Forms.Button
        Me.Proeleysi_Partidas = New System.Windows.Forms.Button
        Me.AnalPartidas = New System.Windows.Forms.Button
        Me.ANALYTICS = New System.Windows.Forms.Button
        Me.SYNTAGES = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.delete = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.GridView1 = New System.Windows.Forms.DataGridView
        Me.Label1 = New System.Windows.Forms.Label
        CType(Me.GridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cFind
        '
        Me.cFind.Location = New System.Drawing.Point(684, 131)
        Me.cFind.Name = "cFind"
        Me.cFind.Size = New System.Drawing.Size(154, 20)
        Me.cFind.TabIndex = 28
        '
        'Button3
        '
        Me.Button3.ImageKey = "reload2.ico"
        Me.Button3.ImageList = Me.ImageList3
        Me.Button3.Location = New System.Drawing.Point(844, 109)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(49, 42)
        Me.Button3.TabIndex = 20
        Me.Button3.UseVisualStyleBackColor = True
        '
        'ImageList3
        '
        Me.ImageList3.ImageStream = CType(resources.GetObject("ImageList3.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList3.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList3.Images.SetKeyName(0, "file.ico")
        Me.ImageList3.Images.SetKeyName(1, "file edit.ico")
        Me.ImageList3.Images.SetKeyName(2, "edit.ico")
        Me.ImageList3.Images.SetKeyName(3, "trash.ico")
        Me.ImageList3.Images.SetKeyName(4, "trash2.ico")
        Me.ImageList3.Images.SetKeyName(5, "binoculars.ico")
        Me.ImageList3.Images.SetKeyName(6, "view-file.ico")
        Me.ImageList3.Images.SetKeyName(7, "find.ico")
        Me.ImageList3.Images.SetKeyName(8, "find_file.ico")
        Me.ImageList3.Images.SetKeyName(9, "printer.ico")
        Me.ImageList3.Images.SetKeyName(10, "print view.ico")
        Me.ImageList3.Images.SetKeyName(11, "reload2.ico")
        Me.ImageList3.Images.SetKeyName(12, "undo.ico")
        Me.ImageList3.Images.SetKeyName(13, "stop.ico")
        Me.ImageList3.Images.SetKeyName(14, "reload_64.png")
        Me.ImageList3.Images.SetKeyName(15, "SEARCH-SMALL.png")
        '
        'DEL_TIMAGOR
        '
        Me.DEL_TIMAGOR.Location = New System.Drawing.Point(606, 108)
        Me.DEL_TIMAGOR.Name = "DEL_TIMAGOR"
        Me.DEL_TIMAGOR.Size = New System.Drawing.Size(80, 23)
        Me.DEL_TIMAGOR.TabIndex = 29
        Me.DEL_TIMAGOR.Text = "Διαγραφή Τιμολ.Αγοράς"
        Me.DEL_TIMAGOR.UseVisualStyleBackColor = True
        Me.DEL_TIMAGOR.Visible = False
        '
        'Button2
        '
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.ImageKey = "find.ico"
        Me.Button2.ImageList = Me.ImageList3
        Me.Button2.Location = New System.Drawing.Point(706, 108)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(100, 23)
        Me.Button2.TabIndex = 27
        Me.Button2.Text = "Αναζήτηση"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'delete_label
        '
        Me.delete_label.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.delete_label.Location = New System.Drawing.Point(454, 112)
        Me.delete_label.Name = "delete_label"
        Me.delete_label.Size = New System.Drawing.Size(80, 39)
        Me.delete_label.TabIndex = 21
        Me.delete_label.Text = "Διαγραφή ετικέττας"
        Me.delete_label.UseVisualStyleBackColor = False
        Me.delete_label.Visible = False
        '
        'Proeleysi_Partidas
        '
        Me.Proeleysi_Partidas.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Proeleysi_Partidas.Location = New System.Drawing.Point(318, 112)
        Me.Proeleysi_Partidas.Name = "Proeleysi_Partidas"
        Me.Proeleysi_Partidas.Size = New System.Drawing.Size(130, 39)
        Me.Proeleysi_Partidas.TabIndex = 26
        Me.Proeleysi_Partidas.Text = "Προελευση Παρτίδας"
        Me.Proeleysi_Partidas.UseVisualStyleBackColor = False
        Me.Proeleysi_Partidas.Visible = False
        '
        'AnalPartidas
        '
        Me.AnalPartidas.Location = New System.Drawing.Point(71, 112)
        Me.AnalPartidas.Name = "AnalPartidas"
        Me.AnalPartidas.Size = New System.Drawing.Size(130, 39)
        Me.AnalPartidas.TabIndex = 25
        Me.AnalPartidas.Text = "Πωλήσεις Παρτίδας"
        Me.AnalPartidas.UseVisualStyleBackColor = True
        Me.AnalPartidas.Visible = False
        '
        'ANALYTICS
        '
        Me.ANALYTICS.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ANALYTICS.Location = New System.Drawing.Point(141, 112)
        Me.ANALYTICS.Name = "ANALYTICS"
        Me.ANALYTICS.Size = New System.Drawing.Size(130, 39)
        Me.ANALYTICS.TabIndex = 24
        Me.ANALYTICS.Text = "Ανάλυση"
        Me.ANALYTICS.UseVisualStyleBackColor = False
        Me.ANALYTICS.Visible = False
        '
        'SYNTAGES
        '
        Me.SYNTAGES.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.SYNTAGES.Location = New System.Drawing.Point(182, 112)
        Me.SYNTAGES.Name = "SYNTAGES"
        Me.SYNTAGES.Size = New System.Drawing.Size(130, 39)
        Me.SYNTAGES.TabIndex = 23
        Me.SYNTAGES.Text = "Συνταγή"
        Me.SYNTAGES.UseVisualStyleBackColor = False
        Me.SYNTAGES.Visible = False
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(540, 112)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(60, 39)
        Me.Button1.TabIndex = 22
        Me.Button1.Text = "Σε Excel"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'delete
        '
        Me.delete.Location = New System.Drawing.Point(606, 129)
        Me.delete.Name = "delete"
        Me.delete.Size = New System.Drawing.Size(80, 22)
        Me.delete.TabIndex = 19
        Me.delete.Text = "Διαγραφή"
        Me.delete.UseVisualStyleBackColor = True
        Me.delete.Visible = False
        '
        'cmdCancel
        '
        Me.cmdCancel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.Location = New System.Drawing.Point(892, 109)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(80, 42)
        Me.cmdCancel.TabIndex = 16
        Me.cmdCancel.Text = "Αποθήκευση Εξοδος"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(17, 125)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(39, 13)
        Me.Label2.TabIndex = 18
        Me.Label2.Text = "Label2"
        '
        'GridView1
        '
        Me.GridView1.AllowUserToDeleteRows = False
        Me.GridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridView1.Location = New System.Drawing.Point(18, 164)
        Me.GridView1.Name = "GridView1"
        Me.GridView1.RowTemplate.Height = 24
        Me.GridView1.Size = New System.Drawing.Size(954, 198)
        Me.GridView1.TabIndex = 17
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(37, 76)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 13)
        Me.Label1.TabIndex = 30
        Me.Label1.Text = "Label1"
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(989, 470)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cFind)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.DEL_TIMAGOR)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.delete_label)
        Me.Controls.Add(Me.Proeleysi_Partidas)
        Me.Controls.Add(Me.AnalPartidas)
        Me.Controls.Add(Me.ANALYTICS)
        Me.Controls.Add(Me.SYNTAGES)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.delete)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GridView1)
        Me.Name = "Form2"
        Me.Text = "Form2"
        CType(Me.GridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cFind As System.Windows.Forms.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ImageList3 As System.Windows.Forms.ImageList
    Friend WithEvents DEL_TIMAGOR As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents delete_label As System.Windows.Forms.Button
    Friend WithEvents Proeleysi_Partidas As System.Windows.Forms.Button
    Friend WithEvents AnalPartidas As System.Windows.Forms.Button
    Friend WithEvents ANALYTICS As System.Windows.Forms.Button
    Friend WithEvents SYNTAGES As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents delete As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
