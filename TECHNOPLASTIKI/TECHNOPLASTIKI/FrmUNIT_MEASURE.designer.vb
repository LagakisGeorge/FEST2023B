<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ergates
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ergates))
        Me.ImageList3 = New System.Windows.Forms.ImageList(Me.components)
        Me.Label1 = New System.Windows.Forms.Label
        Me.GridView1 = New System.Windows.Forms.DataGridView
        Me.Label2 = New System.Windows.Forms.Label
        Me.deleteYLIKA = New System.Windows.Forms.Button
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Button1 = New System.Windows.Forms.Button
        Me.SYNTAGES = New System.Windows.Forms.Button
        Me.ANALYTICS = New System.Windows.Forms.Button
        Me.AnalPartidas = New System.Windows.Forms.Button
        Me.Proeleysi_Partidas = New System.Windows.Forms.Button
        Me.delete_label = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.DEL_TIMAGOR = New System.Windows.Forms.Button
        Me.cFind = New System.Windows.Forms.TextBox
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdRefresh = New System.Windows.Forms.Button
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.ImageList4 = New System.Windows.Forms.ImageList(Me.components)
        Me.n1 = New System.Windows.Forms.Label
        Me.deleteTIMPOL = New System.Windows.Forms.Button
        Me.DELETE_APOG_PART = New System.Windows.Forms.Button
        Me.kinhseis = New System.Windows.Forms.Button
        Me.CMDEPANEKTYPOSI = New System.Windows.Forms.Button
        Me.cmdEdit = New System.Windows.Forms.Button
        Me.analtimpol = New System.Windows.Forms.Button
        Me.symorfosi = New System.Windows.Forms.Button
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.dior_pel = New System.Windows.Forms.Button
        Me.add_pel = New System.Windows.Forms.Button
        CType(Me.GridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
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
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(1, 321)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(0, 15)
        Me.Label1.TabIndex = 5
        Me.Label1.Visible = False
        '
        'GridView1
        '
        Me.GridView1.AllowUserToDeleteRows = False
        Me.GridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridView1.Location = New System.Drawing.Point(6, 73)
        Me.GridView1.Name = "GridView1"
        Me.GridView1.RowTemplate.Height = 24
        Me.GridView1.Size = New System.Drawing.Size(908, 508)
        Me.GridView1.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(28, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(45, 15)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Label2"
        '
        'deleteYLIKA
        '
        Me.deleteYLIKA.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.deleteYLIKA.ImageKey = "-.ico"
        Me.deleteYLIKA.ImageList = Me.ImageList1
        Me.deleteYLIKA.Location = New System.Drawing.Point(572, 24)
        Me.deleteYLIKA.Name = "deleteYLIKA"
        Me.deleteYLIKA.Size = New System.Drawing.Size(94, 37)
        Me.deleteYLIKA.TabIndex = 6
        Me.deleteYLIKA.Text = "Διαγραφή"
        Me.deleteYLIKA.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.deleteYLIKA.UseVisualStyleBackColor = True
        Me.deleteYLIKA.Visible = False
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "+.ico")
        Me.ImageList1.Images.SetKeyName(1, "-.ico")
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(341, 24)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(60, 39)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Σε Excel"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'SYNTAGES
        '
        Me.SYNTAGES.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.SYNTAGES.Location = New System.Drawing.Point(129, 22)
        Me.SYNTAGES.Name = "SYNTAGES"
        Me.SYNTAGES.Size = New System.Drawing.Size(114, 39)
        Me.SYNTAGES.TabIndex = 8
        Me.SYNTAGES.Text = "Συνταγή"
        Me.SYNTAGES.UseVisualStyleBackColor = False
        Me.SYNTAGES.Visible = False
        '
        'ANALYTICS
        '
        Me.ANALYTICS.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ANALYTICS.Location = New System.Drawing.Point(6, 24)
        Me.ANALYTICS.Name = "ANALYTICS"
        Me.ANALYTICS.Size = New System.Drawing.Size(130, 39)
        Me.ANALYTICS.TabIndex = 9
        Me.ANALYTICS.Text = "Ανάλυση"
        Me.ANALYTICS.UseVisualStyleBackColor = False
        Me.ANALYTICS.Visible = False
        '
        'AnalPartidas
        '
        Me.AnalPartidas.Location = New System.Drawing.Point(31, 24)
        Me.AnalPartidas.Name = "AnalPartidas"
        Me.AnalPartidas.Size = New System.Drawing.Size(130, 39)
        Me.AnalPartidas.TabIndex = 10
        Me.AnalPartidas.Text = "Πωλήσεις Παρτίδας"
        Me.AnalPartidas.UseVisualStyleBackColor = True
        Me.AnalPartidas.Visible = False
        '
        'Proeleysi_Partidas
        '
        Me.Proeleysi_Partidas.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.Proeleysi_Partidas.Location = New System.Drawing.Point(179, 24)
        Me.Proeleysi_Partidas.Name = "Proeleysi_Partidas"
        Me.Proeleysi_Partidas.Size = New System.Drawing.Size(130, 39)
        Me.Proeleysi_Partidas.TabIndex = 11
        Me.Proeleysi_Partidas.Text = "Προελευση Παρτίδας"
        Me.Proeleysi_Partidas.UseVisualStyleBackColor = False
        Me.Proeleysi_Partidas.Visible = False
        '
        'delete_label
        '
        Me.delete_label.BackColor = System.Drawing.Color.Lime
        Me.delete_label.Location = New System.Drawing.Point(260, 25)
        Me.delete_label.Name = "delete_label"
        Me.delete_label.Size = New System.Drawing.Size(80, 39)
        Me.delete_label.TabIndex = 6
        Me.delete_label.Text = "Διαγραφή ετικέττας"
        Me.delete_label.UseVisualStyleBackColor = False
        Me.delete_label.Visible = False
        '
        'Button2
        '
        Me.Button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.ImageKey = "find.ico"
        Me.Button2.ImageList = Me.ImageList3
        Me.Button2.Location = New System.Drawing.Point(685, 18)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(100, 23)
        Me.Button2.TabIndex = 12
        Me.Button2.Text = "Αναζήτηση"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'DEL_TIMAGOR
        '
        Me.DEL_TIMAGOR.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.DEL_TIMAGOR.ImageKey = "-.ico"
        Me.DEL_TIMAGOR.ImageList = Me.ImageList1
        Me.DEL_TIMAGOR.Location = New System.Drawing.Point(514, 25)
        Me.DEL_TIMAGOR.Name = "DEL_TIMAGOR"
        Me.DEL_TIMAGOR.Size = New System.Drawing.Size(94, 37)
        Me.DEL_TIMAGOR.TabIndex = 15
        Me.DEL_TIMAGOR.Text = "Διαγραφή Τιμολ.Αγοράς"
        Me.DEL_TIMAGOR.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.DEL_TIMAGOR.UseVisualStyleBackColor = True
        Me.DEL_TIMAGOR.Visible = False
        '
        'cFind
        '
        Me.cFind.Location = New System.Drawing.Point(672, 41)
        Me.cFind.Name = "cFind"
        Me.cFind.Size = New System.Drawing.Size(140, 20)
        Me.cFind.TabIndex = 13
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(863, 30)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(51, 33)
        Me.cmdCancel.TabIndex = 16
        Me.cmdCancel.Text = "Εξοδος"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdRefresh
        '
        Me.cmdRefresh.ImageKey = "reload2.ico"
        Me.cmdRefresh.ImageList = Me.ImageList3
        Me.cmdRefresh.Location = New System.Drawing.Point(822, 30)
        Me.cmdRefresh.Name = "cmdRefresh"
        Me.cmdRefresh.Size = New System.Drawing.Size(35, 33)
        Me.cmdRefresh.TabIndex = 17
        Me.cmdRefresh.UseVisualStyleBackColor = True
        '
        'cmdAdd
        '
        Me.cmdAdd.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.cmdAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAdd.ImageKey = "+.ico"
        Me.cmdAdd.ImageList = Me.ImageList1
        Me.cmdAdd.Location = New System.Drawing.Point(407, 25)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(101, 37)
        Me.cmdAdd.TabIndex = 18
        Me.cmdAdd.Text = "Νέα Εγγραφή"
        Me.cmdAdd.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdAdd.UseVisualStyleBackColor = False
        Me.cmdAdd.Visible = False
        '
        'ImageList4
        '
        Me.ImageList4.ImageStream = CType(resources.GetObject("ImageList4.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList4.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList4.Images.SetKeyName(0, "print_64.png")
        Me.ImageList4.Images.SetKeyName(1, "cancel_48.png")
        Me.ImageList4.Images.SetKeyName(2, "search_48.png")
        Me.ImageList4.Images.SetKeyName(3, "Delete.png")
        Me.ImageList4.Images.SetKeyName(4, "Edit.png")
        Me.ImageList4.Images.SetKeyName(5, "Add.png")
        Me.ImageList4.Images.SetKeyName(6, "lock_48.png")
        Me.ImageList4.Images.SetKeyName(7, "home.ico")
        Me.ImageList4.Images.SetKeyName(8, "home_64.png")
        '
        'n1
        '
        Me.n1.AutoSize = True
        Me.n1.Location = New System.Drawing.Point(694, 8)
        Me.n1.Name = "n1"
        Me.n1.Size = New System.Drawing.Size(22, 15)
        Me.n1.TabIndex = 19
        Me.n1.Text = "....."
        '
        'deleteTIMPOL
        '
        Me.deleteTIMPOL.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.deleteTIMPOL.ImageKey = "-.ico"
        Me.deleteTIMPOL.ImageList = Me.ImageList1
        Me.deleteTIMPOL.Location = New System.Drawing.Point(540, 24)
        Me.deleteTIMPOL.Name = "deleteTIMPOL"
        Me.deleteTIMPOL.Size = New System.Drawing.Size(94, 37)
        Me.deleteTIMPOL.TabIndex = 20
        Me.deleteTIMPOL.Text = "Διαγραφή Τιμολ.Πώλησης"
        Me.deleteTIMPOL.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.deleteTIMPOL.UseVisualStyleBackColor = True
        Me.deleteTIMPOL.Visible = False
        '
        'DELETE_APOG_PART
        '
        Me.DELETE_APOG_PART.Location = New System.Drawing.Point(43, 23)
        Me.DELETE_APOG_PART.Name = "DELETE_APOG_PART"
        Me.DELETE_APOG_PART.Size = New System.Drawing.Size(130, 39)
        Me.DELETE_APOG_PART.TabIndex = 21
        Me.DELETE_APOG_PART.Text = "Διαγρ.Απογ. Παρτίδας"
        Me.DELETE_APOG_PART.UseVisualStyleBackColor = True
        Me.DELETE_APOG_PART.Visible = False
        '
        'kinhseis
        '
        Me.kinhseis.Location = New System.Drawing.Point(401, 24)
        Me.kinhseis.Name = "kinhseis"
        Me.kinhseis.Size = New System.Drawing.Size(63, 38)
        Me.kinhseis.TabIndex = 22
        Me.kinhseis.Text = "Κινήσεις"
        Me.kinhseis.UseVisualStyleBackColor = True
        Me.kinhseis.Visible = False
        '
        'CMDEPANEKTYPOSI
        '
        Me.CMDEPANEKTYPOSI.BackColor = System.Drawing.Color.Lime
        Me.CMDEPANEKTYPOSI.Location = New System.Drawing.Point(398, 26)
        Me.CMDEPANEKTYPOSI.Name = "CMDEPANEKTYPOSI"
        Me.CMDEPANEKTYPOSI.Size = New System.Drawing.Size(91, 36)
        Me.CMDEPANEKTYPOSI.TabIndex = 23
        Me.CMDEPANEKTYPOSI.Text = "Επανεκτύπωση ετικέτας"
        Me.CMDEPANEKTYPOSI.UseVisualStyleBackColor = False
        Me.CMDEPANEKTYPOSI.Visible = False
        '
        'cmdEdit
        '
        Me.cmdEdit.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.cmdEdit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEdit.ImageKey = "Edit.png"
        Me.cmdEdit.ImageList = Me.ImageList4
        Me.cmdEdit.Location = New System.Drawing.Point(6, 22)
        Me.cmdEdit.Name = "cmdEdit"
        Me.cmdEdit.Size = New System.Drawing.Size(117, 39)
        Me.cmdEdit.TabIndex = 24
        Me.cmdEdit.Text = "Διόρθωση εγγραφής"
        Me.cmdEdit.UseVisualStyleBackColor = False
        Me.cmdEdit.Visible = False
        '
        'analtimpol
        '
        Me.analtimpol.Location = New System.Drawing.Point(401, 24)
        Me.analtimpol.Name = "analtimpol"
        Me.analtimpol.Size = New System.Drawing.Size(88, 39)
        Me.analtimpol.TabIndex = 25
        Me.analtimpol.Text = "Ανάλυση Παραστατικού"
        Me.analtimpol.UseVisualStyleBackColor = True
        Me.analtimpol.Visible = False
        '
        'symorfosi
        '
        Me.symorfosi.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.symorfosi.Location = New System.Drawing.Point(210, 25)
        Me.symorfosi.Name = "symorfosi"
        Me.symorfosi.Size = New System.Drawing.Size(113, 38)
        Me.symorfosi.TabIndex = 26
        Me.symorfosi.Text = "Δήλωση Συμμόρφωσης"
        Me.symorfosi.UseVisualStyleBackColor = False
        Me.symorfosi.Visible = False
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(73, 41)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(131, 21)
        Me.ComboBox1.TabIndex = 27
        Me.ComboBox1.Visible = False
        '
        'dior_pel
        '
        Me.dior_pel.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.dior_pel.Location = New System.Drawing.Point(444, 26)
        Me.dior_pel.Name = "dior_pel"
        Me.dior_pel.Size = New System.Drawing.Size(103, 35)
        Me.dior_pel.TabIndex = 28
        Me.dior_pel.Text = "ΔΙΟΡΘΩΣΗ"
        Me.dior_pel.UseVisualStyleBackColor = False
        Me.dior_pel.Visible = False
        '
        'add_pel
        '
        Me.add_pel.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.add_pel.Location = New System.Drawing.Point(157, 22)
        Me.add_pel.Name = "add_pel"
        Me.add_pel.Size = New System.Drawing.Size(97, 39)
        Me.add_pel.TabIndex = 29
        Me.add_pel.Text = "Νέα Εγγραφή"
        Me.add_pel.UseVisualStyleBackColor = False
        Me.add_pel.Visible = False
        '
        'ergates
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ClientSize = New System.Drawing.Size(1089, 853)
        Me.Controls.Add(Me.add_pel)
        Me.Controls.Add(Me.dior_pel)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.symorfosi)
        Me.Controls.Add(Me.analtimpol)
        Me.Controls.Add(Me.cmdEdit)
        Me.Controls.Add(Me.CMDEPANEKTYPOSI)
        Me.Controls.Add(Me.kinhseis)
        Me.Controls.Add(Me.DELETE_APOG_PART)
        Me.Controls.Add(Me.deleteTIMPOL)
        Me.Controls.Add(Me.n1)
        Me.Controls.Add(Me.cmdAdd)
        Me.Controls.Add(Me.cmdRefresh)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cFind)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DEL_TIMAGOR)
        Me.Controls.Add(Me.AnalPartidas)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.deleteYLIKA)
        Me.Controls.Add(Me.ANALYTICS)
        Me.Controls.Add(Me.SYNTAGES)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.delete_label)
        Me.Controls.Add(Me.Proeleysi_Partidas)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GridView1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ergates"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Eργάτες"
        CType(Me.GridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ImageList3 As System.Windows.Forms.ImageList
    Friend WithEvents GridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents deleteYLIKA As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents SYNTAGES As System.Windows.Forms.Button
    Friend WithEvents ANALYTICS As System.Windows.Forms.Button
    Friend WithEvents AnalPartidas As System.Windows.Forms.Button
    Friend WithEvents Proeleysi_Partidas As System.Windows.Forms.Button
    Friend WithEvents delete_label As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents DEL_TIMAGOR As System.Windows.Forms.Button
    Friend WithEvents cFind As System.Windows.Forms.TextBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdRefresh As System.Windows.Forms.Button
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents ImageList4 As System.Windows.Forms.ImageList
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents n1 As System.Windows.Forms.Label
    Friend WithEvents deleteTIMPOL As System.Windows.Forms.Button
    Friend WithEvents DELETE_APOG_PART As System.Windows.Forms.Button
    Friend WithEvents kinhseis As System.Windows.Forms.Button
    Friend WithEvents CMDEPANEKTYPOSI As System.Windows.Forms.Button
    Friend WithEvents cmdEdit As System.Windows.Forms.Button
    Friend WithEvents analtimpol As System.Windows.Forms.Button
    Friend WithEvents symorfosi As System.Windows.Forms.Button
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents dior_pel As System.Windows.Forms.Button
    Friend WithEvents add_pel As System.Windows.Forms.Button
End Class
