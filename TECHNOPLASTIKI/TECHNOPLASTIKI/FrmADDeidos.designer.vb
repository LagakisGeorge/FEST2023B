<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmAddEidos
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmAddEidos))
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.BAROS = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.MON = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.ONO = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.KOD = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtid = New System.Windows.Forms.TextBox
        Me.cmdsave = New System.Windows.Forms.Button
        Me.cmdcancel = New System.Windows.Forms.Button
        Me.n1 = New System.Windows.Forms.Label
        Me.PIECESPERPALLET = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "+.ico")
        Me.ImageList1.Images.SetKeyName(1, "-.ico")
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.PIECESPERPALLET)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.BAROS)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.MON)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.ONO)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.KOD)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtid)
        Me.GroupBox1.Location = New System.Drawing.Point(4, 0)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(554, 192)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'BAROS
        '
        Me.BAROS.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BAROS.Location = New System.Drawing.Point(115, 107)
        Me.BAROS.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.BAROS.MaxLength = 11
        Me.BAROS.Name = "BAROS"
        Me.BAROS.Size = New System.Drawing.Size(99, 24)
        Me.BAROS.TabIndex = 5
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(29, 111)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(55, 17)
        Me.Label7.TabIndex = 102
        Me.Label7.Text = "Βάρος :"
        '
        'MON
        '
        Me.MON.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MON.Location = New System.Drawing.Point(115, 76)
        Me.MON.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MON.MaxLength = 50
        Me.MON.Name = "MON"
        Me.MON.Size = New System.Drawing.Size(432, 24)
        Me.MON.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(1, 84)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(94, 17)
        Me.Label6.TabIndex = 101
        Me.Label6.Text = "Μονάδα Μετ :"
        '
        'ONO
        '
        Me.ONO.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ONO.Location = New System.Drawing.Point(115, 43)
        Me.ONO.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ONO.MaxLength = 100
        Me.ONO.Name = "ONO"
        Me.ONO.Size = New System.Drawing.Size(432, 24)
        Me.ONO.TabIndex = 2
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(1, 47)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(84, 17)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Περιγραφή :"
        '
        'KOD
        '
        Me.KOD.CausesValidation = False
        Me.KOD.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.KOD.Location = New System.Drawing.Point(115, 15)
        Me.KOD.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.KOD.MaxLength = 50
        Me.KOD.Name = "KOD"
        Me.KOD.Size = New System.Drawing.Size(432, 24)
        Me.KOD.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(1, 11)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 17)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "Κωδικός"
        '
        'txtid
        '
        Me.txtid.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtid.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtid.Location = New System.Drawing.Point(115, 15)
        Me.txtid.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtid.Name = "txtid"
        Me.txtid.ReadOnly = True
        Me.txtid.Size = New System.Drawing.Size(103, 24)
        Me.txtid.TabIndex = 100
        Me.txtid.Visible = False
        '
        'cmdsave
        '
        Me.cmdsave.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdsave.Location = New System.Drawing.Point(269, 200)
        Me.cmdsave.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdsave.Name = "cmdsave"
        Me.cmdsave.Size = New System.Drawing.Size(133, 32)
        Me.cmdsave.TabIndex = 8
        Me.cmdsave.Text = "Αποθήκευση"
        Me.cmdsave.UseVisualStyleBackColor = True
        '
        'cmdcancel
        '
        Me.cmdcancel.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdcancel.Location = New System.Drawing.Point(410, 200)
        Me.cmdcancel.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.cmdcancel.Name = "cmdcancel"
        Me.cmdcancel.Size = New System.Drawing.Size(137, 32)
        Me.cmdcancel.TabIndex = 9
        Me.cmdcancel.Text = "Ακυρο"
        Me.cmdcancel.UseVisualStyleBackColor = True
        '
        'n1
        '
        Me.n1.AutoSize = True
        Me.n1.Location = New System.Drawing.Point(76, 235)
        Me.n1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.n1.Name = "n1"
        Me.n1.Size = New System.Drawing.Size(20, 17)
        Me.n1.TabIndex = 10
        Me.n1.Text = "..."
        '
        'PIECESPERPALLET
        '
        Me.PIECESPERPALLET.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PIECESPERPALLET.Location = New System.Drawing.Point(115, 139)
        Me.PIECESPERPALLET.Margin = New System.Windows.Forms.Padding(4)
        Me.PIECESPERPALLET.MaxLength = 11
        Me.PIECESPERPALLET.Name = "PIECESPERPALLET"
        Me.PIECESPERPALLET.Size = New System.Drawing.Size(99, 24)
        Me.PIECESPERPALLET.TabIndex = 103
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(29, 143)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 17)
        Me.Label1.TabIndex = 104
        Me.Label1.Text = "Τεμ/Παλ :"
        '
        'FrmAddEidos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ClientSize = New System.Drawing.Size(557, 245)
        Me.Controls.Add(Me.n1)
        Me.Controls.Add(Me.cmdcancel)
        Me.Controls.Add(Me.cmdsave)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmAddEidos"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Supplier Profile"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ONO As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents KOD As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtid As System.Windows.Forms.TextBox
    Friend WithEvents cmdsave As System.Windows.Forms.Button
    Friend WithEvents cmdcancel As System.Windows.Forms.Button
    Friend WithEvents MON As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents BAROS As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents n1 As System.Windows.Forms.Label
    Friend WithEvents PIECESPERPALLET As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
