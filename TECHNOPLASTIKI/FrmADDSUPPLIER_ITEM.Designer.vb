<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmADDSUPPLIER_ITEM
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtEmail = New System.Windows.Forms.TextBox
        Me.cmdsave = New System.Windows.Forms.Button
        Me.cmdcancel = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtDie = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.N1 = New System.Windows.Forms.Label
        Me.txtName = New System.Windows.Forms.TextBox
        Me.txtTHL = New System.Windows.Forms.TextBox
        Me.RANK = New System.Windows.Forms.TextBox
        Me.ComboCAT = New System.Windows.Forms.ComboBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 35)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Προτεραιότητα :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(109, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Ονομα Ξενοδοχείου :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 85)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Email :"
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(126, 85)
        Me.txtEmail.MaxLength = 50
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(293, 21)
        Me.txtEmail.TabIndex = 0
        '
        'cmdsave
        '
        Me.cmdsave.Location = New System.Drawing.Point(36, 277)
        Me.cmdsave.Name = "cmdsave"
        Me.cmdsave.Size = New System.Drawing.Size(85, 23)
        Me.cmdsave.TabIndex = 1
        Me.cmdsave.Text = "&Αποθήκευση"
        Me.cmdsave.UseVisualStyleBackColor = True
        '
        'cmdcancel
        '
        Me.cmdcancel.Location = New System.Drawing.Point(127, 277)
        Me.cmdcancel.Name = "cmdcancel"
        Me.cmdcancel.Size = New System.Drawing.Size(84, 23)
        Me.cmdcancel.TabIndex = 2
        Me.cmdcancel.Text = "Α&κύρωση"
        Me.cmdcancel.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 12)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(73, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Category ID :"
        '
        'txtDie
        '
        Me.txtDie.Location = New System.Drawing.Point(125, 141)
        Me.txtDie.MaxLength = 50
        Me.txtDie.Name = "txtDie"
        Me.txtDie.Size = New System.Drawing.Size(293, 21)
        Me.txtDie.TabIndex = 12
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(11, 141)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(64, 13)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Διεύθυνση :"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(11, 115)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(128, 13)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Τηλέφωνο Ξενοδοχείου :"
        '
        'N1
        '
        Me.N1.AutoSize = True
        Me.N1.Location = New System.Drawing.Point(363, 225)
        Me.N1.Name = "N1"
        Me.N1.Size = New System.Drawing.Size(38, 13)
        Me.N1.TabIndex = 16
        Me.N1.Text = "Label7"
        Me.N1.Visible = False
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(126, 58)
        Me.txtName.MaxLength = 50
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(293, 21)
        Me.txtName.TabIndex = 17
        '
        'txtTHL
        '
        Me.txtTHL.Location = New System.Drawing.Point(125, 112)
        Me.txtTHL.MaxLength = 50
        Me.txtTHL.Name = "txtTHL"
        Me.txtTHL.Size = New System.Drawing.Size(293, 21)
        Me.txtTHL.TabIndex = 18
        '
        'RANK
        '
        Me.RANK.Location = New System.Drawing.Point(125, 32)
        Me.RANK.MaxLength = 50
        Me.RANK.Name = "RANK"
        Me.RANK.Size = New System.Drawing.Size(293, 21)
        Me.RANK.TabIndex = 20
        '
        'ComboCAT
        '
        Me.ComboCAT.FormattingEnabled = True
        Me.ComboCAT.Items.AddRange(New Object() {"2 ΑΣΤΕΡΩΝ", "3 ΑΣΤΕΡΩΝ", "4 ΑΣΤΕΡΩΝ", "5 ΑΣΤΕΡΩΝ"})
        Me.ComboCAT.Location = New System.Drawing.Point(125, 9)
        Me.ComboCAT.Name = "ComboCAT"
        Me.ComboCAT.Size = New System.Drawing.Size(121, 21)
        Me.ComboCAT.TabIndex = 21
        '
        'FrmADDSUPPLIER_ITEM
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ClientSize = New System.Drawing.Size(456, 334)
        Me.Controls.Add(Me.ComboCAT)
        Me.Controls.Add(Me.RANK)
        Me.Controls.Add(Me.txtTHL)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.N1)
        Me.Controls.Add(Me.txtDie)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cmdcancel)
        Me.Controls.Add(Me.cmdsave)
        Me.Controls.Add(Me.txtEmail)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "FrmADDSUPPLIER_ITEM"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Ξενοδοχεία"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents cmdsave As System.Windows.Forms.Button
    Friend WithEvents cmdcancel As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtDie As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents N1 As System.Windows.Forms.Label
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents txtTHL As System.Windows.Forms.TextBox
    Friend WithEvents RANK As System.Windows.Forms.TextBox
    Friend WithEvents ComboCAT As System.Windows.Forms.ComboBox
End Class
