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
        Me.txtid = New System.Windows.Forms.TextBox
        Me.txtname = New System.Windows.Forms.TextBox
        Me.txtprice = New System.Windows.Forms.TextBox
        Me.cmdsave = New System.Windows.Forms.Button
        Me.cmdcancel = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtcategory = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(27, 35)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Item ID :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 59)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(66, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Item Name :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(15, 100)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(62, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Item Price :"
        '
        'txtid
        '
        Me.txtid.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtid.Location = New System.Drawing.Point(74, 32)
        Me.txtid.Name = "txtid"
        Me.txtid.ReadOnly = True
        Me.txtid.Size = New System.Drawing.Size(63, 21)
        Me.txtid.TabIndex = 8
        '
        'txtname
        '
        Me.txtname.BackColor = System.Drawing.Color.White
        Me.txtname.Location = New System.Drawing.Point(75, 54)
        Me.txtname.MaxLength = 50
        Me.txtname.Multiline = True
        Me.txtname.Name = "txtname"
        Me.txtname.ReadOnly = True
        Me.txtname.Size = New System.Drawing.Size(171, 43)
        Me.txtname.TabIndex = 9
        '
        'txtprice
        '
        Me.txtprice.Location = New System.Drawing.Point(75, 98)
        Me.txtprice.Name = "txtprice"
        Me.txtprice.Size = New System.Drawing.Size(99, 21)
        Me.txtprice.TabIndex = 0
        '
        'cmdsave
        '
        Me.cmdsave.Location = New System.Drawing.Point(3, 125)
        Me.cmdsave.Name = "cmdsave"
        Me.cmdsave.Size = New System.Drawing.Size(66, 23)
        Me.cmdsave.TabIndex = 1
        Me.cmdsave.Text = "&Save"
        Me.cmdsave.UseVisualStyleBackColor = True
        '
        'cmdcancel
        '
        Me.cmdcancel.Location = New System.Drawing.Point(75, 125)
        Me.cmdcancel.Name = "cmdcancel"
        Me.cmdcancel.Size = New System.Drawing.Size(66, 23)
        Me.cmdcancel.TabIndex = 2
        Me.cmdcancel.Text = "&Cancel"
        Me.cmdcancel.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(5, 12)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(73, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Category ID :"
        '
        'txtcategory
        '
        Me.txtcategory.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.txtcategory.Location = New System.Drawing.Point(74, 9)
        Me.txtcategory.Name = "txtcategory"
        Me.txtcategory.ReadOnly = True
        Me.txtcategory.Size = New System.Drawing.Size(60, 21)
        Me.txtcategory.TabIndex = 11
        '
        'FrmADDSUPPLIER_ITEM
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.WhiteSmoke
        Me.ClientSize = New System.Drawing.Size(284, 160)
        Me.Controls.Add(Me.txtcategory)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cmdcancel)
        Me.Controls.Add(Me.cmdsave)
        Me.Controls.Add(Me.txtprice)
        Me.Controls.Add(Me.txtname)
        Me.Controls.Add(Me.txtid)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "FrmADDSUPPLIER_ITEM"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Supplier Item"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtid As System.Windows.Forms.TextBox
    Friend WithEvents txtname As System.Windows.Forms.TextBox
    Friend WithEvents txtprice As System.Windows.Forms.TextBox
    Friend WithEvents cmdsave As System.Windows.Forms.Button
    Friend WithEvents cmdcancel As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtcategory As System.Windows.Forms.TextBox
End Class
