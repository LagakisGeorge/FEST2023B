<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BARCODE
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
        Me.Button1 = New System.Windows.Forms.Button
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.EIDOS = New System.Windows.Forms.TextBox
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.KATHG = New System.Windows.Forms.ComboBox
        Me.Button5 = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(29, 62)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(165, 27)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "ΔΗΜΙΟΥΡΓΙΑ ΕΤΙΚΕΤΑΣ"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(264, 62)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(344, 225)
        Me.ListBox1.TabIndex = 1
        '
        'EIDOS
        '
        Me.EIDOS.Location = New System.Drawing.Point(266, 30)
        Me.EIDOS.Name = "EIDOS"
        Me.EIDOS.Size = New System.Drawing.Size(138, 20)
        Me.EIDOS.TabIndex = 2
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(415, 29)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(193, 20)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "Αναζήτηση"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(17, 312)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(224, 46)
        Me.Button3.TabIndex = 4
        Me.Button3.Text = "ΑΡΧΕΙΟ ΑΠΟΘΗΚΗΣ"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(264, 312)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(224, 46)
        Me.Button4.TabIndex = 5
        Me.Button4.Text = "ΑΡΧΕΙΟ ΣΥΝΤΑΓΩΝ"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'KATHG
        '
        Me.KATHG.FormattingEnabled = True
        Me.KATHG.Location = New System.Drawing.Point(17, 287)
        Me.KATHG.Name = "KATHG"
        Me.KATHG.Size = New System.Drawing.Size(223, 21)
        Me.KATHG.TabIndex = 6
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(17, 384)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(224, 46)
        Me.Button5.TabIndex = 7
        Me.Button5.Text = "ΤΙΜΟΛΟΓΙΑ Α' ΥΛΩΝ"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'BARCODE
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Cyan
        Me.ClientSize = New System.Drawing.Size(1165, 689)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.KATHG)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.EIDOS)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "BARCODE"
        Me.Text = "BARCODE"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents EIDOS As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents KATHG As System.Windows.Forms.ComboBox
    Friend WithEvents Button5 As System.Windows.Forms.Button
End Class
