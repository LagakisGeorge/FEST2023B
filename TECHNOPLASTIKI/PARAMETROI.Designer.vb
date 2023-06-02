<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PARAMETROI
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
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.C1EMAIL = New System.Windows.Forms.TextBox
        Me.C2PWD = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.C3HOST = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(426, 385)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(118, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Καταχώρηση"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Location = New System.Drawing.Point(343, 44)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(200, 20)
        Me.DateTimePicker1.TabIndex = 1
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.Location = New System.Drawing.Point(343, 97)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(200, 20)
        Me.DateTimePicker2.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.Label1.Location = New System.Drawing.Point(60, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(246, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Πρώτη ημέρα παραμονής σε Ξενοδοχεία"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.Label2.Location = New System.Drawing.Point(59, 101)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(247, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Τελική ημέρα παραμονής σε Ξενοδοχεία"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.Label3.Location = New System.Drawing.Point(59, 150)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(198, 16)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "email που θα στέλνει (όχι gmail)"
        '
        'C1EMAIL
        '
        Me.C1EMAIL.Location = New System.Drawing.Point(344, 149)
        Me.C1EMAIL.Name = "C1EMAIL"
        Me.C1EMAIL.Size = New System.Drawing.Size(199, 20)
        Me.C1EMAIL.TabIndex = 6
        '
        'C2PWD
        '
        Me.C2PWD.Location = New System.Drawing.Point(344, 196)
        Me.C2PWD.Name = "C2PWD"
        Me.C2PWD.Size = New System.Drawing.Size(200, 20)
        Me.C2PWD.TabIndex = 8
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.Label4.Location = New System.Drawing.Point(59, 200)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(196, 16)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "password email που θα στέλνει"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(161, Byte))
        Me.Label5.Location = New System.Drawing.Point(61, 241)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(250, 16)
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Smtp_Server.Host π.χ. mailgate.otenet.gr"
        '
        'C3HOST
        '
        Me.C3HOST.Location = New System.Drawing.Point(343, 237)
        Me.C3HOST.Name = "C3HOST"
        Me.C3HOST.Size = New System.Drawing.Size(200, 20)
        Me.C3HOST.TabIndex = 10
        '
        'PARAMETROI
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1006, 574)
        Me.Controls.Add(Me.C3HOST)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.C2PWD)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.C1EMAIL)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DateTimePicker2)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.Button1)
        Me.Name = "PARAMETROI"
        Me.Text = "Form3"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePicker2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents C1EMAIL As System.Windows.Forms.TextBox
    Friend WithEvents C2PWD As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents C3HOST As System.Windows.Forms.TextBox
End Class
