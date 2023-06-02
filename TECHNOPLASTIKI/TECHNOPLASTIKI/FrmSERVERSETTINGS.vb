Imports System.IO
Imports System.Text

'Private Sub Button1_Click(ByVal sender As System.Object, _
'ByVal e As System.EventArgs) Handles Button1.Click
'    Try
'        Dim wFile As System.IO.FileStream
'        Dim byteData() As Byte
'        byteData = Encoding.ASCII.GetBytes("FileStream Test1")
'        wFile = New FileStream("streamtest.txt", FileMode.Append)
'        wFile.Write(byteData, 0, byteData.Length)
'        wFile.Close()
'    Catch ex As IOException
'        MsgBox(ex.ToString)
'    End Try
'End Sub
Public Class FrmSERVERSETTINGS
    Dim CONSTR As String
    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Call writeFileStrData(txtip.Text & ":" & txtservername.Text & ":" & txtusername.Text & ":" & txtpassword.Text & ":" & (IIf(RadioButton1.Checked, 1, 2)), Application.StartupPath & "\Config.ini", , "Unicode")
        checkServer()
        Me.Close()
        'MDIMain.Show()
        ' openedFileStream
        'FrmLOGIN.Show()

    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
        'End
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        ' FrmSet.ShowDialog()
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        txtip.Enabled = False
        txtpassword.Enabled = False
        txtusername.Enabled = False
        txtservername.Focus()
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        txtip.Enabled = True
        txtpassword.Enabled = True
        txtservername.Enabled = True
        txtusername.Enabled = True
        txtip.Focus()
    End Sub
    Public Sub writeFileStrData(ByVal MyData As Object, ByVal filePath As String, Optional ByVal transType As String = "", Optional ByVal dataEncoding As String = "")

        Dim Str As String
        Dim fs As FileStream
        Dim tempBytes() As Byte

        tempBytes = Nothing

        If transType = "" Then
            transType = "Append" 'Set default 
        End If

        If dataEncoding = "" Then
            dataEncoding = "ANSI"
        End If

        Try
            Str = CType(MyData, String)
            'Str = CType(Split(MyData, "-")(0) & Chr(10) & Chr(13) & Split(MyData, "-")(1), String)
            'MsgBox(Str)
            If dataEncoding = "ANSI" Then
                tempBytes = System.Text.Encoding.Default.GetBytes(Str)
            ElseIf dataEncoding = "Unicode" Then
                tempBytes = System.Text.Encoding.Unicode.GetBytes(Str)
            End If

            fs = New FileStream(filePath, FileMode.Create, FileAccess.Write)
            If transType = "Append" Then
                fs.Seek(0, SeekOrigin.End)
            ElseIf transType = "Overwrite" Then
                fs.Seek(0, SeekOrigin.Begin)
            End If

            fs.Write(tempBytes, 0, tempBytes.Length)
            fs.Close()
        Catch ex As Exception

            MsgBox(ex.Message & vbCrLf & ex.StackTrace)
        End Try

    End Sub
End Class