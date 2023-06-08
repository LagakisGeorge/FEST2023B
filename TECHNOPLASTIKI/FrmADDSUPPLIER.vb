Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Interop.Outlook
Imports System.IO

Imports System.Net.Mail
Public Class FrmAddSupplier
    Dim stockID As Integer
    Dim hOldID As Integer
    Dim M_ID As Long = 0
    Dim mIsNew As Boolean = False

    Public Property IsNew() As Integer
        Get
            Return mIsNew
        End Get
        Set(ByVal Value As Integer)
            ' If Value < 1 Or Value > 12 Then
            ' Error processing for invalid value. 
            'Else
            mIsNew = Value
            'End If
        End Set
    End Property

    Public Property ID() As Integer
        Get
            Return M_ID
        End Get
        Set(ByVal Value As Integer)
            ' If Value < 1 Or Value > 12 Then
            ' Error processing for invalid value. 
            'Else
            M_ID = Value
            'End If
        End Set
    End Property




    Private Sub cmdsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdsave.Click
        Dim SQL As String
        Dim mMON As String = Str(Val(onoProsf.Text))

        Dim mb As String = DIE.Text
        mb = Str(Val(mb))
        If Len(email.Text) = 0 Then
            'MsgBox("ΔΕΝ ΒΑΛΑΤΕ email")
            'Exit Sub
        End If
        If Len(ONO.Text) = 0 Then
            MsgBox("ΔΕΝ ΒΑΛΑΤΕ ΕΠΩΝΥΜΙΑ")
            Exit Sub
        End If
        If Len(onoProsf.Text) = 0 Then
            MsgBox("ΔΕΝ ΒΑΛΑΤΕ ΠΡΟΣΦΩΝΗΣΗ")
            Exit Sub
        End If
        Dim cc As String = ""
        For l As Integer = 0 To CheckedListBox1.Items.Count - 1
            If CheckedListBox1.GetItemChecked(l) = True Then
                cc = cc + "1"
            Else
                cc = cc + "0"
            End If
        Next


        Dim cc4 As String = ""
        For l = 0 To CheckedListBox2.Items.Count - 1
            If CheckedListBox2.GetItemChecked(l) = True Then
                cc4 = cc4 + "1"
            Else
                cc4 = cc4 + "0"
            End If
        Next

        Dim mkod As String = email.Text
        Dim mono As String = ONO.Text
        Dim m_mon As String = onoProsf.Text

        Dim mBaros As String = DIE.Text
        Dim ff As String = "MM/dd/yyyy HH:mm"
        Dim ci As String = Format(DTCheckin.Value, ff)
        Dim co As String = Format(DTCheckout.Value, ff)
        Dim aaf As String = Format(DtAirAfixi.Value, ff)
        Dim aan As String = Format(dtAirAnax.Value, ff)

        If DTCheckin.Value > DTCheckout.Value Then
            MsgBox("ΗΜΕΡΟΜΗΝΊΑ CHECKOUT ΜΙΚΡΟΤΕΡΗ ΤΟΥ CHECKIN")
            Exit Sub
        End If

        If DTCheckin.Value < gHMEARX Or DTCheckin.Value > gHMETEL Then
            MsgBox("ΗΜΕΡΟΜΗΝΊΑ CHECKIN EKTOΣ ΔΙΑΣΤΗΜΑΤΟΣ  " + gHMEARX.ToString + "-" + gHMETEL.ToString)
            Exit Sub
        End If
        If DTCheckout.Value < gHMEARX Or DTCheckout.Value > gHMETEL Then
            MsgBox("ΗΜΕΡΟΜΗΝΊΑ CHECKOUT EKTOΣ ΔΙΑΣΤΗΜΑΤΟΣ  " + gHMEARX.ToString + "-" + gHMETEL.ToString)
            Exit Sub
        End If
        If rank.Text.Length = 0 Then rank.Text = "0"


        '  gHMETEL = DT2.Rows(0)("HMETEL")

        Dim idsynodoy As String = "0"
        Dim c3 As String = ""
        If CheckBox1.Checked Then ' einai synodos (δηλαδη δεν ειναι ο κυριωσ προσκεκλημενοσ)
            c3 = "1"
        Else
            c3 = "0"
            'einai o kyrios proskeklhmenos kai exei synodo
            If InStr(SYNODOS.Text, ";") > 0 Then

                idsynodoy = SYNODOS.Text.Split(";")(1)

                'ΕΝΗΜΕΡΩΝΩ ΤΟΝ ΥΙΟ ΜΕ ΤΟ ΙΔ ΤΟΥ ΠΑΤΕΡΑ ΤΟΥ
                ExecuteSQLQuery("UPDATE PEL SET IDPATRIKO=" + Str(ID) + " where ID=" + idsynodoy)
            End If
        End If

        If IsNew Then
            SQL = "insert into PEL (IDSYNODOY,ENERGOS,KINHTO,CH6,AIRPORT,CH5,RANK,SYNODOS,CH4,CH3,CHECKIN,CHECKOUT,AIRAFIXI,AIRANAX,EMAIL,EPO,ONO,DIE) VALUES (" + idsynodoy + "," + c3 + ",'" + KINHTO.Text + "','" + PTHSHANAXC6.Text + "','" + Airport.Text + "','" + PTHSHC5.Text + "'," + rank.Text + ",'" + SYNODOS.Text + "','" + cc4 + "','" + cc + "','" + ci + "','" + co + "','" + aaf + "','" + aan + "','" + email.Text + "','" + Replace(ONO.Text, "'", "`") + "','" + onoProsf.Text + "','" + mBaros + "')"
        Else
            SQL = "UPDATE PEL SET IDSYNODOY=" + idsynodoy + ",ENERGOS=" + c3 + ",AIRAFIXI='" + aaf + "',AIRANAX='" + aan + "',KINHTO='" + KINHTO.Text + "',CH6='" + PTHSHANAXC6.Text + "',AIRPORT='" + cAirport.Text + "',CH5='" + PTHSHC5.Text + "',RANK=" + rank.Text + ",SYNODOS='" + SYNODOS.Text + "',CH4='" + cc4 + "',CH3='" + cc + "',CHECKOUT='" + co + "',CHECKIN='" + ci + "',EMAIL='" + mkod + "',EPO='" + mono + "',ONO='" + m_mon + "',DIE='" + mBaros + "'  WHERE ID=" + Str(ID)
        End If



        Try
            ExecuteSQLQuery(SQL)
        Catch
            MsgBox("ΔΕΝ ΚΑΤΕΧΩΡΗΘΗ " + Err.Description)
        End Try

        Me.Close()

    End Sub


    Private Sub ListBox1_DrawItem(ByVal sender As Object, ByVal e As DrawItemEventArgs)
        ''This code draws a checkbox using the DrawCheckBox method of the ControlPaint class and uses the DrawString method of the Graphics object to draw the text of the item. The if statement inside the DrawCheckBox method sets the state of the checkbox to either ButtonState.Checked or ButtonState.Normal, depending on whether the item is selected.

        '' Draw the background of the ListBox control for each item.
        'e.DrawBackground()

        '' Determine the color of the checkbox based on whether the item is selected.
        'Dim checkboxColor As Color
        'If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
        '    checkboxColor = SystemColors.HighlightText
        'Else
        '    checkboxColor = SystemColors.ControlText
        'End If

        '' Draw the checkbox next to the item text.
        'Dim checkboxRect As New Rectangle(e.Bounds.X, e.Bounds.Y, e.Bounds.Height, e.Bounds.Height)
        'Dim n As Integer
        'If e.Index = 1 Then n = ButtonState.Checked Else n = ButtonState.Normal
        'ControlPaint.DrawCheckBox(e.Graphics, checkboxRect, n) 'If(e.Index Mod 2 = 0, n, ButtonState.Normal)
        'e.Graphics.DrawString(ListBox1.Items(e.Index), Me.Font, New SolidBrush(checkboxColor), e.Bounds.X + checkboxRect.Width, e.Bounds.Y)
    End Sub


    Private Sub ListBox1_Click(ByVal sender As Object, ByVal e As EventArgs)

    End Sub
    'This code uses the IndexFromPoint method of the ListBox to get the index of the clicked item, and the GetItemRectangle method to get the bounds of the item. It then creates a Rectangle object for the checkbox based on the item bounds, and checks whether the click occurred within this rectangle using the Contains method.

    'If the click occurred within the checkbox, the code sets the selected state of the item to True using the SetSelected method, and toggles the checked state of the checkbox using the SetItemChecked method. Note that you need to set the selected state to True in order for the checked state to be updated correctly.












    Private Sub FrmAddSupplier_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DTCheckin.CustomFormat = "dd/MM/yyyy HH:mm"
        DTCheckout.CustomFormat = "dd/MM/yyyy HH:mm"
        'Dim SQL As String
        'Dim mMON As String = Str(Val(AFM.Text))

        'Dim mb As String = DIE.Text
        'mb = Str(Val(mb))
        'If Len(KOD.Text) = 0 Then
        '    MsgBox("ΔΕΝ ΒΑΛΑΤΕ ΚΩΔΙΚΟ")
        '    Exit Sub
        'End If
        'If Len(ONO.Text) = 0 Then
        '    MsgBox("ΔΕΝ ΒΑΛΑΤΕ ΠΕΡΙΓΡΑΦΗ")
        '    Exit Sub
        'End If
        'If Len(AFM.Text) = 0 Then
        '    MsgBox("ΔΕΝ ΒΑΛΑΤΕ ΜΟΝΑΔΑ ΜΕΤΡΗΣΗΣ")
        '    Exit Sub
        'End If

    End Sub

    Private Sub DTCheckin_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTCheckin.ValueChanged

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles send.Click















        Try
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential(gC1EMAIL, gC2PWD)
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = gC3HOST

            e_mail = New MailMessage()
            e_mail.From = New MailAddress(Trim(email.Text))
            If mAttachment.Text.Length > 1 Then
                Dim attachment As System.Net.Mail.Attachment
                attachment = New System.Net.Mail.Attachment(mAttachment.Text)
                e_mail.Attachments.Add(attachment)
            End If


            e_mail.To.Add(email.Text)
            'Dim item As System.Net.Mail.Attachment
            'e_mail.Attachments.Add(item)

            e_mail.Subject = Subject.Text '"Email Sending"
            e_mail.IsBodyHtml = False
            e_mail.Body = txtMessage.Text
            Smtp_Server.Send(e_mail)
            MsgBox("Mail Sent")

        Catch
            MsgBox(Err.Description)
        End Try


        'Try
        '    Dim oMail As New SmtpMail("TryIt")
        '    ' Set sender email address, please change it to yours
        '    oMail.From = "test@emailarchitect.net"
        '    ' Set recipient email address, please change it to yours
        '    oMail.To = "support@emailarchitect.net"

        '    ' Set email subject
        '    oMail.Subject = "test HTML email with attachment"
        '    ' Set HTML body
        '    oMail.HtmlBody = "<font size=5>This is</font> <font color=red><b>a test</b></font>"

        '    ' Add attachment from local disk
        '    oMail.AddAttachment("d:\test.pdf")

        '    ' Add attachment from remote website
        '    oMail.AddAttachment("http://www.emailarchitect.net/webapp/img/logo.jpg")

        '    ' Your SMTP server address
        '    Dim oServer As New SmtpServer("smtp.emailarchitect.net")

        '    ' User and password for ESMTP authentication
        '    oServer.User = "test@emailarchitect.net"
        '    oServer.Password = "testpassword"

        '    ' Most mordern SMTP servers require SSL/TLS connection now.
        '    ' ConnectTryTLS means if server supports SSL/TLS, SSL/TLS will be used automatically.
        '    oServer.ConnectType = SmtpConnectType.ConnectTryTLS

        '    ' If your SMTP server uses 587 port
        '    ' oServer.Port = 587

        '    ' If your SMTP server requires SSL/TLS connection on 25/587/465 port
        '    ' oServer.Port = 25 ' 25 or 587 or 465
        '    ' oServer.ConnectType = SmtpConnectType.ConnectSSLAuto

        '    Console.WriteLine("start to send email with attachment ...")

        '    Dim oSmtp As New SmtpClient()
        '    oSmtp.SendMail(oServer, oMail)

        '    Console.WriteLine("email was sent successfully!")
        'Catch ep As Exception
        '    Console.WriteLine("failed to send email with the following error:")
        '    Console.WriteLine(ep.Message)
        'End Try



    End Sub



    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
        mAttachment.Text = OpenFileDialog1.FileName
    End Sub

    Private Sub cmdcancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdcancel.Click
        Me.Close()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        SendMailDemo()
    End Sub

    Public Sub SendMailDemo()
        Try
            ' Create the Outlook application.
            Dim outlookApp As New Outlook.Application()
            ' Create a new mail item.
            Dim outlookMailItem = DirectCast(outlookApp.CreateItem(OlItemType.olMailItem), MailItem)

            outlookMailItem.HTMLBody = txtMessage.Text   '"Hello, Jawed your message body will go here!!"
            'Add an attachment.
            Dim displayNameForAttachment = mAttachment.Text
            Dim position As Integer = CInt(outlookMailItem.Body.Length) + 1
            Dim attachType = CInt(OlAttachmentType.olByValue)

            Dim wordDocPathName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, mAttachment.Text)
            Dim theAttachment As Outlook.Attachment = outlookMailItem.Attachments.Add(
                wordDocPathName, attachType, position, displayNameForAttachment)

            outlookMailItem.Subject = Subject.Text
            Dim recipients As Recipients = outlookMailItem.Recipients
            ' Change the recipient in the next line if necessary.
            Dim oRecip = recipients.Add(Trim(email.Text))
            oRecip.Resolve()
            outlookMailItem.Send()
            ' Clean up.
            oRecip = Nothing
            recipients = Nothing
            outlookMailItem = Nothing
            outlookApp = Nothing
            'end of try block
        Catch ex As System.Exception
            Console.WriteLine("Put breakpoint here")
        End Try
    End Sub

End Class