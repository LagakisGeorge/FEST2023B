Public Class PARAMETROI

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ff As String = "MM/dd/yyyy HH:mm"
        Dim ci1 As String = Format(DateTimePicker1.Value, ff)
        Dim ci2 As String = Format(DateTimePicker2.Value, ff)



        'Smtp_Server.Credentials = New Net.NetworkCredential("lagakis@otenet.gr", "a8417!")
        'Smtp_Server.Port = 587
        'Smtp_Server.EnableSsl = True
        'Smtp_Server.Host = "mailgate.otenet.gr"


        ExecuteSQLQuery("update MEM SET HMEARX='" + ci1 + "',HMETEL='" + ci2 + "' WHERE ID=1 ")
        ExecuteSQLQuery("update MEM SET C1='" + C1EMAIL.Text + "' WHERE ID=1 ")
        ExecuteSQLQuery("update MEM SET C2='" + C2PWD.Text + "' WHERE ID=1 ")
        ExecuteSQLQuery("update MEM SET C3='" + C3HOST.Text + "' WHERE ID=1 ")

        Dim DT2 As New DataTable
        DT2 = Execute2SQLQuery("SELECT ISNULL(HMEARX,GETDATE()) AS HMEARX,ISNULL(HMETEL,GETDATE()) AS HMETEL,ISNULL(C1,'') AS C1,ISNULL(C2,'') AS C2,ISNULL(C3,'') AS C3 FROM MEM WHERE ID=1")
        DateTimePicker1.Value = DT2.Rows(0)("HMEARX")
        DateTimePicker2.Value = DT2.Rows(0)("HMETEL")
        C1EMAIL.Text = DT2.Rows(0)("C1")
        C2PWD.Text = DT2.Rows(0)("C2")
        C3HOST.Text = DT2.Rows(0)("C3")

        gHMEARX = DT2.Rows(0)("HMEARX")
        gHMETEL = DT2.Rows(0)("HMETEL")
        gC1EMAIL = DT2.Rows(0)("C1")
        gC2PWD = DT2.Rows(0)("C2")
        gC3HOST = DT2.Rows(0)("C3")

        

        







    End Sub

    Private Sub PARAMETROI_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim DT2 As New DataTable
        DT2 = Execute2SQLQuery("SELECT ISNULL(HMEARX,GETDATE()) AS HMEARX,ISNULL(HMETEL,GETDATE()) AS HMETEL,ISNULL(C1,'') AS C1,ISNULL(C2,'') AS C2,ISNULL(C3,'') AS C3 FROM MEM WHERE ID=1")
        DateTimePicker1.Value = DT2.Rows(0)("HMEARX")
        DateTimePicker2.Value = DT2.Rows(0)("HMETEL")
        C1EMAIL.Text = DT2.Rows(0)("C1")
        C2PWD.Text = DT2.Rows(0)("C2")
        C3HOST.Text = DT2.Rows(0)("C3")

        gHMEARX = DT2.Rows(0)("HMEARX")
        gHMETEL = DT2.Rows(0)("HMETEL")
        gC1EMAIL = DT2.Rows(0)("C1")
        gC2PWD = DT2.Rows(0)("C2")
        gC3HOST = DT2.Rows(0)("C3")






    End Sub
End Class