Imports System.Data.OleDb
Imports System.Net.NetworkInformation

Public Class BARCODE

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        report.Show()




    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ' ExecuteSQLQuery("SELECT * FROM EID WHERE ONO LIKE '" + EIDOS.Text + "%'")
        FillListBox("SELECT ONO,KOD AS CID  FROM YLIKA WHERE N1=4 AND ONO LIKE '" + EIDOS.Text + "%'", ListBox1)
    End Sub

    Private Sub BARCODE_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        checkServer()
        ExecuteSQLQuery("SELECT * FROM PINAKES WHERE TYPOS=11 ORDER BY AYJON")

        For K As Integer = 0 To sqlDT.Rows.Count - 1
            KATHG.Items.Add(sqlDT(K)("AYJON").ToString + ";" + sqlDT(K)("PERIGRAFH").ToString)

        Next

        KATHG.Text = KATHG.Items(0).ToString

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim N As Long
        N = ListBox1.SelectedIndex
        Dim C As String = Split(ListBox1.Items(N).ToString, ";")(1)
        ExecuteSQLQuery("UPDATE MEM SET MEMO='" + Split(ListBox1.Items(N).ToString, ";")(0) + "' WHERE ID=1")
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim mergates As New ergates()
        'Set the Parent Form of the Child window.
        mergates.MdiParent = MDIMain
        'Display the new form.





        For k = 0 To 20
            mergates.widths(7) = 100
        Next
        '  ExecuteSQLQuery("update YLIKA SET N1=(SELECT TOP 1 AEG FROM EID WHERE KOD=YLIKA.KOD) ")

        Dim Mn1 As String
        Mn1 = Split(KATHG.Text, ";")(0)


        mergates.Text = "Αρχείο Υλικών"
        mergates.Label1.Text = "SELECT ONO AS [Ονομα ],KOD AS [ΚΩΔ],N1 AS [ΚΑΤΗΓ],BAROS AS [ΒΑΡΟΣ],C1,C2  FROM YLIKA WHERE N1=" + Mn1 + " ORDER BY KOD "

        ' ergates.MdiParent = Me
        mergates.WindowState = FormWindowState.Maximized
        mergates.STHLHONOMATOS_ID = 0
        mergates.STHLHTOY_ID = 1
        mergates.widths(1) = 100
        mergates.QUERY_AFTER = "update YLIKA SET N1=" + Mn1 + " WHERE N1 IS NULL"
        For KK As Integer = 0 To 6
            mergates.widths(KK) = 100
        Next
        mergates.Label2.Text = KATHG.Text
        mergates.widths(0) = 400
        gMenu = 22
        mergates.Show()


        Exit Sub

    End Sub

    'End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'Form1.Show()
        For k = 0 To 20
            ergates.widths(7) = 100
        Next
        '  ExecuteSQLQuery("update YLIKA SET N1=(SELECT TOP 1 AEG FROM EID WHERE KOD=YLIKA.KOD) ")

        Dim Mn1 As String
        'Mn1 = Split(KATHG.Text, ";")(0)


        ergates.Text = "Αρχείο Συνταγών"
        ergates.Label1.Text = "SELECT KOD AS [ΚΩΔ],KODSYNOD AS [ΣΥΣΤΑΤΙΚΑ],POSOSTO AS [ΠΟΣΟΣΤΟ],ID  FROM SYNTAGES  ORDER BY KOD,KODSYNOD "

        ' ergates.MdiParent = Me
        ergates.WindowState = FormWindowState.Maximized
        ergates.STHLHONOMATOS_ID = 0
        ergates.STHLHTOY_ID = 3
        ergates.widths(1) = 100
        ergates.QUERY_AFTER = ""   ' update YLIKA SET N1=" + Mn1 + " WHERE N1 IS NULL"
        For KK As Integer = 0 To 6
            ergates.widths(KK) = 100
        Next
        ergates.Label2.Text = KATHG.Text

        gMenu = 22
        ergates.Show()


        Exit Sub





    End Sub

    Private Sub EIDOS_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EIDOS.TextChanged

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        TIMOLOGIApol.Show()

    End Sub
End Class