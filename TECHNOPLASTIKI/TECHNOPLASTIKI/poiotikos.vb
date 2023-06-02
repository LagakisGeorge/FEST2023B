
Imports CrystalDecisions.CrystalReports.Engine

Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.OleDb
Imports System.Net.NetworkInformation
Imports System.Transactions


Public Class VARDIA

    Private Sub VARDIA_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FILLComboBox("select EPO,space(300)+STR(ID) FROM VARDIA order by EPO ", VARDIES)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        ' Exit Sub
        Dim TELBARCODE As String
        TELBARCODE = ""  '"'0205200000" + Mid(kodPROION + Space(6), 1, 6) + "300" + Mid(Str(100000 + minTem), 3, 5) + "10000" + Mid(LTrim(Str(MP)), 1, 7)



        '"UPDATE MEM SET MEMO='" + onomaProion + "',
        'TPSKETO=" + Str(minTem) + ",HME=GETDATE(),
        'PAR2='" + kodPROION + "',PAR1='" + TELBARCODE + "'")
        Dim R As New DataTable

        'Dim R As New DataTable
        ExecuteSQLQuery("SELECT COUNT(*) AS N FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_NAME  = 'POIOTIKOS'", R)    'On Error Resume Next

        If R(0)(0) = 0 Then
            ExecuteSQLQuery("CREATE TABLE POIOTIKOS ( KOD VARCHAR(20) NOT NULL,HME DATETIME NOT NULL,ID int identity(1,1) NOT NULL )", R)
        End If









        ExecuteSQLQuery("INSERT INTO POIOTIKOS ( KOD ,HME  ) VALUES ('" + Split(PROION.Text, "*")(0) + "',getdate() )", R)





        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open("C:\mercvb\poiot-label.xlsx")
        xlWorkSheet = xlWorkBook.Worksheets("sh1")
        'display the cells value B2
        '    MsgBox(xlWorkSheet.Cells(6, 1).value)
        'edit the cell with new value

        xlWorkSheet.Cells(16, 2) = Mid(PROION.Text, 8, 50)
        xlWorkSheet.Cells(16, 1) = Mid(PROION.Text, 1, 7)
        xlWorkSheet.Cells(13, 3) = VARDIES.Text

        'xlWorkSheet.Cells(8, 1) = "ΠΟΣΟΤΗΤΑ: " + "" 'Str(minTem)
        'xlWorkSheet.Cells(9, 1) = "ΠΑΡΤΙΔΑ/LOT :" + "" ' Str(MP)
        'xlWorkSheet.Cells(15, 1) = TELBARCODE
        'xlWorkSheet.Cells(18, 2) = PROION.Text ' kodPROION
        'Globals.xlworkSheet.PrintOut(From:=1, To:=1, Copies:=2, Preview:=True)
        xlWorkBook.Save()

        xlWorkSheet.PrintOut(From:=1, To:=1, Copies:=1, Preview:=False)


        xlWorkBook.Save()


        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)




        Me.Close()







    End Sub
End Class