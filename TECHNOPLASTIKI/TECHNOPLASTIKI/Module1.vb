'Created on August 16, 2010
'Tan, Angelito S.

'Date update dec 12, 2010
Imports System.Data.OleDb
Imports System.Net.NetworkInformation
Imports Excel = Microsoft.Office.Interop.Excel
Module ModCon
    'Public fso As New filesystemobject
    'Public ParamDVFrom As New CrystalDecisions.Shared.ParameterDiscreteValue
    'Public ParamDVTo As New CrystalDecisions.Shared.ParameterDiscreteValue
    'Public ParamCompanyName As New CrystalDecisions.Shared.ParameterDiscreteValue
    'Public ParamCompanyLoc As New CrystalDecisions.Shared.ParameterDiscreteValue
    'Public ParamCompanyContact As New CrystalDecisions.Shared.ParameterDiscreteValue
    'Public ParamCompanyTIN As New CrystalDecisions.Shared.ParameterDiscreteValue
    'Public _USER As New CrystalDecisions.Shared.ParameterDiscreteValue
    'Public mReport As New CrystalDecisions.CrystalReports.Engine.ReportDocument

    Declare Function SetDefaultPrinter Lib "winspool.drv" Alias "SetDefaultPrinterA" (ByVal pszPrinter As String) As Boolean
    Declare Function GetDefaultPrinter Lib "winspool.drv" Alias "GetDefaultPrinterA" (ByVal pszBuffer() As String, ByVal pcchBuffer As Integer) As Boolean




    Public sqlDT As New DataTable

    Public sqlDaTaSet As New DataSet
    Public sqlDTx As New DataTable
    Public openedFileStream As System.IO.Stream
    Dim xsize As Integer
    'Public Const cnString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=../database/SaleInv_DB.mdb"
    'Public Const cnString As String = "Provider=SQLNCLI10;Server=CPAT;Database=SaleInv_DB; Trusted_Connection=yes;"
    'Public Const cnString As String = "Provider=SQLNCLI10;Server=CPAT;Database=SaleInv_DB;Uid=sa; Pwd=angelito;"

    'Public Const cnstring As String = "Provider=SQLOLEDB;" & _
    '                                  "Data Source=;" & _
    '                                  "Network=CPAT;" & _
    '                                  "Initial Catalog=SaleInv_DB;" & _
    '                                  "User Id=sa;" & _
    '                                  "Password=angelito"
    '192.168.1.104;" & _'                           

    Public gConnect As String
    Public gConSQL As String
    Public gMenu As Integer

    'Public GDB As New ADODB.Connection

    Public gFPA As Single
    ' Public gMenu As Integer


    'Public conn As OleDbConnection = New OleDbConnection(cnString)
    ' Public DataFileLock As New System.Threading.ReaderWriterLock
    Public gPHDHMAGENNA As Integer
    Public gPHDHMAnextOXEIA As Integer
    Public gAPOGnextOXEIA As Integer





    Public sqlSTR As String
    Public Rpt_SqlStr As String
    Public pass As Boolean
    Public VAT As Double
    Public username As String
    Public xUser_ID As Integer
    Public xUser_Access As String
    Public Pending_ID As Integer
    Public Pending_QTY As Integer
    Public Pending_Item_ID As Integer
    Public dataBytes() As Byte
    Public xpass As Boolean
    Public howx As Integer
    Public xid(1) As Integer
    Public xlock As Boolean
    Public iMin As Integer
    Public tmpStr As String
    Public LOGID As Integer
    Public PreviousPage, NextPage As Integer
    Public i_Print As Integer
    Public gMHNAS As String

    Public Function toTeleia(ByVal c As String) As String
        toTeleia = Replace(c, ",", ".")
    End Function
    Public Function cNull(ByVal c As Object) As String
        If IsDBNull(c) Then
            cNull = ""
        Else
            cNull = c
        End If


    End Function


    Public Function checkServer() As Boolean
        Dim c As String
        '  gMHNAS = Format(Now, "yyyymm")
        '  gMHNAS = InputBox("ΜΗΝΑΣ ΕΡΓΑΣΙΑΣ ", "", gMHNAS)

        c = Application.StartupPath & "\Config.ini"

        Try

            With FrmSERVERSETTINGS
                .OpenFileDialog1.FileName = c
                openedFileStream = .OpenFileDialog1.OpenFile()
            End With

            ReDim dataBytes(openedFileStream.Length - 1) 'Init 
            openedFileStream.Read(dataBytes, 0, openedFileStream.Length)
            openedFileStream.Close()
            tmpStr = System.Text.Encoding.Unicode.GetString(dataBytes)

            With FrmSERVERSETTINGS
                If Split(tmpStr, ":")(4) = "1" Then
                    'network
                    gConnect = "Provider=SQLOLEDB.1;Persist Security Info=True;" & _
                               "Data Source=" & Split(tmpStr, ":")(1) & _
                               ";Initial Catalog=" & Trim(Split(tmpStr, ":")(5)) & _
                               ";User Id=" & Split(tmpStr, ":")(2) & _
                               ";Password=" & Split(tmpStr, ":")(3)
                    gConSQL = "Server=" & Split(tmpStr, ":")(1) & ";Database=" & Split(tmpStr, ":")(5) & ";User Id=" & Split(tmpStr, ":")(2) & ";Password=" & Split(tmpStr, ":")(3)
                    'Server=myServerName\myInstanceName;Database=myDataBase;User Id=myUsername;
                    'Password=myPassword;
                Else
                    'local
                    'MsgBox(Split(tmpStr, ":")(1))
                    gConnect = "Provider=SQLOLEDB;Server=" & Split(tmpStr, ":")(1) & _
                               ";Database=MERCURY; Trusted_Connection=yes;"

                    gConSQL = "Data Source=" & Split(tmpStr, ":")(1) & ";Integrated Security=True;database=MERCURY"
                    'cnString = "Data Source=localhost\SQLEXPRESS;Integrated Security=True;database=YGEIA"


                    '
                End If
            End With
            Dim sqlCon As New OleDbConnection
            sqlCon.ConnectionString = gConnect
            sqlCon.Open()
            checkServer = True
            sqlCon.Close()


            ADD_FIELD("MEM", "GS1", "BIGINT")
            ADD_FIELD("YLIKA", "PIECESPERPALLET", "BIGINT")

            ADD_FIELD("PALETES", "DATE", "DATETIME")



            '            Dim GDB As New ADODB.Connection
            'GDB.Open(gConnect)
        Catch ex As Exception
            checkServer = False
            MsgBox("εξοδος λογω μη σύνδεσης με βάση δεδομένων. Ελέγξτε το config.ini")
            End
        End Try
    End Function

    Public Function Execute2SQLQuery(ByVal SQLQuery As String) As System.Data.DataTable
        Try
            Dim sqlCon As New OleDbConnection(gConnect)

            Dim sqlDA As New OleDbDataAdapter(SQLQuery, sqlCon)

            Dim sqlCB As New OleDbCommandBuilder(sqlDA)
            sqlDT.Reset() ' refresh 
            sqlDA.Fill(sqlDT)
            'rowsAffected = command.ExecuteNonQuery();
            ' sqlDA.Fill(sqlDaTaSet, "PEL")

        Catch ex As Exception
            MsgBox("Error: " & ex.ToString)
            If Err.Number = 5 Then
                MsgBox("Invalid Database, Configure TCP/IP", MsgBoxStyle.Exclamation, "Sales and Inventory")
                End

            Else
                MsgBox("Error : " & ex.Message)
            End If
            MsgBox("Error No. " & Err.Number & " Invalid database or no database found !! Adjust settings first", MsgBoxStyle.Critical, "Sales And Inventory")
            MsgBox(SQLQuery)
        End Try
        Return sqlDT
    End Function

    Function Kau_Aritmoy(ByVal mposo As String) As String
        'καθαρισμα αριθμου
        'Dim mPoso As String = xlWorkSheet.Cells(N, 9).VALUE.ToString

        If InStr(mPoso, ",") > InStr(mPoso, ".") Then
            If InStr(mPoso, ".") = 0 Then  ' 1234,50
                mPoso = Replace(mPoso, ",", ".")
            Else  '12.234,67
                mPoso = Replace(mPoso, ".", "")
                mPoso = Replace(mPoso, ",", ".")
            End If
        Else  ' 12,350.56
            If InStr(mPoso, ",") = 0 Then  ' 1234.50
                ' ok
            Else  '12,234.67
                mPoso = Replace(mPoso, ",", "")

            End If



        End If
        Kau_Aritmoy = mposo

    End Function


    Public Function ExecuteSQLQuery(ByVal SQLQuery As String) As DataTable
        Try
            Dim sqlCon As New OleDbConnection(gConnect)

            Dim sqlDA As New OleDbDataAdapter(SQLQuery, sqlCon)

            Dim sqlCB As New OleDbCommandBuilder(sqlDA)
            sqlDT.Reset() ' refresh 
            sqlDA.Fill(sqlDT)
            'Dim rowsAffected As Integer = sqlDT.Rows.Count
            ' sqlDA.Fill(sqlDaTaSet, "PEL")

        Catch ex As Exception
            MsgBox("Error: " & ex.ToString)
            If Err.Number = 5 Then
                MsgBox("Invalid Database, Configure TCP/IP", MsgBoxStyle.Exclamation, "Sales and Inventory")
                ' End

            Else
                MsgBox("Error : " & ex.Message)
            End If
            MsgBox("Error No. " & Err.Number & " Invalid database or no database found !! Adjust settings first", MsgBoxStyle.Critical, "Sales And Inventory")
            MsgBox(SQLQuery)
        End Try
        Return sqlDT
    End Function
    Public Sub ExecuteSQLQuery(ByVal SQLQuery As String, ByRef SQLDT As DataTable)
        'αν χρησιμοποιώ  byref  tote prepei να δηλωθεί   
        'Dim DTI As New DataTable


        Try
            Dim sqlCon As New OleDbConnection(gConnect)

            Dim sqlDA As New OleDbDataAdapter(SQLQuery, sqlCon)

            Dim sqlCB As New OleDbCommandBuilder(sqlDA)
            SQLDT.Reset() ' refresh 
            sqlDA.Fill(SQLDT)
            ' sqlDA.Fill(sqlDaTaSet, "PEL")

        Catch 'ex As Exception
            'MsgBox("Error: " & ex.ToString)
            'If Err.Number = 5 Then
            '    MsgBox("Invalid Database, Configure TCP/IP", MsgBoxStyle.Exclamation, "Sales and Inventory")
            'Else
            '    MsgBox("Error : " & ex.Message)
            'End If
            'MsgBox("Error No. " & Err.Number & " Invalid database or no database found !! Adjust settings first", MsgBoxStyle.Critical, "Sales And Inventory")
            MsgBox(Err.Description + Chr(13) + SQLQuery)
        End Try
        'Return sqlDT
    End Sub
    Public Sub FILLComboBox(ByVal sql As String, ByVal cb As ComboBox)
        Dim conn As OleDbConnection = New OleDbConnection(gConnect)
        cb.Items.Clear()
        Try
            conn.Open()
            Dim cmd As OleDbCommand = New OleDbCommand(sql, conn)
            Dim rdr As OleDbDataReader = cmd.ExecuteReader
            While rdr.Read
                cb.Items.Add(rdr(0).ToString & " - " & rdr(1).ToString)
            End While
            rdr.Close()
        Catch ex As Exception
            MsgBox("Error:" & ex.ToString)
        Finally
            conn.Close()
        End Try
    End Sub
    Public Sub FillListBox(ByVal sql As String, ByVal cb As ListBox)
        Dim conn As OleDbConnection = New OleDbConnection(gConnect)
        cb.Items.Clear()
        Try
            conn.Open()
            Dim cmd As OleDbCommand = New OleDbCommand(sql, conn)
            Dim rdr As OleDbDataReader = cmd.ExecuteReader
            While rdr.Read
                cb.Items.Add(rdr(0).ToString & " ; " & rdr(1).ToString)
            End While
            rdr.Close()
        Catch ex As Exception
            MsgBox("Error:" & ex.ToString)
        Finally
            conn.Close()
        End Try
    End Sub

    Public Sub FILLComboBox2(ByVal sql As String, ByVal cb As ComboBox, ByRef ids As Array)
        '   FILLComboBox2("select NAME,ID FROM ERGATES", ergates, iderg)  π πίνακας (IDERG) παίρνει το  (ID)
        Dim conn As OleDbConnection = New OleDbConnection(gConnect)
        cb.Items.Clear()
        Try
            conn.Open()
            Dim cmd As OleDbCommand = New OleDbCommand(sql, conn)
            Dim rdr As OleDbDataReader = cmd.ExecuteReader
            Dim n As Integer = 0
            While rdr.Read
                cb.Items.Add(rdr(0).ToString)
                ids(n) = rdr(1)
                n = n + 1
            End While
            rdr.Close()

        Catch ex As Exception
            MsgBox("Error:" & ex.ToString)
        Finally
            conn.Close()
        End Try
    End Sub

    Function MakeSSCC() As String
        '   ADD_FIELD("MEM", "GS1", "BIGINT")
        '  ADD_FIELD("YLIKA", "PIECESPERPALLET", "BIGINT")
        Dim YLIKA As New DataTable
        ExecuteSQLQuery("update MEM SET GS1=ISNULL(GS1,0)+1 WHERE ID=1", YLIKA)
        Dim cGS1 As String = GetValue("select GS1 FROM MEM WHERE ID=1")
        Dim L As Long = Val(cGS1)
        Dim F As String = Format(L, "0000000")
        F = "000521301114" + F
        ' ean > 9,999,999 => "001521..."

        MakeSSCC = F + findCheckDigit(F)

        'Dim CH As String
        'Dim SUMA As Long = 0

        'For K As Integer = 1 To 17
        '    CH = Mid(F, K, 1)
        '    If K Mod 2 = 1 Then
        '        SUMA = SUMA + 3 * Val(CH)
        '    Else
        '        SUMA = SUMA + Val(CH)
        '    End If
        'Next
        'Dim N As Integer = SUMA Mod 10
        'If N > 0 Then
        '    N = 10 - N
        'End If
        'Dim CN As String = Format(N, "0")

        'MakeGS1 = F + CN




    End Function


    Public Function findCheckDigit(ByVal f As String) As String
        Dim CH As String
        Dim SUMA As Long = 0

        For K As Integer = 1 To Len(f)
            CH = Mid(f, K, 1)
            If K Mod 2 = 1 Then
                SUMA = SUMA + 3 * Val(CH)
            Else
                SUMA = SUMA + Val(CH)
            End If
        Next
        Dim N As Integer = SUMA Mod 10
        If N > 0 Then
            N = 10 - N
        End If
        Dim CN As String = Format(N, "0")

        findCheckDigit = CN
    End Function






    Function GetValue(ByVal query As String) As String
        Dim YLIKA As New DataTable

        ExecuteSQLQuery(query, YLIKA)
        GetValue = YLIKA.Rows(0).Item(0).ToString

        YLIKA = Nothing


    End Function



    Public Function DataSourceConnection_Report()
        'If Split(tmpStr, ":")(4) = "1" Then
        '    'mReport.DataSourceConnections
        '    'mReport()
        '    'mReport.DataSourceConnections(0).SetConnection(Split(tmpStr, ":")(1), "SaleInv_DB", Split(tmpStr, ":")(2), Split(tmpStr, ":")(3))
        '    mReport.DataSourceConnections(0).SetConnection(Split(tmpStr, ":")(1), "SaleInv_DB", False)
        '    'MsgBox(Split(tmpStr, ":")(2) & "  " & Split(tmpStr, ":")(3))
        '    mReport.DataSourceConnections(0).SetLogon(Split(tmpStr, ":")(2), Split(tmpStr, ":")(3))
        'Else

        '    mReport.DataSourceConnections(0).SetConnection(Split(tmpStr, ":")(1), "SaleInv_DB", True)
        'End If
        ''MsgBox(mReport.DataSourceConnections(0).ServerName.ToString)
        'Return 0
    End Function

    Public Sub FillListView(ByVal sqlData As DataTable, ByVal lvList As ListView, ByVal imageID As Integer)
        Dim i As Integer
        Dim j As Integer
        'lvList.Refresh()
        lvList.Clear()
        For i = 0 To sqlData.Columns.Count - 1
            lvList.Columns.Add(sqlData.Columns(i).ColumnName)
        Next i

        For i = 0 To sqlData.Rows.Count - 1
            lvList.Items.Add(sqlData.Rows(i).Item(0), imageID)
            For j = 1 To sqlData.Columns.Count - 1
                If Not IsDBNull(sqlData.Rows(i).Item(j)) Then
                    lvList.Items(i).SubItems.Add(sqlData.Rows(i).Item(j))
                Else
                    lvList.Items(i).SubItems.Add("")
                End If
            Next j
        Next i

        For i = 0 To sqlData.Columns.Count - 1
            xsize = lvList.Width / sqlData.Columns.Count - 8
            'MsgBox(xsize)
            'If xsize > 1440 Then
            lvList.Columns(i).Width = xsize
            'Else
            '   lvList.Columns(i).Width = 2000
            'End If
            'lvList.Columns(i).AutoResize(ColumnHeaderAutoResizeStyle.HeaderSize)
        Next i
    End Sub


    'his articles helps user to Insert, Update, Delete, and Select data in Excel files using the OLEDBDataProvider in VB.NET.

    'Here is the connection string to connect with Excel using OleDBDataProvider:

    'Hide   Copy Code
    'Here is the code on the button click event to select and insert data in an Excel file:

    'Hide   Shrink    Copy Code
    'Private Sub EXCEL7(ByVal FILE As String)
    '    Dim connstring As String = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    '    "Data Source=" + FILE + ";Extended Properties=""Excel 8.0;HDR=YES;"""
    '    Dim pram As OleDbParameter
    '    Dim dr As DataRow
    '    Dim olecon As OleDbConnection
    '    Dim olecomm As OleDbCommand
    '    Dim olecomm1 As OleDbCommand
    '    Dim oleadpt As OleDbDataAdapter
    '    Dim ds As DataSet
    '    Try
    '        olecon = New OleDbConnection
    '        olecon.ConnectionString = connstring
    '        olecomm = New OleDbCommand
    '        olecomm.CommandText = _
    '           "Select FirstName, LastName, Age, Phone from [Sheet1$]"
    '        olecomm.Connection = olecon
    '        olecomm1 = New OleDbCommand
    '        olecomm1.CommandText = "Insert into [Sheet1$] " & _
    '            "(FirstName, LastName, Age, Phone) values " & _
    '            "(@FName, @LName, @Age, @Phone)"
    '        olecomm1.Connection = olecon
    '        pram = olecomm1.Parameters.Add("@FName", OleDbType.VarChar)
    '        pram.SourceColumn = "FirstName"
    '        pram = olecomm1.Parameters.Add("@LName", OleDbType.VarChar)
    '        pram.SourceColumn = "LastName"
    '        pram = olecomm1.Parameters.Add("@Age", OleDbType.VarChar)
    '        pram.SourceColumn = "Age"
    '        pram = olecomm1.Parameters.Add("@Phone", OleDbType.VarChar)
    '        pram.SourceColumn = "Phone"
    '        oleadpt = New OleDbDataAdapter(olecomm)
    '        ds = New DataSet
    '        olecon.Open()
    '        oleadpt.Fill(ds, "Sheet1")
    '        If IsNothing(ds) = False Then
    '            dr = ds.Tables(0).NewRow
    '            dr("FirstName") = "Raman"
    '            dr("LastName") = "Tayal"
    '            dr("Age") = 24
    '            dr("Phone") = 98989898
    '            ds.Tables(0).Rows.Add(dr)
    '            oleadpt = New OleDbDataAdapter
    '            oleadpt.InsertCommand = olecomm1
    '            Dim i As Integer = oleadpt.Update(ds, "Sheet1")
    '            MessageBox.Show(i & " row affected")
    '        End If
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    Finally
    '        olecon.Close()
    '        olecon = Nothing
    '        olecomm = Nothing
    '        oleadpt = Nothing
    '        ds = Nothing
    '        dr = Nothing
    '        pram = Nothing
    '    End Try
    'End Sub


    'Imports Excel = Microsoft.Office.Interop.Excel
    'Public Class Form1

    '    Private Sub Button1_Click(ByVal sender As System.Object, _
    '    ByVal e As System.EventArgs) Handles Button1.Click
    '        MsgBox(Read_from_excel("c:\test.xlsx", "sheet1", 1, 1))

    '    End Sub

    '#Region "Read and write to excel file, use functions Read_from_excel and Text_to_excel"
    Public Function Read_from_excel(ByVal filename As String, ByVal sheetname As String, ByVal row As Integer, ByVal column As Integer)
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open(filename)
        xlWorkSheet = xlWorkBook.Worksheets(sheetname)

        Dim value As String
        value = xlWorkSheet.Cells(row, column).value

        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)
        Return value
    End Function

    Public Sub Text_to_excel(ByVal filename As String, ByVal sheetname As String, ByVal row As Integer, ByVal column As Integer, ByVal text As String)
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open(filename)
        xlWorkSheet = xlWorkBook.Worksheets(sheetname)

        xlWorkSheet.Cells(row, column) = text

        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)
    End Sub


    Public Sub mreleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    '#End Region

    Public Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


    Private Function getMacAddress() As String
        Try
            Dim adapters As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces()
            Dim adapter As NetworkInterface
            Dim myMac As String = String.Empty

            For Each adapter In adapters
                Select Case adapter.NetworkInterfaceType
                    'Exclude Tunnels, Loopbacks and PPP
                    Case NetworkInterfaceType.Tunnel, NetworkInterfaceType.Loopback, NetworkInterfaceType.Ppp
                    Case Else
                        If Not adapter.GetPhysicalAddress.ToString = String.Empty And Not adapter.GetPhysicalAddress.ToString = "00000000000000E0" Then
                            myMac = adapter.GetPhysicalAddress.ToString
                            Exit For ' Got a mac so exit for
                        End If

                End Select
            Next adapter

            Return myMac
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function


    'This is the function that returns all the MAC addresses in an string array

    Public Function GETMAC() As String()
        Dim MACAddresses(0) As String
        Dim i As Integer = 0
        Dim NIC As NetworkInterface

        For Each NIC In NetworkInterface.GetAllNetworkInterfaces
            ReDim Preserve MACAddresses(i)
            MACAddresses(i) = String.Format("{0}", NIC.GetPhysicalAddress())
            i += 1
        Next
        Return MACAddresses
    End Function

    '  Dim mc As System.Management.ManagementClass
    '  Dim mo As ManagementObject
    'mc = New ManagementClass("Win32_NetworkAdapterConfiguration")
    '  Dim moc As ManagementObjectCollection = mc.GetInstances()
    'For Each mo In moc
    '   If mo.Item("IPEnabled") = True Then
    '      ListBox1.Items.Add("MAC address " & mo.Item("MacAddress").ToString())
    '   End If
    'Next

    'ADD THIS LINE 
    'OF CODE INSIDE THE 
    'WINDOWS FORM GENERATED CODE

    '<System.STAThread()> _
    'WIN XP THEMS STIS FORMES
    'Public 
    'Shared 
    '   Sub Main()

    '       System.Windows.Forms.Application.EnableVisualStyles()

    '   System.Windows.Forms.Application.Run(New 
    'frmDecode)  ' replace frmDecode by the name of your 
    'form!!!

    '       End
    'Sub
    Sub ADD_FIELD(ByVal Table As String, ByVal FIELD As String, ByVal FIELDTYPE As String)
        '-----------------------------------------------------------------------------------
        Dim R As New DataTable
        Dim sql As String
        sql = "SELECT TOP 0 * FROM " + Table + ";"
        ExecuteSQLQuery(sql)
        Dim k As Integer, OK As String
        OK = 0
        For Each column As DataColumn In sqlDT.Columns
            If FIELD = column.ColumnName Then
                OK = 1
            End If
        Next

        If OK = 0 Then
            sql = "alter table " + Table + " ADD " + FIELD + " " + FIELDTYPE
            ExecuteSQLQuery(sql)
        End If

    End Sub

    Public Function DefaultPrinterName() As String
        Dim oPS As New System.Drawing.Printing.PrinterSettings

        Try
            DefaultPrinterName = oPS.PrinterName
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Capturing Default Printer", MessageBoxButtons.OK)
        Finally
            oPS = Nothing
        End Try
    End Function





End Module
