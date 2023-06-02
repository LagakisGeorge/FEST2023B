Imports System.Data.OleDb
Imports System.Net.NetworkInformation
Imports Excel = Microsoft.Office.Interop.Excel



Public Class TIMOLOGIA

    Dim GDB As New ADODB.Connection

    '  Private Sub cmdCommand1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCommand1.Click

    Dim R As New ADODB.Recordset



    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ' ExecuteSQLQuery("SELECT * FROM EID WHERE ONO LIKE '" + EIDOS.Text + "%'")
        If Len(EIDKOD.Text) = 0 Then
            FillListBox("SELECT ONO,KOD   FROM YLIKA WHERE  ONO LIKE '" + EIDOS.Text + "%'", ListBox1)
        Else
            FillListBox("SELECT ONO,KOD   FROM YLIKA WHERE  KOD LIKE '" + EIDKOD.Text + "%'", ListBox1)
        End If





    End Sub



    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim N As Long

        'Dim i As Integer
        'Dim str() As String

        'str() = Split("0;1;2;3;4;5", ";")

        'For i = 0 To 5
        '    List1.AddItem(str(i), i)
        'Next i




        N = ListBox1.SelectedIndex
        If N = -1 Then Exit Sub

        ONO.Text = Split(ListBox1.Items(N).ToString, ";")(0)

        ' If N = -1 Then Exit Sub

        KOD.Text = Split(ListBox1.Items(N).ToString, ";")(1)



        FillListBox("SELECT left(ATIM+space(10),10)+CONVERT(CHAR(10),HME,3),str(YPOL,6,1)+' '+PROM    FROM TIMS WHERE YPOL>0  AND KOD='" + KOD.Text + "' ORDER BY HME", Excelimport)
        Excelimport.Items.Insert(0, "ΠΑΡ/ΚΟ    HMEΡ/ΝΙΑ      ΥΠΟΛ. ΠΡΟΜΗΘΕΥΤΗΣ ")

        katax.Enabled = True





        HME.Focus()
        ' ExecuteSQLQuery("UPDATE MEM SET MEMO='" + Split(ListBox1.Items(N).ToString, ";")(0) + "' WHERE ID=1")
    End Sub

    Private Sub katax_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles katax.Click
        Dim sql As String
        If Val(POSO.Text) = 0 Then
            MsgBox("Ποσότητα;")
            Exit Sub
        End If

        If Len(KOD.Text) = 0 Then
            MsgBox("ΕΙΔΟΣ ;")
            Exit Sub
        End If





        If Len(ATIM.Text) = 0 Then
            MsgBox("Αρ.Τιμολογίου;")
            Exit Sub
        End If


        Dim MFI As New DataTable
        ExecuteSQLQuery("select count(*) from TIMS WHERE ATIM='" + ATIM.Text + "' AND HME='" + Format(HME.Value, "MM/dd/yyyy") + "'", MFI)

        If MFI(0)(0) > 0 Then
            MsgBox("ΥΠΑΡΧΕΙ ΗΔΗ ΤΟ ΤΙΜΟΛΟΓΙΟ")
        End If













        Dim mposo As String
        Dim mgr As Double
        mgr = Val(Replace(POSO.Text, ",", "."))
        If Mid(monades.Text, 1, 2) = "GR" Then
            mgr = mgr / 1000
        ElseIf Mid(monades.Text, 1, 2) = "KG" Then
            mgr = mgr

        ElseIf Mid(monades.Text, 1, 2) = "TE" Then
            mgr = mgr

        ElseIf Mid(monades.Text, 1, 2) = "TO" Then ' TON
            mgr = 1000 * mgr
        End If
        mposo = Replace(Str(mgr), ",", ".")

        sql = "insert into TIMS (HME,POSO,ATIM,KOD,PROM,YPOL) VALUES("
        sql = sql + "'" + Format(HME.Value, "MM/dd/yyyy") + "',"
        sql = sql + mposo + ","
        sql = sql + "'" + ATIM.Text + "',"
        sql = sql + "'" + LTrim(KOD.Text) + "',"
        sql = sql + "'" + Split(PROMITH.Text, ";")(0).ToString + "',"
        sql = sql + mposo + ")"

        Try
            ExecuteSQLQuery(sql)
            katax.Enabled = False

        Catch ex As Exception
            MsgBox("Δεν Αποθηκεύθηκε")
            Exit Sub
        End Try


        Dim N As Long





        'N = ListBox1.SelectedIndex
        'ONO.Text = Split(ListBox1.Items(N).ToString, ";")(0)
        'KOD.Text = Split(ListBox1.Items(N).ToString, ";")(1)


        FillListBox("SELECT left(ATIM+space(10),10)+CONVERT(CHAR(10),HME,3),str(YPOL,6,1)+' '+PROM    FROM TIMS WHERE YPOL>0  AND KOD='" + KOD.Text + "' ORDER BY HME", Excelimport)
        Excelimport.Items.Insert(0, "ΠΑΡ/ΚΟ    HMEΡ/ΝΙΑ      ΥΠΟΛ. ΠΡΟΜΗΘΕΥΤΗΣ ")
        katax.Enabled = True


        Excelimport.Items.Add(ATIM.Text + ";" + Format(HME.Value, "dd/MM/yyyy") + ";" + mposo)







    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        For k = 0 To 20
            ergates.widths(7) = 100
        Next
        '  ExecuteSQLQuery("update YLIKA SET N1=(SELECT TOP 1 AEG FROM EID WHERE KOD=YLIKA.KOD) ")

        Dim Mn1 As String
        Mn1 = "" ' Split(KATHG.Text, ";")(0)


        ergates.Text = "Αρχείο Υλικών"
        ergates.Label1.Text = "SELECT HME,POSO,ATIM,KOD,PROM,YPOL,ID  FROM TIMS ORDER BY HME "

        ' ergates.MdiParent = Me
        ergates.WindowState = FormWindowState.Maximized
        ergates.STHLHONOMATOS_ID = 6
        ergates.STHLHTOY_ID = 6
        ergates.widths(1) = 100
        ergates.QUERY_AFTER = " " 'update YLIKA SET N1=" + Mn1 + " WHERE N1 IS NULL"
        For KK As Integer = 0 To 6
            ergates.widths(KK) = 50
        Next
        ergates.Label2.Text = "Τιμολόγια"
        ergates.widths(0) = 100
        gMenu = 24
        ergates.Show()
    End Sub

    Private Sub TIMOLOGIA_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        monades.Text = monades.Items(0).ToString
        GDB.Open(gConnect)
        FILLComboBox("select EPO+';'+AFM,space(30)+AFM FROM PEL WHERE EIDOS='r' order by EPO ", PROMITH)

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, _
                              ByVal e As System.EventArgs) _
            Handles Button3.Click

        Dim ANS = MsgBox("Είναι η Γραμογράφηση Τιμ.Αγοράς ΟΚ  ??", MsgBoxStyle.YesNo)

        If ANS = MsgBoxResult.No Then
            Exit Sub
        End If

        OpenFileDialog1.ShowDialog()




        Dim ISOK As Boolean = True
        Dim NEA_TIM As Integer = 0
        IMPORT_TIMAGOR(1, ISOK, NEA_TIM)

        If ISOK = False Then
            MsgBox("ΔΙΟΡΘΩΣΤΕ ΤΑ ΕΙΔΗ ΠΟΥ ΔΕΝ ΥΠΑΡΧΟΥΝ")
        Else
            If NEA_TIM = 0 Then
                MsgBox("ΔΕΝ ΥΠΑΡΧΟΥΝ NEA TIMOΛOΓIA")
                Exit Sub

            Else
                ISOK = True
                NEA_TIM = 0
                IMPORT_TIMAGOR(0, ISOK, NEA_TIM)
            End If

        End If

    End Sub

    Private Sub IMPORT_TIMAGOR(ByVal ELEGXOS As Integer, ByRef ISOK As Boolean, ByRef NEA_TIM As Integer)
        Dim ANS As Integer
        Dim r As New ADODB.Recordset

        Dim line As String
        Dim line2 As String
        Dim cPel As String
        Dim cEID As String
        Dim mHME As String
        Dim mAtim As String

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
        xlWorkSheet = xlWorkBook.Worksheets(1)
        'display the cells value B2
        '    MsgBox(xlWorkSheet.Cells(6, 1).value)
        'edit the cell with new value

        'xlWorkSheet.Cells(7, 2) = onomaProion
        'xlWorkSheet.Cells(8, 1) = "ΠΟΣΟΤΗΤΑ: " + Str(minTem)
        'xlWorkSheet.Cells(15, 1) = TELBARCODE
        'xlWorkSheet.Cells(18, 2) = kodPROION
        ''Globals.xlworkSheet.PrintOut(From:=1, To:=1, Copies:=2, Preview:=True)

        Dim N As Integer = 11
        Dim D As String
        Dim flagPel As Integer
        Dim flagEID As Integer
        Dim merror As Integer = 0


        If ELEGXOS = 0 Then
            GDB.BeginTrans()
        End If

        Dim mon As String


        '*********************** LOOP EXCEL ***************************************************
        Do

            ' line = xlWorkSheet.Cells(N, 1).VALUE.ToString

            If N > xlWorkSheet.UsedRange.Rows.Count Then

                If xlWorkSheet.Cells(N, 4).VALUE = Nothing Then
                    Exit Do
                End If

            End If

            'If xlWorkSheet.Cells(N, 4).VALUE = Nothing Then
            '    Exit Do
            'End If
            Dim YPARXEI_HDH As Boolean = False
            Try

                'ΨΑΧΝΩ ΝΑ ΔΡΩ ΤΟΝ ΚΩΔΙΚΟ ΤΟΥ ΠΡΟΜΗΘΕΥΤΗ
                flagPel = 0

                Try
                    '==========================================================================
                    While True

                        If N > xlWorkSheet.UsedRange.Rows.Count Then

                            If xlWorkSheet.Cells(N, 4).VALUE = Nothing Then
                                Exit While
                            End If

                        End If

                        If String.IsNullOrEmpty(xlWorkSheet.Cells(N, 4).VALUE) Then

                            cPel = "ΩΩΩΩ"

                        Else

                            cPel = xlWorkSheet.Cells(N, 4).VALUE.ToString
                            D = xlWorkSheet.Cells(N, 2).VALUE.ToString
                        End If

                        If IsNumeric(Mid(cPel, 1, 6)) Then
                            mHME = Mid(xlWorkSheet.Cells(N, 2).VALUE.ToString, 5, 10)
                            mHME = Mid(mHME, 4, 2) + "/" + Mid(mHME, 1, 2) + "/" + Mid(mHME, 7, 4)
                            mAtim = xlWorkSheet.Cells(N, 3).VALUE.ToString



                            If ELEGXOS = 1 Then
                                Dim MFI As New DataTable
                                ExecuteSQLQuery("select count(*) from TIMS WHERE ATIM='" + mAtim + "' AND HME='" + mHME + "'", MFI)
                                If MFI(0)(0) > 0 Then
                                    ListBox1.Items.Add(mAtim + ";" + " ΥΠΑΡΧΕΙ ΗΔΗ ΤΟ ΤΙΜΟΛΟΓΙΟ ")
                                    YPARXEI_HDH = True
                                Else
                                    ListBox1.Items.Add(" ;" + mAtim + ";" + "NEO TIM. ")
                                    NEA_TIM = NEA_TIM + 1
                                End If
                            Else  ' ΕΛΕΓΧΩ ΑΠΟ ΤΟ LISTBOX

                                For L = 0 To ListBox1.Items.Count - 1
                                    If mAtim = Split(ListBox1.Items(L).ToString, ";")(0) Then
                                        YPARXEI_HDH = True
                                    End If

                                Next



                            End If






                            If Len(mAtim) > 10 Then

                            End If

                            flagPel = 1
                            'line=
                            N = N + 1
                            Exit While
                        End If

                            N = N + 1
                    End While
                    '==========================================================================
                Catch ex As Exception
                    MsgBox("ΛΑΘΟΣ ΣΤΟ WHILE ΠΡΟΜΗΘΕΥΤΗ")
                End Try

                'ΔΕΝ ΒΡΗΚΑ ΠΕΛΑΤΗ ΒΓΕΣ ΑΠΟ LOOP
                If flagPel = 0 Then
                    Exit Do
                End If



                'ΨΑΧΝΩ ΝΑ BΡΩ ΤΟYΣ ΚΩΔΙΚΟΥΣ ΤΩΝ  ΕΙΔΩΝ
                flagEID = 0
                '-------------------------------------------------------------------------------------
                While True '  Not xlWorkSheet.Cells(N, 2).VALUE = Nothing

                    'If N > xlWorkSheet.UsedRange.Rows.Count Then
                    If xlWorkSheet.Cells(N, 2).VALUE = Nothing Then
                        Exit While
                    End If

                    'End If

                    cEID = Mid(xlWorkSheet.Cells(N, 2).VALUE.ToString, 8, 6)

                    If IsNumeric(cEID) Then
                        flagEID = 1
                        'βρηκα ενα ειδος και κανω INSERT TO RECORD
                        '..ADD_RECORD
                        'r.Open("select * FROM TIMS WHERE KODSYNOD='" + D + "' AND  KOD='" + cPel + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                        'καθαρισμα αριθμου
                        Dim mPoso As String '= xlWorkSheet.Cells(N, 9).VALUE.ToString

                        'αν η στηλη με τις ποσότητες δεν ειναι στο Ι θα είναι στο J
                        If String.IsNullOrEmpty(xlWorkSheet.Cells(N, 9).VALUE) Then

                            mPoso = xlWorkSheet.Cells(N, 10).VALUE.ToString

                        Else
                            mPoso = xlWorkSheet.Cells(N, 9).VALUE.ToString
                        End If

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
                        If InStr(mAtim, "ΔΑ") > 0 Then
                            mPoso = -mPoso
                        End If


                        'If ELEGXOS = 1 Then
                        r.Open("select * FROM YLIKA WHERE KOD='" + cEID + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If r.EOF Then
                            MsgBox(" ΔΕΝ ΥΠΑΡΧΕΙ Ο ΚΩΔΙΚΟΣ " + cEID + ". ΔΙΑΚΟΠΗΚΕ Η ΔΙΑΔΙΚΑΣΙΑ")
                            merror = 1
                            Exit Do
                        End If
                        'End If

                        'If r.EOF Then
                        '    MsgBox(" ΔΕΝ ΥΠΑΡΧΕΙ Ο ΚΩΔΙΚΟΣ " + cEID + ". ΔΙΑΚΟΠΗΚΕ Η ΔΙΑΔΙΚΑΣΙΑ")
                        '    merror = 1
                        '    Exit Do

                        'Else

                        mon = r("c1").Value.ToString

                        If Val(Mid(mon, 1, 3)) > 0 Then
                            mPoso = Replace(Str(Val(mPoso) * 100), ",", ".")

                        Else

                        End If

                        'Dim MFI As New DataTable
                        'ExecuteSQLQuery("select count(*) from TIMS WHERE ATIM='" + mAtim + "' AND HME='" + mHME + "'", MFI)

                        'If MFI(0)(0) > 0 Then
                        '    MsgBox("ΥΠΑΡΧΕΙ ΗΔΗ ΤΟ ΤΙΜΟΛΟΓΙΟ")
                        '    ' Else

                        'End If
                        If YPARXEI_HDH = False Then
                            If ELEGXOS = 0 Then
                                Try
                                    GDB.Execute("INSERT INTO TIMS (C1,KOD,PROM,ATIM,HME,POSO,YPOL) VALUES ('" + Mid(cPel, 1, 6) + "','" + cEID + "','" + Mid(cPel, 8, 16) + "','" + mAtim + "','" + mHME + "'," + mPoso + "," + mPoso + ")")

                                Catch ex As Exception
                                    MsgBox("ΛΑΘΟΣ ΣΤΟ ΤΙΜΟΛΟΓΙΟ " + Str(N))
                                End Try
                            End If


                        End If



                        '     End If

                        r.Close()

                    Else ' ΤΕΛΕΙΩΣΕ ΤΟ ΤΙΜΟΛΟΓΙΟ ΚΑΙ ΠΑΩ ΠΑΡΑΚΑΤΩ
                        Exit While

                    End If

                    N = N + 1
                End While
                ''-------------------------------------------------------------------------------------
                'ΔΕΝ ΒΡΗΚΑ ΠΕΛΑΤΗ ΒΓΕΣ ΑΠΟ LOOP
                'If flagPel = 0 Then
                '    Exit Do
                'End If

                ' GDB.Execute("update SYNTAGES SET POSOSTO=" + Replace(Str(xlWorkSheet.Cells(N, 11).VALUE / 1), ",", ".") + " WHERE  KOD='" + cPel + "' AND KODSYNOD='" + D + "'")
            Catch ex As Exception
                MsgBox("ΛΑΘΟΣ ΣΤΗΝ ΣΕΙΡΑ " + Str(N) + " " + cPel + " == " + D + "--" + ex.Message)
                'Exit Do

            End Try

            N = N + 1
            Me.Text = N
        Loop Until False 'xlWorkSheet.Cells(N, 9).VALUE Is Nothing
        '*********************** LOOP EXCEL ***************************************************
        '    MsgBox("OK " + Str(N))


        If ELEGXOS = 0 Then
            ANS = MsgBox("ΝΑ ΑΠΟΘΗΚΕΥΤΟΥΝ ΟΙ ΕΙΣΑΓΩΓΕΣ ΤΙΜΟΛΟΓΙΩΝ (ΣΕΙΡΕΣ EXCEL " + Str(N) + " )  ", MsgBoxStyle.YesNo)

            If ANS = MsgBoxResult.Yes Then
                GDB.CommitTrans()

            Else
                GDB.RollbackTrans()
            End If
        End If




        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        If merror = 1 Then
            ISOK = False
        End If


    End Sub

    Private Sub PYLONIMPORT_TIMAGOR(ByVal ELEGXOS As Integer, ByRef ISOK As Boolean, ByRef NEA_TIM As Integer)

        Dim mDAY, mMONTH, mYEAR As Integer
        mDAY = HME.Value.Day
        mMONTH = HME.Value.Month
        mYEAR = HME.Value.Year
        'ΕΝΗΜΕΡΩΝΟΥΝ ΜΕ ΑΓΟΡΕΣ ΤΗΝ ΑΠΟΘΗΚΗ ΤΑ: ΤΔΠ, ΔΠΡ (ΑΦΑΙΡΕΙΣ ΤΙΑ)   'ΤΔΠ-',ΔΠΡ-'

        'ΕΝΗΜΕΡΩΝΟΥΝ ΜΕ ΕΠΙΣΤΡΟΦΕΣ ΣΤΟΝ ΠΡΟΜΗΘΕΥΤΗ ΤΗΝ ΑΠΟΘΗΚΗ ΤΑ : ΔΕΠ, ΕΠΠ  'ΔΕΠ-', 'ΕΠΠ-'

        Dim GPYL As New ADODB.Connection
        GPYL.Open("DSN=PYLONTECHNOPLASTIKI;uid=sa;pwd=p@ssw0rd")
        Dim R32 As New ADODB.Recordset
        Dim SQLP As String = "SELECT    HEDOCCODE AS ATIM,HEDOCNUM,[HEITEMCODE] AS KODE"
        SQLP = SQLP + ",[HEITEMDESCRIPTION] AS ONO ,[HEAQTY] AS [POSOTHTA],[HEOFFICIALDATE] AS HME,"
        SQLP = SQLP + "(SELECT HECODE FROM HESUPPLIERS WHERE HEID=C.HEBILLSPLRID) AS KODIKOSPELATH,C.HEENTITYDESCR  FROM [HECENTLINES] E "
        SQLP = SQLP + " INNER JOIN [HEDOCENTRIES] T  ON E.HEDENTID=T.HEID LEFT JOIN [HECOMMERCIALENTRIES] C ON E.HEDENTID=C.HEDENTID "
        SQLP = SQLP + "WHERE T.HESTATUS=0 AND LEFT(HEDOCCODE,4) IN ('ΤΔΠ-','ΔΠΡ-','ΔΕΠ-', 'ΕΠΠ-') AND DAY(HEOFFICIALDATE)=" + Str(mDAY) + " AND MONTH(HEOFFICIALDATE)=" + Str(mMONTH) + " AND YEAR(HEOFFICIALDATE)=" + Str(mYEAR) + "  ORDER BY HME ,ATIM"


        R32.Open(SQLP, GPYL, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)























        Dim ANS As Integer
        Dim r As New ADODB.Recordset

        Dim line As String
        Dim line2 As String
        Dim cPel As String
        Dim cEID As String
        Dim mHME As String
        Dim mAtim As String

        'Dim xlApp As Excel.Application
        'Dim xlWorkBook As Excel.Workbook
        'Dim xlWorkSheet As Excel.Worksheet

        'xlApp = New Excel.ApplicationClass
        'xlWorkBook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
        'xlWorkSheet = xlWorkBook.Worksheets(1)
        'display the cells value B2
        '    MsgBox(xlWorkSheet.Cells(6, 1).value)
        'edit the cell with new value

        'xlWorkSheet.Cells(7, 2) = onomaProion
        'xlWorkSheet.Cells(8, 1) = "ΠΟΣΟΤΗΤΑ: " + Str(minTem)
        'xlWorkSheet.Cells(15, 1) = TELBARCODE
        'xlWorkSheet.Cells(18, 2) = kodPROION
        ''Globals.xlworkSheet.PrintOut(From:=1, To:=1, Copies:=2, Preview:=True)

        Dim N As Integer = 11
        Dim D As String
        Dim flagPel As Integer
        Dim flagEID As Integer
        Dim merror As Integer = 0


        If ELEGXOS = 0 Then
            GDB.BeginTrans()
        End If

        Dim mon As String


        '*********************** LOOP EXCEL ***************************************************
        'Do
        Do While Not R32.EOF
            ' line = xlWorkSheet.Cells(N, 1).VALUE.ToString

            'If N > xlWorkSheet.UsedRange.Rows.Count Then

            '    If xlWorkSheet.Cells(N, 4).VALUE = Nothing Then
            '        Exit Do
            '    End If

            'End If

          
            Dim YPARXEI_HDH As Boolean = False
            Try

                'ΨΑΧΝΩ ΝΑ ΔΡΩ ΤΟΝ ΚΩΔΙΚΟ ΤΟΥ ΠΡΟΜΗΘΕΥΤΗ
                flagPel = 0

                Try
                    '==========================================================================
                    ' While True

                    'If N > xlWorkSheet.UsedRange.Rows.Count Then

                    '    If xlWorkSheet.Cells(N, 4).VALUE = Nothing Then
                    '        Exit While
                    '    End If

                    'End If

                    'If String.IsNullOrEmpty(xlWorkSheet.Cells(N, 4).VALUE) Then

                    '    cPel = "ΩΩΩΩ"

                    'Else

                    '    cPel = xlWorkSheet.Cells(N, 4).VALUE.ToString
                    '    D = xlWorkSheet.Cells(N, 2).VALUE.ToString
                    'End If

                    cPel = R32("KODIKOSPELATH").Value.ToString
                    mHME = Format(R32("HME").Value, "MM/dd/yyyy") '    .ToString
                    mAtim = R32("ATIM").Value.ToString


                    'If IsNumeric(Mid(cPel, 1, 6)) Then
                    ' mHME = Mid(xlWorkSheet.Cells(N, 2).VALUE.ToString, 5, 10)
                    '  mHME = Mid(mHME, 4, 2) + "/" + Mid(mHME, 1, 2) + "/" + Mid(mHME, 7, 4)
                    ' mAtim = xlWorkSheet.Cells(N, 3).VALUE.ToString



                    If ELEGXOS = 1 Then
                        Dim MFI As New DataTable
                        ExecuteSQLQuery("select count(*) from TIMS WHERE ATIM='" + mAtim + "' AND HME='" + mHME + "'", MFI)
                        If MFI(0)(0) > 0 Then
                            ListBox1.Items.Add(mAtim + ";" + " ΥΠΑΡΧΕΙ ΗΔΗ ΤΟ ΤΙΜΟΛΟΓΙΟ ")
                            YPARXEI_HDH = True
                        Else
                            ListBox1.Items.Add(" ;" + mAtim + ";" + "NEO TIM. ")
                            NEA_TIM = NEA_TIM + 1
                        End If
                    Else  ' ΕΛΕΓΧΩ ΑΠΟ ΤΟ LISTBOX

                        For L = 0 To ListBox1.Items.Count - 1
                            If mAtim = Split(ListBox1.Items(L).ToString, ";")(0) Then
                                YPARXEI_HDH = True
                            End If

                        Next



                    End If

                    flagPel = 1
                    'line=

                    'Exit While
                    ' End If


                    ' End While
                    '==========================================================================
                Catch ex As Exception
                    MsgBox("ΛΑΘΟΣ ΣΤΟ WHILE ΠΡΟΜΗΘΕΥΤΗ")
                End Try

                'ΔΕΝ ΒΡΗΚΑ ΠΕΛΑΤΗ ΒΓΕΣ ΑΠΟ LOOP
                'If flagPel = 0 Then
                'Exit Do
                ' End If



                'ΨΑΧΝΩ ΝΑ BΡΩ ΤΟYΣ ΚΩΔΙΚΟΥΣ ΤΩΝ  ΕΙΔΩΝ
                flagEID = 0
                '-------------------------------------------------------------------------------------
                ' LOOP IDIOY  TIMOLOGIOY
                While R32("ATIM").Value.ToString = mAtim '  Not xlWorkSheet.Cells(N, 2).VALUE = Nothing


            
                    'If xlWorkSheet.Cells(N, 2).VALUE = Nothing Then
                    '    Exit While
                    'End If


                    cEID = R32("KODE").Value.ToString
                    'cEID = Mid(xlWorkSheet.Cells(N, 2).VALUE.ToString, 8, 6)

                    If IsNumeric(cEID) Then
                        flagEID = 1
                        'βρηκα ενα ειδος και κανω INSERT TO RECORD
                        '..ADD_RECORD
                        'r.Open("select * FROM TIMS WHERE KODSYNOD='" + D + "' AND  KOD='" + cPel + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

                        'καθαρισμα αριθμου
                        Dim mPoso As String '= xlWorkSheet.Cells(N, 9).VALUE.ToString

                        'αν η στηλη με τις ποσότητες δεν ειναι στο Ι θα είναι στο J
                        mPoso = R32("POSOTHTA").Value
                        'ΔΕΠ-', 'ΕΠΠ-'
                        If InStr(mAtim, "ΔΕΠ") > 0 Or InStr(mAtim, "ΕΠΠ") > 0 Then
                            mPoso = -mPoso
                        End If


                        'If ELEGXOS = 1 Then
                        r.Open("select * FROM YLIKA WHERE KOD='" + cEID + "'", GDB, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                        If r.EOF Then
                            MsgBox(" ΔΕΝ ΥΠΑΡΧΕΙ Ο ΚΩΔΙΚΟΣ " + cEID + ". ΔΙΑΚΟΠΗΚΕ Η ΔΙΑΔΙΚΑΣΙΑ")
                            merror = 1
                            Exit Do
                        End If
                        mon = r("c1").Value.ToString

                        If Val(Mid(mon, 1, 3)) > 0 Then
                            mPoso = Replace(Str((mPoso) * 100), ",", ".")

                        Else

                        End If


                        If YPARXEI_HDH = False Then
                            If ELEGXOS = 0 Then
                                Try
                                    GDB.Execute("INSERT INTO TIMS (C1,KOD,PROM,ATIM,HME,POSO,YPOL) VALUES ('" + Mid(cPel, 1, 6) + "','" + cEID + "','" + Mid(cPel, 8, 16) + "','" + mAtim + "','" + mHME + "'," + Replace(Str((mPoso)), ",", ".") + "," + Replace(Str((mPoso)), ",", ".") + ")")

                                Catch ex As Exception
                                    MsgBox("ΛΑΘΟΣ ΣΤΟ ΤΙΜΟΛΟΓΙΟ " + Str(N))
                                End Try
                            End If


                        End If



                        '     End If

                        r.Close()

                    Else ' ΤΕΛΕΙΩΣΕ ΤΟ ΤΙΜΟΛΟΓΙΟ ΚΑΙ ΠΑΩ ΠΑΡΑΚΑΤΩ
                        ' Exit While

                    End If


                        R32.MoveNext()

                        If R32.EOF Then
                            Exit While
                        End If


                        N = N + 1
                End While
                ''-------------------------------------------------------------------------------------
                'ΔΕΝ ΒΡΗΚΑ ΠΕΛΑΤΗ ΒΓΕΣ ΑΠΟ LOOP
                'If flagPel = 0 Then
                '    Exit Do
                'End If

                ' GDB.Execute("update SYNTAGES SET POSOSTO=" + Replace(Str(xlWorkSheet.Cells(N, 11).VALUE / 1), ",", ".") + " WHERE  KOD='" + cPel + "' AND KODSYNOD='" + D + "'")
            Catch ex As Exception
                MsgBox("ΛΑΘΟΣ ΣΤΗΝ ΣΕΙΡΑ " + Str(N) + " " + cPel + " == " + D + "--" + ex.Message)
                'Exit Do

            End Try

            N = N + 1
            Me.Text = N
        Loop 'Until False 'xlWorkSheet.Cells(N, 9).VALUE Is Nothing
        '*********************** LOOP EXCEL ***************************************************
        '    MsgBox("OK " + Str(N))


        If ELEGXOS = 0 Then
            ANS = MsgBox("ΝΑ ΑΠΟΘΗΚΕΥΤΟΥΝ ΟΙ ΕΙΣΑΓΩΓΕΣ ΤΙΜΟΛΟΓΙΩΝ (ΣΕΙΡΕΣ EXCEL " + Str(N) + " )  ", MsgBoxStyle.YesNo)

            If ANS = MsgBoxResult.Yes Then
                GDB.CommitTrans()

            Else
                GDB.RollbackTrans()
            End If
        End If




        'xlWorkBook.Close()
        'xlApp.Quit()

        'releaseObject(xlApp)
        'releaseObject(xlWorkBook)
        'releaseObject(xlWorkSheet)

        If merror = 1 Then
            ISOK = False
        End If


    End Sub







    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim HM As String = Format(HME.Value, "MM/dd/yyyy")
        Dim sql As String
        If Val(POSO.Text) = 0 Then
            MsgBox("Ποσότητα;")
            Exit Sub
        End If

        If Len(KOD.Text) = 0 Then
            MsgBox("ΕΙΔΟΣ ;")
            Exit Sub
        End If

        If Val(arpart.Text) = 0 Then
            MsgBox("ΠΑΡΤΙΔΑ ;")
            Exit Sub
        End If



        If Len(ATIM.Text) = 0 Then
            MsgBox("Αρ.Τιμολογίου;")
            Exit Sub
        End If

        Dim mposo As String
        Dim mgr As Double
        mgr = Val(Replace(POSO.Text, ",", "."))
        If Mid(monades.Text, 1, 2) = "GR" Then
            MsgBox("ΜΟΝΟ ΤΕΜΑΧΙΑ")
            Exit Sub
        End If

        If Mid(monades.Text, 1, 2) = "KG" Then
            MsgBox("ΜΟΝΟ ΤΕΜΑΧΙΑ")
            Exit Sub
        End If

        If Mid(monades.Text, 1, 2) = "TO" Then
            MsgBox("ΜΟΝΟ ΤΕΜΑΧΙΑ")
            Exit Sub
        End If

        mposo = Replace(Str(mgr), ",", ".")

        Dim R2 As New DataTable
        ExecuteSQLQuery("SELECT COUNT(*) FROM PARTIDES WHERE PARTIDA=" + Str(arpart.Text), R2)
        If R2(0)(0) > 0 Then
            MsgBox("ΥΠΑΡΧΕΙ ΗΔΗ Η ΠΑΡΤΙΔΑ ")
            Exit Sub
        End If






        sql = "INSERT INTO PARTIDES (PARTIDA,HME,KOD,TIMOLOGIA,TEMAXIA,YPOL,N1) VALUES("
        sql = sql + arpart.Text + ",'" + HM + "','" + LTrim(KOD.Text) + "','" + "απογ" + "'," + POSO.Text + "," + POSO.Text + ",0)"
        ExecuteSQLQuery(sql)
        MsgBox("ΚΑΤΕΧΩΡΗΘΗ")
        arpart.Text = ""

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim PARTIDES As New DataTable
        Dim MP As Long

        ExecuteSQLQuery("SELECT max(PARTIDA) FROM PARTIDES WHERE PARTIDA LIKE '150%' ", PARTIDES)
        If IsDBNull(PARTIDES(0)(0)) Then
            MP = 1500000
        Else
            MP = PARTIDES(0)(0) + 1
        End If

        arpart.Text = MP
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim ANS = MsgBox("NA ΓINEI EIΣAΓΩΓΗ ΑΠΟ PYLON;", MsgBoxStyle.YesNo)

        If ANS = MsgBoxResult.No Then
            Exit Sub
        End If

        ' OpenFileDialog1.ShowDialog()




        Dim ISOK As Boolean = True
        Dim NEA_TIM As Integer = 0
        PYLONIMPORT_TIMAGOR(1, ISOK, NEA_TIM)

        If ISOK = False Then
            MsgBox("ΔΙΟΡΘΩΣΤΕ ΤΑ ΕΙΔΗ ΠΟΥ ΔΕΝ ΥΠΑΡΧΟΥΝ")
        Else
            If NEA_TIM = 0 Then
                MsgBox("ΔΕΝ ΥΠΑΡΧΟΥΝ NEA TIMOΛOΓIA")
                Exit Sub

            Else
                ISOK = True
                NEA_TIM = 0
                PYLONIMPORT_TIMAGOR(0, ISOK, NEA_TIM)
            End If

        End If

    End Sub
End Class