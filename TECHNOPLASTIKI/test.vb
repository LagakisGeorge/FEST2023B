﻿Imports System.Data.OleDb
Imports System.Data.SqlClient
'Imports Microsoft.Office.Interop
'Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Public Class test
    Dim f1_row As Integer
    Dim f1_col As Integer
    Dim f2_row As Integer
    Dim f2_col As Integer
    Dim f_CanMoveLeft As Integer '' αν προκειται να κοψω στα 2 την κράτηση να μην μπορει να παει αριστερότερα
    ' γιατι την ιδια μερα θα κοιμαται σε 2 ξενοδοχεια

    Dim f_idpel As String
    Dim f_idHotRoomDays As String
    Dim f_2idHotRoomDays As String
    Dim F_REM_DAYS As String ' ΔΙΑΝΥΚΤΕΡΕΥΣΕΙΣ ΠΟΥ ΑΠΟΜΕΝΟΥΝ ΣΤΟΝ ΠΕΛΑΤΗ ΠΟΥ ΕΚΑΝΑ ΤΟ ΠΡΟΤΟ ΚΛΙΚ
    ' Declare the ContextMenuStrip control.
    Private fruitContextMenuStrip As ContextMenuStrip
    'Sub COPYALLTOCLIPBOARD()

    '    DataGridView1.SelectAll()
    '    Dim dataObj As DataObject = DataGridView1.GetClipboardContent()
    '    If Not IsDBNull(dataObj) Then
    '        Clipboard.SetDataObject(dataObj)
    '    End If
    'End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click




        ' COPYALLTOCLIPBOARD()
        'Dim xlexcel As Microsoft.Office.Interop.Excel.Application
        'Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
        'Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        'Dim misValue As Object = System.Reflection.Missing.Value
        'xlexcel = New Excel.Application()
        'xlexcel.Visible = True
        'xlWorkBook = xlexcel.Workbooks.Add(misValue)
        'xlWorkSheet = xlexcel.xlWorkbook.Worksheets(1) '    xlexcel.Workbooks.Worksheets.get_Item(1)

        Dim filename As String = "c:\mercvb\ektyp.xlsx"
        Dim sheetname As String = "Φύλλο1"
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xl As Excel.Worksheet
        xlApp = New Excel.ApplicationClass
        xlWorkBook = xlApp.Workbooks.Add 'αν ηθελα να το ανοιξω αντι για add -> Open(filename)
        xlWorkBook.Worksheets.Add()  '(1)
        xl = xlWorkBook.Worksheets(1) ' .Add

        xl.Range("a1", "z1000").Font.Size = 8
        xl.Name = "fest"
        Dim r, c As Integer
        For c = 0 To DGV.Columns.Count - 1
            xl.Cells(1, c + 1) = DGV.Columns(c).HeaderCell.Value
        Next


        For r = 0 To DGV.Rows.Count - 1
            Me.Text = r
            For c = 0 To DGV.Columns.Count - 1
                If IsDBNull(DGV.Rows(r).Cells(c).Value) Then


                Else
                    Try


                        If DGV.Rows(r).Cells(c).Value <> Nothing Then
                            If Mid(DGV.Rows(r).Cells(c).Value.ToString, 1, 2) = "_ " Then
                            Else
                                xl.Cells(r + 2, c + 1) = DGV.Rows(r).Cells(c).Value 'Columns(nn).HeaderCell.Value
                                xl.Cells(r + 2, c + 1).Interior.Color = RGB(200, 200, 200)
                                xl.Cells(r + 2, c + 1).Style.WrapText = True

                            End If
                        End If
                    Catch ex As Exception

                    End Try

                End If
                ' Else

                ' End If


            Next
        Next
        xlApp.Visible = True  'ΜΠΟΡΩ ΝΑ ΤΟ ΒΛΕΠΩ
        MsgBox("ok")

        Exit Sub



        'Dim cr As Excel.Range = xl.Cells(1, 1)
        'cr.Select()
        'xl.PasteSpecial(cr, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, True)

        'Dim test As Integer
        'Dim HTR As New DataTable
        '' ΒΑΖΩ ΕΠΙΚΕΦΑΛΙΔΕΣ
        'ExecuteSQLQuery("select MAX(DATECHECKIN) AS MAX1,MIN(DATECHECKIN) AS MIN1 from HOTROOMDAYS  ", HTR)
        'Dim mApo As Date = HTR.Rows(0)("MIN1")
        'Dim mEos As Date = HTR.Rows(0)("MAX1")
        'Dim hmeres As Integer = DateDiff("d", mApo, mEos) + 1
        'Dim hmera As Date = mApo

        'DGV.RowHeadersVisible = False
        'DGV.Columns(0).HeaderCell.Value = "Ξεν"
        'DGV.Columns(1).HeaderCell.Value = "Δωμ"
        'Dim pl As Integer = 80
        '' ΔΕΙΧΝΩ ΤΙΣ ΜΕΡΕΣ
        'For nn As Integer = 2 To hmeres + 1
        '    'DGV.Columns(nn).Width = pl
        '    ' DGV.Columns(nn).HeaderCell.Value = Format(hmera, "dd/MM")
        '    xl.Cells(1, nn) = Format(hmera, "dd/MM")
        '    hmera = DateAdd("d", 1, hmera)

        'Next




        '' DEIXNΩ ΔΩΜΑΤΙΑ & ΞΕΝΟΔΟΧΕΙΑ
        'ExecuteSQLQuery("select H.NAME,R.ROOMN from HOTROOMS R inner join HOTELS H ON H.ID=R.HOTELID ORDER BY H.NAME,R.ROOMN ", HTR)

        'For K As Integer = 0 To HTR.Rows.Count - 1
        '    ' DGV.Rows.Add()
        '    xl.Cells(K + 1, 1) = HTR.Rows(K)("name")
        '    ' DGV.Rows(K).Cells(0).Value = HTR.Rows(K)("name")
        '    ' DGV.Rows(K).Cells(1).Value = HTR.Rows(K)("roomn")
        '    xl.Cells(K + 1, 2) = HTR.Rows(K)("roomn")
        'Next


        ''diastash toy excel   hmeres: sthles
        'Dim rowsn As Integer = HTR.Rows.Count


        'ExecuteSQLQuery("select HOTELNAME,ROOMN,DATECHECKIN,ISNULL(IDPEL,0) AS IDPEL,ISNULL((SELECT EPO+'+'+ISNULL(SYNODOS,'') FROM PEL WHERE ID=IDPEL),'-') AS EPO ,ID from HOTROOMDAYS  ORDER BY HOTELNAME,ROOMN ", HTR)
        'Dim X As String, R As String
        'For K As Integer = 0 To HTR.Rows.Count - 1
        '    'ΠΡΟΣΔΙΟΡΙΖΩ ΤΗΝ ΣΕΙΡΑ
        '    X = Trim(HTR.Rows(K)("HOTELname"))
        '    R = Trim(HTR.Rows(K)("ROOMN")) ' ariumos domatioy p.x. 102
        '    Dim seira As Integer = 0
        '    For i As Integer = 1 To hmeres 'DGV.Rows.Count - 1
        '        'αν βρηκα το ξενοδοχείο(Χ) και το δωμάτιο(R) τοτε τσιμπα  την σειρά
        '        If X = xl.Cells(i, 1).ToString And R = xl.Cells(i, 2).ToString Then '  DGV.Rows(i).Cells(0).Value And R = DGV.Rows(i).Cells(1).Value Then
        '            seira = i
        '            Exit For
        '        End If
        '    Next

        '    'ΠΡΟΣΔΙΟΡΙΖΩ ΤΗΝ ΣΤΗΛΗ  sygkrinontas to HOTROOMDAYS DATECHECKIN ME TON TITLO THS STHLHS KAI TSIMPAO THN STHLH
        '    Dim cc As String = Format(HTR.Rows(K)("datecheckin"), "dd/MM")
        '    Dim sthlh As Integer = 0
        '    For i = 1 To hmeres ' DGV.Columns.Count - 1
        '        If cc = xl.Cells(1, i).ToString Then 'DGV.Columns(i).HeaderCell.Value Then
        '            sthlh = i
        '            Exit For
        '        End If
        '    Next
        '    If seira > 0 And sthlh > 0 Then
        '        If HTR.Rows(K)("idpel") > 0 Then
        '            'DGV.Rows(seira).Cells(sthlh).Value
        '            xl.Cells(seira, sthlh) = HTR.Rows(K)("EPO") + "_                               " + Str(HTR.Rows(K)("id")) ' HTR.Rows(K)("idpel")
        '            xl.Cells(sthlh).Style.BackColor = Color.LightGreen

        '        Else
        '            ' If xl.Cells(seira, sthlh).Style.BackColor = Color.LightGreen Then
        '            ' test = 1
        '            ' Else
        '            xl.Cells(seira, sthlh) = "_ " '+ Str(HTR.Rows(K)("id")) ' HTR.Rows(K)("IDPEL")
        '            xl.Cells(seira, sthlh).Style.BackColor = Color.Red
        '            ' End If

        '        End If
        '    End If


        'Next



    End Sub
    Private Sub DataGridView1_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = 3 Then
            MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)
        End If
    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        DataGridView1.Rows(1).Cells(2).Value = "SSS"
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        paint_grid()
    End Sub
    Private Sub paint_grid()
        Dim test As Integer
        Dim HTR As New DataTable
        ' ΒΑΖΩ ΕΠΙΚΕΦΑΛΙΔΕΣ
        ExecuteSQLQuery("select MAX(DATECHECKIN) AS MAX1,MIN(DATECHECKIN) AS MIN1 from HOTROOMDAYS  ", HTR)
        Dim mApo As Date = HTR.Rows(0)("MIN1")
        Dim mEos As Date = HTR.Rows(0)("MAX1")
        Dim hmeres As Integer = DateDiff("d", mApo, mEos) + 1
        Dim hmera As Date = mApo
        'If DGV.Columns.Count < hmeres Then
        '    For l As Integer = 1 To DGV.Columns.Count - hmeres
        '        DGV.Columns.Add("aaa", "--")


        '    Next
        'End If
        DGV.RowHeadersVisible = False
        DGV.Columns(0).HeaderCell.Value = "Ξεν"
        DGV.Columns(1).HeaderCell.Value = "Δωμ"
        Dim pl As Integer = 80
        ' ΔΕΙΧΝΩ ΤΙΣ ΜΕΡΕΣ
        For nn As Integer = 2 To hmeres + 1
            DGV.Columns(nn).Width = pl
            'DGV.Rows(0).Cells(nn).Value = Format(hmera, "dd/MM")
            DGV.Columns(nn).HeaderCell.Value = Format(hmera, "dd/MM")
            hmera = DateAdd("d", 1, hmera)

        Next

        DGV.RowHeadersVisible = True


        ' DEIXNΩ ΔΩΜΑΤΙΑ & ΞΕΝΟΔΟΧΕΙΑ
        ExecuteSQLQuery("select H.NAME,R.ROOMN from HOTROOMS R inner join HOTELS H ON H.ID=R.HOTELID ORDER BY H.NAME,R.ROOMN ", HTR)

        For K As Integer = 0 To HTR.Rows.Count - 1
            DGV.Rows.Add()

            DGV.Rows(K).Cells(0).Value = HTR.Rows(K)("name")
            DGV.Rows(K).Cells(1).Value = HTR.Rows(K)("roomn")

        Next


        ExecuteSQLQuery("select HOTELNAME,ROOMN,DATECHECKIN,ISNULL(IDPEL,0) AS IDPEL,ISNULL((SELECT EPO+'+'+ISNULL(SYNODOS,'') FROM PEL WHERE ID=IDPEL),'-') AS EPO ,ID from HOTROOMDAYS  ORDER BY HOTELNAME,ROOMN ", HTR)
        Dim X As String, R As String
        For K As Integer = 0 To HTR.Rows.Count - 1
            'ΠΡΟΣΔΙΟΡΙΖΩ ΤΗΝ ΣΕΙΡΑ
            X = Trim(HTR.Rows(K)("HOTELname"))
            R = Trim(HTR.Rows(K)("ROOMN")) ' ariumos domatioy p.x. 102
            Dim seira As Integer = 0
            For i As Integer = 0 To DGV.Rows.Count - 1
                'αν βρηκα το ξενοδοχείο(Χ) και το δωμάτιο(R) τοτε τσιμπα  την σειρά
                If X = DGV.Rows(i).Cells(0).Value And R = DGV.Rows(i).Cells(1).Value Then
                    seira = i
                    Exit For
                End If
            Next

            'ΠΡΟΣΔΙΟΡΙΖΩ ΤΗΝ ΣΤΗΛΗ  sygkrinontas to HOTROOMDAYS DATECHECKIN ME TON TITLO THS STHLHS KAI TSIMPAO THN STHLH
            Dim cc As String = Format(HTR.Rows(K)("datecheckin"), "dd/MM")
            Dim sthlh As Integer = 0
            For i = 0 To DGV.Columns.Count - 1
                If cc = DGV.Columns(i).HeaderCell.Value Then
                    sthlh = i
                    Exit For
                End If
            Next

            If HTR.Rows(K)("idpel") > 0 Then
                DGV.Rows(seira).Cells(sthlh).Value = HTR.Rows(K)("EPO") + "_                               " + Str(HTR.Rows(K)("id")) ' HTR.Rows(K)("idpel")
                DGV.Rows(seira).Cells(sthlh).Style.BackColor = Color.LightGreen

            Else
                If DGV.Rows(seira).Cells(sthlh).Style.BackColor = Color.LightGreen Then
                    test = 1
                Else
                    DGV.Rows(seira).Cells(sthlh).Value = "_                        " + Str(HTR.Rows(K)("id")) ' HTR.Rows(K)("IDPEL")
                    DGV.Rows(seira).Cells(sthlh).Style.BackColor = Color.Red
                End If

            End If



        Next
        DGV.Refresh()


    End Sub


    Private Sub Krarhseis_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Krarhseis.Click

        If UCase(username) = "ADMIN" Then
        Else
            Exit Sub

        End If



        '------------------------------------------  bookings ----------------------------------------------------------------13032 FESTIV
        Dim PEL As New DataTable
        Dim mRankPEL As String

        ExecuteSQLQuery("UPDATE PEL SET ENERGOS=0 WHERE ENERGOS IS NULL")



        Dim TEST As Integer
        'SELECT EPO as [ΦΙΛΟΞΕΝΟΥΜΕΝΟΙ ΧΩΡΙΣ ΚΡΑΤΗΣΗ],CHECKIN,CHECKOUT FROM PEL P LEFT JOIN HOTROOMDAYS H ON P.ID=H.IDPEL WHERE IDROOM IS NULL
        ' ΦΟΡΤΩΝΩ ΤΟΥΣ ΠΕΛΑΤΕΣ
        Dim MDAY As String, dcin As Date, DCOUT As Date
        ExecuteSQLQuery("select  convert(date,CHECKOUT) as CHECKOUTD,convert(date,CHECKIN) as CHECKIND,* from PEL  P LEFT JOIN HOTROOMDAYS H ON P.ID=H.IDPEL WHERE IDROOM IS NULL AND ENERGOS<>1", PEL)
        ListBox1.Items.Add("ΑΠΕΤΥΧΑΝ ΝΑ ΚΑΝΟΥΝ ΚΡΑΡΗΣΗ:")
        Dim ISOK As Boolean = True

        'ολοι οι πελάτες που δεν έχουν κάνει checkin
        For K As Integer = 0 To PEL.Rows.Count - 1
            ISOK = True
            If IsDBNull(PEL.Rows(K)("CHECKIN")) Or IsDBNull(PEL.Rows(K)("CHECKOUT")) Then
                MsgBox("ΔΕΝ ΕΧΕΙ ΗΜΕΡΟΜΗΝΙΕΣ CHECKIN-CHECOUT O/H" + PEL.Rows(K)("EPO"))
                ISOK = False
            End If
            If IsDBNull(PEL.Rows(K)("rank")) Then
                MsgBox("ΔΕΝ ΕΧΕΙ RANK O/H" + PEL.Rows(K)("EPO"))
                ISOK = False
            End If
            mRankPEL = PEL.Rows(K)("rank").ToString


            If ISOK Then


                MDAY = Format(PEL.Rows(K)("CHECKIN"), "dd/MM/yyyy") ' βρηκα την ημερα checkin
                dcin = PEL.Rows(K)("CHECKIN") ' βρηκα την ημερα checkin  PROSOXH EXEI KAI TIME
                DCOUT = PEL.Rows(K)("CHECKOUT")
                Dim DCIND As Date = PEL.Rows(K)("CHECKIND")
                Dim DCOUTD As Date = PEL.Rows(K)("CHECKOUTD")
                'Dim DCINDAYONLY As String = Format(PEL.Rows(K)("CHECKIN"), "MM/dd/yyyy")
                'Dim dcin2d As DateTime = Convert.ToDateTime(DCINDAYONLY)

                'Dim DCoutDAYONLY As String = Format(PEL.Rows(K)("CHECKout"), "MM/dd/yyyy")
                'Dim dcout2d As DateTime = Convert.ToDateTime(DCoutDAYONLY)



                Dim hmeres As Integer = DateDiff("d", DCIND, DCOUTD)
                ' α τροπος κρατησεις με database    ( b me pinaka datagridview)
                Dim HTR As New DataTable
                'ΒΡΙΣΚΩ ΤΑ ΚΕΝΑ ΣΕ ΑΥΤΟ ΤΟ ΔΙΑΣΤΗΜΑ

                'ΕΛΕΥΘΕΡΕΣ ΟΛΕΣ ΟΙ ΗΜΕΡΕΣ

                Dim A As String = PEL.Rows(K)("EPO").ToString()



                'Dim Sql As String = "DATECHECKIN>=" + MDAY + "' AND DATECKECKOUT AND (IDPEL IS NULL OR IDPEL=0) ORDER BY RANK "
                'ExecuteSQLQuery("select  DATECHECKIN,IDPEL,D.ID AS ID,HOTELID,IDPEL from HOTROOMDAYS D INNER JOIN HOTELS H ON D.HOTELID=H.ID  WHERE " + Sql, HTR)

                'HTR ----> ΑΚΡΙΒΩΣ ΓΙΑ ΤΗΝ ΗΜΕΡΑ ΤΟΥ CHECKIN ΒΛΕΠΩ ΤΑ ΔΙΑΘΕΣΙΜΑ ΔΩΜΑΤΙΑ
                ExecuteSQLQuery("select  DATECHECKIN,IDPEL,D.ID AS ID,HOTELID,IDPEL,IDROOM,H.NAME,D.ROOMN AS ROOMN,H.MAXNIGHTS from HOTROOMDAYS D INNER JOIN HOTELS H ON D.HOTELID=H.ID  WHERE  H.RANK<=" + mRankPEL + " AND CONVERT(CHAR(10),DATECHECKIN,103)='" + MDAY + "' AND (IDPEL IS NULL OR IDPEL=0) ORDER BY H.RANK desc ", HTR)
                For L As Integer = 0 To HTR.Rows.Count - 1 ' ΟΛΑ ΤΑ ΔΙΑΘΕΣΙΜΑ

                    If PEL.Rows(K)("ID") = 13032 Then
                        TEST = 1
                    End If




                    ' ΠΡΕΠΕΙ ΝΑ ΕΛΕΓΞΩ ΤΙΣ ΜΕΡΕΣ ΔΙΑΜΟΝΗΣ ΤΟΥ ΠΡΟΣΚΕΚΛΗΜΕΝΟΥ ΑΝ ΕΙΝΑΙ ΔΙΑΘΕΣΙΜΕΣ ΣΤΟ ΙΔΙΟ ΔΩΜΑΤΙΟ IDROOM
                    Dim HRDAYS As New DataTable
                    '
                    Dim OK As Integer = 0
                    Try
                        Dim KATEIL_nyxtes As New DataTable
                        ExecuteSQLQuery("select count(*) from HOTROOMDAYS WHERE isnull(IDPEL,0)>0 and HOTELID=" + HTR(L)("HOTELID").ToString(), KATEIL_nyxtes)


                        ' πρέπει οι  αγορασμένες νύχτες να ξεπερνουν το άθροισμα ( κατειλημένων ηδη ημερών + ημέρες του νέου προσκεκλημένου)
                        If HTR(L)("MAXNIGHTS") > KATEIL_nyxtes(0)(0) + hmeres Then

                            ExecuteSQLQuery("select count(*) from HOTROOMDAYS WHERE DATECHECKIN>='" + Format(DCIND, "MM/dd/yyyy") + "' AND DATECHECKIN<'" + Format(DCOUTD, "MM/dd/yyyy") + "' AND IDROOM=" + HTR(L)("IDROOM").ToString + " AND IDPEL=0", HRDAYS)
                            If HRDAYS(0)(0) = hmeres Then  ' εχει διαθεσιμες ολες τις ημερες οποτε οκ
                                ExecuteSQLQuery("UPDATE PEL SET CH2='" + HTR.Rows(0)("NAME") + "',CH1=" + HTR.Rows(0)("ROOMN") + ",HOTELID=" + HTR.Rows(0)("HOTELID").ToString + " WHERE ID=" + PEL(K)("ID").ToString)
                                ExecuteSQLQuery("update HOTROOMDAYS set IDPEL=" + PEL(K)("ID").ToString + " WHERE DATECHECKIN>='" + Format(DCIND, "MM/dd/yyyy") + "' AND DATECHECKIN<'" + Format(DCOUTD, "MM/dd/yyyy") + "' AND IDROOM=" + HTR(L)("IDROOM").ToString)
                                OK = 1
                                Exit For
                            End If

                        End If

                    Catch ex As Exception

                    End Try
                    If OK = 0 Then
                        ListBox1.Items.Add(PEL(0)("EPO"))
                    End If


                    'ExecuteSQLQuery("UPDATE HOTROOMDAYS SET IDPEL=" + PEL(K)("ID").ToString + " WHERE ID=" + HTR(0)("ID").ToString)

                Next



            End If

        Next
        PAINTGRID1()
        'UPDATE HOTROOMDAYS SET IDPEL=0
        MsgBox("OK")
        paint_grid()
    End Sub



    Private Sub DGV_CellContentClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DGV.CellContentClick
        'Dim R, C As Integer
        'R = e.RowIndex
        'C = e.ColumnIndex

    End Sub

    Private Sub DGV_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DGV.CellClick
        '-------------------------------------1o CLICK----------------------------------------------------------------------------

        If UCase(username) = "ADMIN" Then
        Else
            Exit Sub

        End If



        Dim R, C As Integer
        R = e.RowIndex
        C = e.ColumnIndex
        If R < 0 Or C < 1 Then Exit Sub



        '----------- 2o KLIK ----------------------------------------
        If F_REM_DAYS > 0 Then '  DGV.Rows(R).Cells(C).Style.BackColor = Color.YellowGreen Then  '-------------- 2o click -----------------------------------


            '---------------  τσιμπαω το id του 2ου κλικ ------------------------------------------------
            Dim d As String = DGV.Rows(R).Cells(C).Value()
            Dim s As String = ""
            f2_row = R : f2_col = C

            If f_CanMoveLeft = 0 Then
                If f2_col < f1_col Then
                    MsgBox("δεν μπορεί να ειναι σε 2 δωμάτια την ιδια μέρα")
                    Exit Sub
                End If

            End If


            Try
                s = d.Split("_")(0)
            Catch ex As Exception

            End Try
            If Len(s) > 0 Then
                MsgBox("Kateilhm;eno Δωμάτιο")
                Exit Sub

                ' f_2idHotRoomDays = s  ' βρίσκω το id ΤΟΥ 2oy ΚΛΙΚ

                ' DGV.Rows(R).Cells(C).Style.BackColor = Color.YellowGreen
            Else
                ' F_REM_DAYS = 0
                'f_idHotRoomDays = 0
            End If
            '----------------------------------------------  βγαζω το popup menu ---------------------------------------
            Dim currentCell As DataGridViewCell = DGV.CurrentCell

            Dim cellDisplayRect As Rectangle = DGV.GetCellDisplayRectangle(currentCell.ColumnIndex, currentCell.RowIndex, False)

            Dim cellAbsolutePos As Point = DGV.PointToScreen(cellDisplayRect.Location)
            Dim X, Y As Long : X = cellAbsolutePos.X : Y = cellAbsolutePos.Y
            ContextMenuStrip1.Show(DGV, New Point(IIf(X - 100 > 0, X - 100, 0), IIf(Y - 20 > 0, Y - 20, 0))) '
            ' F_REM_DAYS = 0
            Dim n As Integer
            n = 0
        Else '-------------------------------------------------- 1o click ---------------------------------------------------- 
            Dim d As String = DGV.Rows(R).Cells(C).Value()
            Dim s As String = ""
            f1_row = R : f1_col = C
            If d = Nothing Then
                Exit Sub
            End If

            If Len(Trim(d.Split("_")(0))) = 0 Then
                Exit Sub
            End If
            Try
                s = d.Split("_")(1)
            Catch ex As Exception

            End Try
            If Len(s) > 0 Then
                f_idHotRoomDays = s  ' βρίσκω το id ΤΟΥ ΑΡΧΙΚΟΥ ΚΛΙΚ
                F_REM_DAYS = GRIDfind_right_days(s, f1_row, f1_col) ' Pelfind_right_days(s) 'ΑΠΟΜΕΝΟΥΣΕΣ ΔΙΑΝΥΚΤΕΡΕΥΣΕΙΣ

                f_idpel = Str(GETn_VALUE("SELECT IDPEL FROM HOTROOMDAYS WHERE ID=" + f_idHotRoomDays))
                If Val(f_idpel) = 0 Then
                    F_REM_DAYS = 0
                End If
            Else
                F_REM_DAYS = 0
                f_idHotRoomDays = 0
                f_idpel = "0"
            End If
            If F_REM_DAYS > 0 Then
                DGV.Rows(R).Cells(C).Style.BackColor = Color.YellowGreen
            End If


            'ΒΡΙΣΚΩ ΤΟ ΑΡΙΣΤΕΡΟ ΚΟΥΤΑΚΙ ΑΠΟ ΑΥΤΟ ΠΟΥ ΕΚΑΝΑ ΚΛΙΚ ΜΗΝ ΤΥΧΟΝ ΑΝΗΚΕΙ ΣΤΟΝ ΙΔΙΟ ΠΕΛΑΤΗ
            d = DGV.Rows(R).Cells(C - 1).Value()
            If d = Nothing Then
                '  Exit Sub
            End If


            If C - 1 < 2 Then ' einai akrh
                f_CanMoveLeft = 0
            Else
                s = d.Split("_")(1)
                Dim aristID As String = Str(GETn_VALUE("SELECT IDPEL FROM HOTROOMDAYS WHERE ID=" + s))
                ' αν προκειται να κοψω στα 2 την κράτηση να μην μπορει να παει αριστερότερα
                ' γιατι την ιδια μερα θα κοιμαται σε 2 ξενοδοχεια
                If Val(aristID) = Val(f_idpel) Then
                    f_CanMoveLeft = 0
                Else
                    f_CanMoveLeft = 1

                End If
            End If









        End If



    End Sub
    Private Function GRIDfind_right_days(ByVal IDPEL, ByVal f1_row, ByVal f1_col)
        Dim d As String
        For K As Integer = f1_col To DGV.Columns.Count - 1
            d = DGV.Rows(f1_row).Cells(K).Value()

            If DGV.Rows(f1_row).Cells(K).Value() = Nothing Then
                GRIDfind_right_days = K - f1_col
                Exit For
            Else

                If GETn_VALUE("SELECT IDPEL FROM HOTROOMDAYS WHERE ID=" + IDPEL) = GETn_VALUE("SELECT IDPEL FROM HOTROOMDAYS WHERE ID=" + d.Split("_")(1)) Then
                    'OK
                Else
                    GRIDfind_right_days = K - f1_col
                    Exit For
                End If
            End If

        Next

    End Function

    Private Sub DGV_CellMouseUp(ByVal sender As Object, ByVal e As DataGridViewCellMouseEventArgs) Handles DGV.CellMouseUp
        ''-----------------  MENU ------------------------------------------------------
        If e.Button = MouseButtons.Right Then
            If My.Computer.Keyboard.ShiftKeyDown Then
                delete_day
                Exit Sub
            End If
            F_REM_DAYS = 0



            Dim currentCell As DataGridViewCell = DGV.CurrentCell
            If InStr(currentCell.Value.ToString, "_") = 0 Then
                Exit Sub
            End If
            Dim ID As String = currentCell.Value.ToString.Split("_")(1)

            Dim HT2 As New DataTable
            ExecuteSQLQuery("select * from HOTROOMDAYS WHERE ID=" + ID, HT2)
            Dim nextID As String = (Val(ID) + 1).ToString
            Dim HOTELID As Long = HT2(0)("HOTELID")
            Dim IDPEL As Long = HT2(0)("IDPEL")
            Dim datecheckin As Date = HT2(0)("DATECHECKIN")

            Dim CHE As String = datecheckin.ToString


            Dim HT3 As New DataTable
            ExecuteSQLQuery("select * from HOTROOMDAYS WHERE ID=" + nextID, HT3)


            If HOTELID = HT3(0)("HOTELID") Then
                'OK
            Else
                MsgBox("ΑΛΛΑΓΗ ΞΕΝΟΔΟΧΕΙΟΥ.ΔΕΝ ΕΠΙΤΡΕΠΕΤΑΙ Η ΠΡΟΣΘΕΣΗ ΔΙΑΝΥΚΤΕΡΕΥΣΗΣ")
                Exit Sub
            End If

            If HT2(0)("ROOMN") = HT3(0)("ROOMN") Then
                'OK
            Else
                MsgBox("ΑΛΛΑΓΗ ΔΩMATIOY.ΔΕΝ ΕΠΙΤΡΕΠΕΤΑΙ Η ΠΡΟΣΘΕΣΗ ΔΙΑΝΥΚΤΕΡΕΥΣΗΣ")
                Exit Sub
            End If

            ' Η ΕΠΟΜΕΝΗ ΜΕΡΑ ΣΤΟ ΙΔΙΟ ΔΩΜΑΤΙΟ ΕΙΝΑΙ ΚΕΝΗ
            If HT3(0)("IDPEL") = 0 Then
                'OK
            Else
                'ΕΛΕΓΧΩ ΜΗΠΩΣ ΘΕΛΕΙ ΝΑ ΠΡΟΣΘΕΣΩ ΠΡΙΝ ΤΗΝ 1Η ΗΜΕΡΑ
                Dim HTPREV As New DataTable
                ExecuteSQLQuery("select * from HOTROOMDAYS WHERE ID=" + (Val(ID) - 1).ToString, HTPREV)
                If HTPREV(0)("IDPEL") = 0 Then  ' ΑΡΙΣΤΕΡΑ ΕΧΕΙ ΚΕΝΟ ΠΟΥ ΕΙΝΑΙ ΑΔΕΙΟ
                    Dim ans2 As Integer = MsgBox("Να προστεθεί μία διανυκτέρευση;", vbYesNo)
                    If ans2 = vbYes Then
                        ' ΕΛΕΓΧΩ ΑΝ Η ΠΡΩΤΗ ΜΕΡΑ ΤΟΥ HOTROOMDAYS EINAI IDIA ME TO CHECKIN TOY ΠΡΟΣΚΕΚΛΗΜΕΝΟΥ
                        ' ΜΗΝ ΤΥΧΟΝ ΔΗΛΑΔΗ ΕΧΕΙ ΣΠΑΣΙΜΟ ΣΕ 2 ΞΕΝΟΔΟΧΕΙΑ
                        ' ΚΑΙ ΠΑΩ ΣΤΟ 2ο ΣΠΑΣΙΜΟ ΚΑΙ ΠΡΟΣΘΕΣΩ ΜΙΑ ΜΕΡΑ ΑΡΙΣΤΕΡΑ 
                        ' ΜΕ ΑΠΟΤΕΛΕΣΜΑ Η ΙΔΙΑ ΜΕΡΑ ΝΑ ΕΙΝΑΙ 2 ΦΟΡΕΣ ΣΤΟΝ ΠΡΟΣΚΕΚΛΗΜΕΝΟ

                        Dim HT4 As New DataTable
                        ExecuteSQLQuery("SELECT CHECKIN FROM PEL  WHERE ID=" + IDPEL.ToString, HT4)
                        If HT4(0)(0).DAY = datecheckin.Day Then

                            ExecuteSQLQuery("UPDATE HOTROOMDAYS SET IDPEL=" + IDPEL.ToString + "  WHERE ID=" + (Val(ID) - 1).ToString, HT4)
                            ExecuteSQLQuery("UPDATE PEL SET CHECKIN=DATEADD(day,-1,CHECKIN)      WHERE ID=" + IDPEL.ToString)

                            paint_grid()
                            Exit Sub
                        Else
                            MsgBox("EXEI ΔΙΑΜΟΝΗ ΣΕ 2 ΞΕΝΟΔΟΧΕΙΑ/ΔΩΜΑΤΙΑ. ΔΕΝ ΕΠΙΤΡΕΠΕΤΑΙ Η ΠΡΟΣΘΕΣΗ ΔΙΑΝΥΚΤΕΡΕΥΣΗΣ")
                            Exit Sub
                        End If

                    Else
                        'MsgBox("ΕΠΟΜΕΝΗ ΗΜΕΡΑ ΚΑΤΕΙΛΗΜΕΝΗ. ΔΕΝ ΕΠΙΤΡΕΠΕΤΑΙ Η ΠΡΟΣΘΕΣΗ ΔΙΑΝΥΚΤΕΡΕΥΣΗΣ")
                        Exit Sub

                    End If


                End If  'If HTPREV(0)("IDPEL") = 0 Then  ' ΑΡΙΣΤΕΡΑ ΕΧΕΙ ΚΕΝΟ ΠΟΥ ΕΙΝΑΙ ΑΔΕΙΟ
            End If ' HT3(0)("IDPEL") = 0 Then

            Dim ans As Integer = MsgBox("Να προστεθεί μία διανυκτέρευση;", vbYesNo)
            If ans = vbYes Then
                Dim HT4 As New DataTable
                ExecuteSQLQuery("SELECT CHECKOUT FROM PEL  WHERE ID=" + IDPEL.ToString, HT4)
                Dim CCHE As String = HT4(0)(0).ToString

                If HT4(0)(0).DAY = datecheckin.Day + 1 Then
                    ExecuteSQLQuery("UPDATE HOTROOMDAYS SET IDPEL=" + IDPEL.ToString + "  WHERE ID=" + nextID.ToString, HT4)
                    ExecuteSQLQuery("UPDATE PEL SET CHECKOUT=DATEADD(day,1,CHECKOUT)      WHERE ID=" + IDPEL.ToString)
                    paint_grid()
                Else
                    MsgBox("EXEI ΔΙΑΜΟΝΗ ΣΕ 2 ΞΕΝΟΔΟΧΕΙΑ/ΔΩΜΑΤΙΑ. ΔΕΝ ΕΠΙΤΡΕΠΕΤΑΙ Η ΠΡΟΣΘΕΣΗ ΔΙΑΝΥΚΤΕΡΕΥΣΗΣ")
                    Exit Sub
                End If

            End If


            '    Dim cellDisplayRect As Rectangle = DGV.GetCellDisplayRectangle(currentCell.ColumnIndex, currentCell.RowIndex, False)

            '    Dim cellAbsolutePos As Point = DGV.PointToScreen(cellDisplayRect.Location)
            '    Dim X, Y As Long : X = cellAbsolutePos.X : Y = cellAbsolutePos.Y
            '    ContextMenuStrip1.Show(DGV, New Point(IIf(X - 200 > 0, X - 200, 0), IIf(Y - 300 > 0, Y - 300, 0))) ' Button1.Height))

        End If  'If e.Button = MouseButtons.Right Then


    End Sub
    Sub Delete_day()
        '================================
        Dim currentCell As DataGridViewCell = DGV.CurrentCell
        If InStr(currentCell.Value.ToString, "_") = 0 Then  'adeio
            Exit Sub
        End If



        Dim ID As String = currentCell.Value.ToString.Split("_")(1)

        Dim HT2 As New DataTable
        ExecuteSQLQuery("select * from HOTROOMDAYS WHERE ID=" + ID, HT2)
        Dim nextID As String = (Val(ID) + 1).ToString
        Dim HOTELID As Long = HT2(0)("HOTELID")
        Dim IDPEL As Long = HT2(0)("IDPEL")
        Dim datecheckin As Date = HT2(0)("DATECHECKIN")

        Dim CHE As String = datecheckin.ToString
        If IDPEL = 0 Then
            Exit Sub

        End If


        Dim ans As Integer = MsgBox("Να αφαιρεθεί μία διανυκτέρευση;", vbYesNo)
        If ans = vbYes Then
            Dim HT4 As New DataTable
            ExecuteSQLQuery("SELECT CHECKOUT,CHECKIN FROM PEL  WHERE ID=" + IDPEL.ToString, HT4)
            If HT4.Rows.Count = 0 Then
                Exit Sub
            End If
            Dim CCHE As String = HT4(0)(0).ToString

            If HT4(0)("CHECKOUT").DAY - 1 = datecheckin.Day Then
                ExecuteSQLQuery("UPDATE HOTROOMDAYS SET IDPEL=0  WHERE ID=" + ID.ToString, HT4)
                ExecuteSQLQuery("UPDATE PEL SET CHECKOUT=DATEADD(day,-1,CHECKOUT)      WHERE ID=" + IDPEL.ToString)
                paint_grid()
                Exit Sub

            End If
            If HT4(0)("CHECKIN").DAY = datecheckin.Day Then
                ExecuteSQLQuery("UPDATE HOTROOMDAYS SET IDPEL=0  WHERE ID=" + ID.ToString, HT4)
                ExecuteSQLQuery("UPDATE PEL SET CHECKIN=DATEADD(day,1,CHECKIN)      WHERE ID=" + IDPEL.ToString)
                paint_grid()
                Exit Sub

            End If



            MsgBox("δεν έγινε η αφαίρεση")



        End If


    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If UCase(username) = "ADMIN" Then
        Else
            Exit Sub

        End If
        Dim ANS As Integer
        ANS = MsgBox("ΠΡΟΣΟΧΗ ΘΑ ΔΙΑΓΡΑΦΟΥΝ ΟΙ ΚΡΑΤΗΣΕΙΣ", MsgBoxStyle.YesNo)
        If ANS = vbYes Then

            ExecuteSQLQuery("UPDATE HOTROOMDAYS SET IDPEL=0")
            ExecuteSQLQuery("UPDATE PEL SET CH2='',CH1='',NUM2=0 ")
            ExecuteSQLQuery("update PEL SET [HOTELID]=0;")


            ' ExecuteSQLQuery("update PEL SET [CHECKIN]='05/01/2023',[CHECKOUT]='05/13/2023',[HOTELID]=0;")

            ' ExecuteSQLQuery("update HOTROOMS SET [APO]='05/01/2023',[EOS]='05/13/2023'")
        End If






    End Sub
    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        '---------------------------------------------move booking------------------------------------------------------------------
        'τα 2 IDs f_2idHotRoomDays , f_idHotRoomDays
        ' F_REM_DAYS ποσες μερες δεξια του τσιμπάω
        ' f1_row = R : f1_col = C  1o klik
        ' ελεγχω πρωτον ότι δεξια του νέου μέρους εχω τις απαραίτητεσ ημερες
        '
        Dim nc, n As Integer
        nc = 0
        For n = f2_col To f2_col + F_REM_DAYS - 1
            If DGV.Rows(f2_row).Cells(n).Style.BackColor = Color.Red Then
                nc = nc + 1
            End If
        Next


        If nc = F_REM_DAYS Then ' εχω τος απαραιτητεσ μερες
            ' γεμιζω το νεο  antigrafontas ta palia kelia
            Dim nc1 As Integer = f1_col  ' κολονα παλιου που ισως να μην συμπιπτει με την κολανα του νεου (αν λειξει ενδιαμεσα ο πλελατης)
            Dim SMAX, SMIN As String ' Η ΠΡΩΤΗ ΚΑΙ Η ΤΕΛΕΥΤΑΙΑ ΗΜΕΡΟΜΗΝΙΑ 
            For n = f2_col To f2_col + F_REM_DAYS - 1
                Dim DPALIO As String = DGV.Rows(f2_row).Cells(n).Value.ToString
                DGV.Rows(f2_row).Cells(n).Value = DGV.Rows(f1_row).Cells(nc1).Value
                DGV.Rows(f1_row).Cells(n).Style.BackColor = Color.LightGreen
                Dim d As String = DGV.Rows(f1_row).Cells(nc1).Value()
                Dim s As String = ""
                s = DPALIO.Split("_")(1)
                If n = f2_col Then SMAX = s
                If n = f2_col + F_REM_DAYS - 1 Then SMIN = s
                ExecuteSQLQuery("update HOTROOMDAYS set IDPEL=" + f_idpel + " WHERE ID= " + s) ' AND DATECHECKIN<'" + Format(DCOUTD, "MM/dd/yyyy") + "' AND IDROOM=" + HTR(L)("IDROOM").ToString)

                nc1 = nc1 + 1
                'End If
            Next
            'ΕΝΗΜΕΡΩΝΩ ΤΟΝ ΠΕΛΑΤΗ ΜΕ ΤΙς ΝΕΕς ΗΜΕΡΟΜΗΝΙΕΣ
            If f2_col <> f1_col Then
                ExecuteSQLQuery("UPDATE  PEL SET CHECKOUT=DATEADD(DAY," + Str(f2_col - f1_col) + ",CHECKOUT) WHERE ID=" + f_idpel)
                ExecuteSQLQuery("UPDATE  PEL SET CHECKIN=DATEADD(DAY," + Str(f2_col - f1_col) + ",CHECKIN) WHERE ID=" + f_idpel)
            End If
            'ΕΝΗΜΕΡΩΝΩ ΤΟΝ ΠΕΛΑΤΗ ΜΕ ΤΟ ΝΕΟ ΞΕΝΟΔΟΧΕΙΟ
            ExecuteSQLQuery("UPDATE  PEL SET CH2='" + Mid(DGV.Rows(f2_row).Cells(0).Value.ToString, 1, 30) + "' WHERE ID=" + f_idpel)

            ' ελευθερωνω το παλιο
            n = f1_col
            For n = f1_col To f1_col + F_REM_DAYS - 1
                Dim DD As String = DGV.Rows(f1_row).Cells(n).Value.ToString
                DGV.Rows(f1_row).Cells(n).Value = "_" + DD.Split("_")(1)
                DGV.Rows(f1_row).Cells(n).Style.BackColor = Color.Red


                Dim s As String = DD.Split("_")(1)
                ExecuteSQLQuery("update HOTROOMDAYS set IDPEL=0 WHERE ID= " + s) ' AND DATECHECKIN<'" + Format(DCOUTD, "MM/dd/yyyy") + "' AND IDROOM=" + HTR(L)("IDROOM").ToString)

                ' nc = nc + 1
                'End If
            Next
            DGV.Refresh()

            ' ενημερωνω την database me tis αλλαγες που εγιναν αφου γίνει επιβεβαίωση
            'ExecuteSQLQuery("select count(*) from HOTROOMDAYS WHERE DATECHECKIN>='" + Format(DCIND, "MM/dd/yyyy") + "' AND DATECHECKIN<'" + Format(DCOUTD, "MM/dd/yyyy") + "' AND IDROOM=" + HTR(L)("IDROOM").ToString, HRDAYS)
            'If HRDAYS(0)(0) = hmeres Then  ' εχει διαθεσιμες ολες τις ημερες οποτε οκ
            '    ExecuteSQLQuery("UPDATE PEL SET CH2='" + HTR.Rows(0)("NAME") + "',CH1=" + HTR.Rows(0)("ROOMN") + ",NUM2=" + HTR.Rows(0)("HOTELID").ToString + " WHERE ID=" + PEL(K)("ID").ToString)
            '    ExecuteSQLQuery("update HOTROOMDAYS set IDPEL=" + PEL(K)("ID").ToString + " WHERE DATECHECKIN>='" + Format(DCIND, "MM/dd/yyyy") + "' AND DATECHECKIN<'" + Format(DCOUTD, "MM/dd/yyyy") + "' AND IDROOM=" + HTR(L)("IDROOM").ToString)
            '    OK = 1
            '    Exit For
            'End If

        Else
            MsgBox(" δεν φτανουν οι μερες")



        End If
        F_REM_DAYS = 0
        paint_grid()

        'Dim HRDAYS As New DataTable
        ''
        'Dim OK As Integer = 0
        'Try

        '    '------------------- ΕΛΕΓΧΩ ΑΝ ΔΕΞΙΑ ΜΟΥ ΕΧΩ ΤΙΣ ΑΠΑΙΤΟΥΜΕΝΕΣ ΜΕΡΕΣ
        '    ExecuteSQLQuery("select count(*) from HOTROOMDAYS WHERE DATECHECKIN>='" + Format(DCIND, "MM/dd/yyyy") + "' AND DATECHECKIN<'" + Format(DCOUTD, "MM/dd/yyyy") + "' AND IDROOM=" + HTR(L)("IDROOM").ToString, HRDAYS)
        '    If HRDAYS(0)(0) = hmeres Then  ' εχει διαθεσιμες ολες τις ημερες οποτε οκ
        '        ExecuteSQLQuery("UPDATE PEL SET CH2='" + HTR.Rows(0)("NAME") + "',CH1=" + HTR.Rows(0)("ROOMN") + ",NUM2=" + HTR.Rows(0)("HOTELID").ToString + " WHERE ID=" + PEL(K)("ID").ToString)
        '        ExecuteSQLQuery("update HOTROOMDAYS set IDPEL=" + PEL(K)("ID").ToString + " WHERE DATECHECKIN>='" + Format(DCIND, "MM/dd/yyyy") + "' AND DATECHECKIN<'" + Format(DCOUTD, "MM/dd/yyyy") + "' AND IDROOM=" + HTR(L)("IDROOM").ToString)
        '        OK = 1
        '        Exit For
        '    End If
        'Catch ex As Exception

        'End Try





    End Sub
    Function Pelfind_right_days(ByVal id As String) As Integer
        ' βρισκει πόσες διανυκτερεύσεις μαζί με την τρέχουσα του απομένουν TOY PELATH POY EINAI SE AYTO TO KOYTAKI (ΑΠΟ ΤΟ ΤΡΕΧΟΝ ΚΟΥΤΑΚΙ ΜΕΧΡΙ ΤΟ ΤΕΛΟΣ ΤΗΣ ΔΙΑΜΟΝΗΣ
        Dim HTR As New DataTable

        ExecuteSQLQuery("select  H.*,convert(date,H.DATECHECKIN) AS CHECKIND,convert(date,PEL.CHECKOUT) as CHECKOUTD from HOTROOMDAYS H inner JOIN PEL ON H.IDPEL=PEL.ID WHERE H.ID=" + id, HTR)
        Dim hmeres As Integer = 0 ' DateDiff("d", HTR(0)("CHECKIND"), HTR(0)("CHECKOUTD"))
        If HTR.Rows.Count > 0 Then
            hmeres = DateDiff("d", HTR(0)("CHECKIND"), HTR(0)("CHECKOUTD"))
        End If

        Pelfind_right_days = hmeres

    End Function



    Function GETn_VALUE(ByVal QUERY As String) As Single
        Dim DT77 As New DataTable
        ExecuteSQLQuery(QUERY, DT77)
        If DT77.Rows.Count > 0 Then
            GETn_VALUE = IIf(IsDBNull(DT77.Rows(0)(0)), 0, DT77.Rows(0)(0))
        Else
            GETn_VALUE = -1
        End If
    End Function








    Private Sub test_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        paint_grid()
        PAINTGRID1()
        '  Dim Q As String = "SELECT EPO as [ΦΙΛΟΞΕΝΟΥΜΕΝΟΙ ΧΩΡΙΣ ΚΡΑΤΗΣΗ],CHECKIN,CHECKOUT FROM PEL WHERE HOTELID=0 OR HOTELID IS NULL"

        'Try
        '    Dim sqlCon As New OleDbConnection(gConnect)

        '    Dim sqlDA As New OleDbDataAdapter(Q, sqlCon)

        '    Dim sqlCB As New OleDbCommandBuilder(sqlDA)
        '    sqlDT.Reset() ' refresh 
        '    sqlDA.Fill(sqlDT, "PEL")


        'conn.Open()

        'da = New SqlDataAdapter(SQLqry, conn)

        ''create command builder
        'Dim cb As SqlCommandBuilder = New SqlCommandBuilder(da)
        'ds.Clear()
        ''fill dataset
        'da.Fill(ds, "PEL")
        'GridView1.ClearSelection()
        'GridView1.DataSource = ds
        'GridView1.DataMember = "PEL"

        ' cnString = gConSQL


        'Dim SQLqry
        '' SQLqry = Label1.Text '"SELECT NAME,N1,ID FROM ERGATES " ' ORDER BY HME "
        'Dim conn As SqlConnection = New SqlConnection(gConSQL)
        '' Try
        '' Open connection
        'conn.Open()

        'Dim da As SqlDataAdapter = New SqlDataAdapter(Q, conn)

        ''create command builder
        'Dim cb As SqlCommandBuilder = New SqlCommandBuilder(da)
        'Dim ds As DataSet
        'ds = New DataSet ' ds.Clear()
        ''fill dataset
        'da.Fill(ds, "PEL")
        'DataGridView1.ClearSelection()
        'DataGridView1.DataSource = ds
        'DataGridView1.DataMember = "PEL"
















        'DataGridView1.ClearSelection()
        'DataGridView1.DataSource = sqlDT
        'DataGridView1.Refresh()
        ''DataGridView1.DataMember = "PEL"
        ' Catch ex As Exception

        ' End Try
    End Sub




    Private Sub PAINTGRID1()
        Dim Q As String = "SELECT EPO as [ΦΙΛΟΞΕΝΟΥΜΕΝΟΙ ΧΩΡΙΣ ΚΡΑΤΗΣΗ],CHECKIN,CHECKOUT FROM PEL P LEFT JOIN HOTROOMDAYS H ON P.ID=H.IDPEL WHERE IDROOM IS NULL AND ENERGOS<>1" ' HOTELID=0 OR HOTELID IS NULL"
        Dim SQLqry
        ' SQLqry = Label1.Text '"SELECT NAME,N1,ID FROM ERGATES " ' ORDER BY HME "
        Dim conn As SqlConnection = New SqlConnection(gConSQL)
        ' Try
        ' Open connection
        conn.Open()

        Dim da As SqlDataAdapter = New SqlDataAdapter(Q, conn)

        'create command builder
        Dim cb As SqlCommandBuilder = New SqlCommandBuilder(da)
        Dim ds As DataSet
        ds = New DataSet ' ds.Clear()
        'fill dataset
        da.Fill(ds, "PEL")
        DataGridView1.ClearSelection()
        DataGridView1.DataSource = ds
        DataGridView1.DataMember = "PEL"

    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        ' ακυρωση κράτησης

        Dim ans As Integer
        ans = MsgBox("Να ακυρωθεί η κράτηση;", MsgBoxStyle.YesNoCancel)
        If ans = MsgBoxResult.Yes Then
            ExecuteSQLQuery("UPDATE  PEL SET CHECKOUT=null,HOTELID=0 WHERE ID=" + f_idpel)
            ExecuteSQLQuery("UPDATE  PEL SET CHECKIN=null WHERE ID=" + f_idpel)
            ExecuteSQLQuery("UPDATE  PEL SET CH2=null WHERE ID=" + f_idpel) 'ΞΕΝΟΔΟΧΕΙΟ =''
            ExecuteSQLQuery("update HOTROOMDAYS set IDPEL=0 where IDPEL=" + f_idpel)
            F_REM_DAYS = 0
            paint_grid()
            PAINTGRID1()






        End If





    End Sub

    Private Sub ContextMenuStrip1_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub

    Private Sub ΕξοδοςToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ΕξοδοςToolStripMenuItem.Click
        DGV.Rows(f1_row).Cells(f1_col).Style.BackColor = Color.LightGreen
        F_REM_DAYS = 0
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        FindInGrid(Me.DGV, bres.Text)
    End Sub
    Private Function FindInGrid(ByRef dgvGrid As DataGridView, ByVal strFind As String, Optional ByVal bStartAtBeggining As Boolean = False) As Point

        'call it like:
        '    ' forces a start at the grid cell 0,0
        'FindInGrid(Me.dgvMyGridView, Me.m_sFind, True)

        '    ' starts one cell beyond the staic member X, if X = last column then increments static member Y and sets X = 0
        'FindInGrid(Me.dgvMyGridView, Me.m_sFind)

        Dim bFound As Boolean = False
        Static pResult As Point
        Static X As Integer
        Static Y As Integer
        Dim dgvCell As DataGridViewCell
        Dim dgvRow As DataGridViewRow = Nothing
        For Each dgvRow In dgvGrid.SelectedRows()
            dgvRow.Selected = False
        Next
        If bStartAtBeggining Then
            pResult.X = 0
            pResult.Y = 0
        Else
            If pResult.X < dgvGrid.Columns.Count - 1 Then
                pResult.X += 1
            Else
                pResult.X = 0
                If pResult.Y < dgvGrid.Rows.Count - 1 Then
                    pResult.Y += 1
                Else
                    pResult.Y = 0
                End If
            End If
        End If
        For Y = pResult.Y To dgvGrid.Rows.Count - 1
            For X = pResult.X To dgvGrid.Columns.Count - 1
                dgvCell = dgvGrid(X, Y)
                If Not IsDBNull(dgvCell.Value) Then
                    If Not dgvCell.Value = Nothing Then
                        If dgvCell.Value.ToString.ToLower.Contains(strFind.ToLower) Then
                            If Not dgvCell.Value = Nothing Then
                                pResult.X = X
                                pResult.Y = Y
                                bFound = True
                                Exit For
                            End If
                        End If
                    End If
                End If


            Next
            If bFound Then
                Exit For
            Else
                pResult.X = 0
            End If
            If Y = dgvGrid.Rows.Count - 1 Then
                pResult.Y = 0
                If dgvGrid.CurrentCell.Value.ToString.ToLower.Contains(strFind.ToLower) Then
                    MessageBox.Show("δεν υπάρχει αλλο " & strFind & " για εύρεση.",
                    "Grid Search...", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                Else
                    MessageBox.Show("'" & strFind & "' δεν βρέθηκε ",
                    "Grid Search...", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
                End If
            End If
        Next
        If bFound Then
            Try
                dgvGrid.Item(pResult.X, pResult.Y).Selected = True
                dgvGrid.CurrentCell = dgvGrid.Item(pResult.X, pResult.Y)

            Catch ex As Exception

            End Try

        End If
    End Function

    Private Sub DGV_MouseMove(sender As Object, e As MouseEventArgs) Handles DGV.MouseMove

    End Sub
End Class