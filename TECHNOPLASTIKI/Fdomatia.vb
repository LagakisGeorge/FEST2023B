Public Class Fdomatia
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
            Return m_ID
        End Get
        Set(ByVal Value As Integer)
            ' If Value < 1 Or Value > 12 Then
            ' Error processing for invalid value. 
            'Else
            m_ID = Value
            'End If
        End Set
    End Property

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click




        If D1.Value > D2.Value Then
            MsgBox("ΗΜΕΡΟΜΗΝΊΑ CHECKOUT ΜΙΚΡΟΤΕΡΗ ΤΟΥ CHECKIN")
            Exit Sub

        End If

        If D1.Value < gHMEARX Or D1.Value > gHMETEL Then
            MsgBox("ΗΜΕΡΟΜΗΝΊΑ CHECKIN EKTOΣ ΔΙΑΣΤΗΜΑΤΟΣ  " + gHMEARX.ToString + "-" + gHMETEL.ToString)
            Exit Sub
        End If
        If D2.Value < gHMEARX Or D2.Value > gHMETEL Then
            MsgBox("ΗΜΕΡΟΜΗΝΊΑ CHECKOUT EKTOΣ ΔΙΑΣΤΗΜΑΤΟΣ  " + gHMEARX.ToString + "-" + gHMETEL.ToString)
            Exit Sub
        End If
        'If cat.Text.Length = 0 Then
        '    MsgBox("ΔΕΝ ΕΧΕΤΕ ΒΑΛΕΙ ΚΑΤΗΓΟΡΙΑ")
        '    Exit Sub
        'End If











        Dim SQL As String = "INSERT INTO HOTROOMS (ROOMN,HOTELID,KREBATIA,APO,EOS) VALUES(" + domatio.Text + "," + Str(ID) + "," + kreb.Text + ",'" + Format(D1.Value, "MM/dd/yyyy") + "','" + Format(D2.Value, "MM/dd/yyyy") + "')"
        ExecuteSQLQuery(Sql)
        PAINTROOMS()


        '        CREATE TABLE [dbo].[HOTROOMS](
        '	[ROOMN] [varchar](5) NOT NULL,
        '	[HOTELID] [int] NOT NULL,
        '	[DOMATIA] [int] NULL,
        '	[CATEGORY] [int] NULL,
        '	[APO] [datetime] NULL,
        '	[EOS] [datetime] NULL,
        '	[N1] [real] NULL,
        '	[N2] [real] NULL,
        '	[N3] [real] NULL,
        '	[C1] [nvarchar](50) NULL,
        '	[C2] [nvarchar](50) NULL,
        '	[C3] [nvarchar](50) NULL,
        '	[H1] [datetime] NULL,
        '	[H2] [datetime] NULL,
        '	[H3] [datetime] NULL,
        '	[B1] [bit] NULL,
        '	[B2] [bit] NULL,
        '	[ID] [int] IDENTITY(1,1) NOT NULL,
        ' CONSTRAINT [PK_HOTROOMS] PRIMARY KEY CLUSTERED 
        '(
        '	[ROOMN] ASC,
        '        [HOTELID](Asc)
        ')WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
        ') ON [PRIMARY]
        'Dim DT2 As New DataTable
        'DT2 = Execute2SQLQuery("select * from HOTROOMS WHERE HOTELID=" + Str(ID))
        'Dim N As Integer
        'For K As Integer = 0 To DT2.Rows.Count - 1

        '    ListView1.Items.Add(DT2.Rows(K)("ROOMN"))
        '    ListView1.Items(N).SubItems.Add(DT2.Rows(K)("APO").ToString)
        '    ListView1.Items(N).SubItems.Add(DT2.Rows(K)("EOS").ToString)
        '    ListView1.Items(N).SubItems.Add(DT2.Rows(K)("DOMATIA").ToString)

        '    N = N + 1

        'Next
        'ListView1.Items.Add("77")
        'ListView1.Items.Add("")
        'For k As Integer = 0 To ListView1.Items.Count - 1

        '    ListView1.Items(k).SubItems.Add("1lagakis")
        '    ListView1.Items(k).SubItems.Add("2Accounting")
        '    ListView1.Items(k).SubItems.Add("3John Smith")
        '    ListView1.Items(k).SubItems.Add("4Accounting")
        'Next

        
    End Sub

    Private Sub Fdomatia_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
     D1.Value = gHMEARX 
    D2.Value = gHMEARX 


        PAINTROOMS()

    End Sub

    Private Sub DIAGRAFI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DIAGRAFI.Click
        Try
            Dim a As String = ListView1.Items(ListView1.SelectedIndices(0)).SubItems(0).Text
            Dim n As Integer = ListView1.FocusedItem.Index
            ExecuteSQLQuery("delete from HOTROOMS WHERE ROOMN=" + a + "")
            PAINTROOMS()
        Catch
            MsgBox("Διαλέξτε δωμάτιο")
        End Try

    End Sub


    Private Sub PAINTROOMS()

        ListView1.Items.Clear()
        ListView1.Columns.Clear()
        ListView1.Columns.Add("Δωμάτιο", 100, HorizontalAlignment.Center) 'Column 1
        ListView1.Columns.Add("Από", 100, HorizontalAlignment.Center) 'Column 1
        ListView1.Columns.Add("Εως", 100, HorizontalAlignment.Center) 'Column 1
        ListView1.Columns.Add("Αριθ.Κλινών", 100, HorizontalAlignment.Center) 'Column 1
        ListView1.Columns.Add("Κατηγορία", 100, HorizontalAlignment.Center) 'Column 1
        Dim DT2 As New DataTable
        DT2 = Execute2SQLQuery("select * from HOTROOMS WHERE HOTELID=" + Str(ID))
        Dim N As Integer
        For K As Integer = 0 To DT2.Rows.Count - 1

            ListView1.Items.Add(DT2.Rows(K)("ROOMN"))
            ListView1.Items(N).SubItems.Add(Format(DT2.Rows(K)("APO"), "dd/MM/yyyy"))
            ListView1.Items(N).SubItems.Add(Format(DT2.Rows(K)("eos"), "dd/MM/yyyy")) 'DT2.Rows(K)("EOS").ToString)
            ListView1.Items(N).SubItems.Add(DT2.Rows(K)("KREBATIA").ToString)
            ListView1.Items(N).SubItems.Add(DT2.Rows(K)("CATEGORY").ToString)

            N = N + 1

        Next
    End Sub






    Private Sub DIORTOSI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DIORTOSI.Click

        ' ΕΝΗΜΕΡΩΣΗ
        If DIORTOSI.BackColor = Color.Green Then
            If Len(cat.Text) = 0 Then
                MsgBox("ΚΑΤΗΓΟΡΊΑ;")

                cat.Focus()
                Exit Sub


            End If
            If Len(kreb.Text) = 0 Then
                MsgBox("ΚΛΙΝΕΣ;")

                kreb.Focus()
                Exit Sub

            End If
            ExecuteSQLQuery("update HOTROOMS set KREBATIA=" + kreb.Text + " WHERE ROOMN='" + domatio.Text + "'")
            'Format(D1.Value, "MM/dd/yyyy") + "'
            ExecuteSQLQuery("update HOTROOMS set APO='" + Format(D1.Value, "MM/dd/yyyy") + "' WHERE ROOMN='" + domatio.Text + "'")

            ExecuteSQLQuery("update HOTROOMS set EOS='" + Format(D2.Value, "MM/dd/yyyy") + "' WHERE ROOMN='" + domatio.Text + "'")

            ExecuteSQLQuery("update HOTROOMS set CATEGORY=" + cat.Text + " WHERE ROOMN='" + domatio.Text + "'")


            DIORTOSI.BackColor = Color.LightGray

            DIORTOSI.Text = "Διόρθωση"
            domatio.Enabled = True
            PAINTROOMS()
            DIAGRAFI.Enabled = True
            Button1.Enabled = True
            ListView1.Enabled = True


        Else   'ΔΙΟΡΘΩΣΗ





            Dim n As Integer
            Try
                n = ListView1.FocusedItem.Index ' OR ListView1.SelectedIndices(0)
                DIAGRAFI.Enabled = False
                Button1.Enabled = False
                ListView1.Enabled = False


                domatio.Text = ListView1.Items(n).SubItems(0).Text
                D1.Value = ListView1.Items(n).SubItems(1).Text
                D2.Value = ListView1.Items(n).SubItems(2).Text
                kreb.Text = ListView1.Items(n).SubItems(3).Text
                cat.Text = ListView1.Items(n).SubItems(4).Text
                domatio.Enabled = False

                'Dim n As Integer = ListView1.FocusedItem.Index
                ' ExecuteSQLQuery("delete from HOTELROOMS WHERE ROOMN=" + a + "")
                'PAINTROOMS()
                DIORTOSI.BackColor = Color.Green
                DIORTOSI.Text = "Καταχώρηση"
            Catch
                MsgBox("Διαλέξτε δωμάτιο")
            End Try

        End If
    End Sub

  
    Private Sub CreateDays_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CreateDays.Click
        Dim HOTROOMS As New DataTable
        ExecuteSQLQuery("select COUNT(*) from HOTROOMDAYS WHERE HOTELID=" + Str(ID), HOTROOMS)
        If HOTROOMS(0)(0) > 0 Then
            Dim ANS As Integer
            ANS = MsgBox("ΥΠΑΡΧΟΥΝ " + Str(HOTROOMS(0)(0)) + " ΕΓΓΡΑΦΕΣ. Να σβηστούν;", MsgBoxStyle.YesNo)
            If ANS = MsgBoxResult.No Then
                Exit Sub
            End If
            ExecuteSQLQuery("delete from HOTROOMDAYS WHERE HOTELID=" + Str(ID), HOTROOMS)


        End If
        ExecuteSQLQuery("select * from HOTROOMS WHERE HOTELID=" + Str(ID), HOTROOMS)
        Dim N As Integer
        For K As Integer = 0 To HOTROOMS.Rows.Count - 1
            Dim mApo As Date = HOTROOMS.Rows(K)("apo")
            Dim mEos As Date = HOTROOMS.Rows(K)("eos")
            Dim hmeres As Integer = DateDiff("d", mApo, mEos) + 1
            Dim hmera As Date = mApo
            Dim SQL As String
            For nn As Integer = 1 To hmeres



                SQL = "('" + HotelName.Text + "','" + HOTROOMS(K)("ROOMN") + "'," + HOTROOMS.Rows(K)("ID").ToString + "," + HOTROOMS.Rows(K)("HOTELID").ToString
                SQL = SQL + ",'" + Format(hmera, "MM/dd/yyyy") + "')"
                ExecuteSQLQuery("INSERT INTO HOTROOMDAYS (HOTELNAME,ROOMN,IDROOM,HOTELID,DATECHECKIN) VALUES " + SQL)
                hmera = DateAdd("d", 1, hmera)

            Next

        Next
        MsgBox("ok")
    End Sub
End Class