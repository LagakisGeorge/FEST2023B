Public Class FrmADDSUPPLIER_ITEM

    Private Sub cmdcancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdcancel.Click
        Me.Close()
    End Sub

    Private Sub cmdsave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdsave.Click
        Dim i As Integer
        With FrmAddSupplier.lstProducts

        End With
        '        With FrmAddSupplier.lstProducts
        '            .Items.Add(lstItems.Items(i).Text)
        '        End With
        '        For x = 0 To FrmAddSupplier.lstProducts.Items.Count - 1
        '            FrmAddSupplier.lstProducts.Items(x).SubItems.Add(lstItems.Items(i).SubItems(1).Text)
        '        Next
        If txtid.Text = "" Then
            MsgBox("Cannot save this record, select an item first ", MsgBoxStyle.Information, "Sales and Inventory")
            Exit Sub
        End If
        With FrmAddSupplier
            For i = 0 To .lstProducts.Items.Count - 1
                If .lstProducts.Items(i).SubItems(1).Text = txtid.Text Then
                    MsgBox("Selected item is already on the list !!", MsgBoxStyle.Exclamation, "Sales and Inventory")
                    Exit Sub
                End If
            Next
            .lstProducts.Items.Add(txtcategory.Text)
            .lstProducts.Items(.lstProducts.Items.Count - 1).SubItems.Add(txtid.Text)
            .lstProducts.Items(.lstProducts.Items.Count - 1).SubItems.Add(txtname.Text)
            .lstProducts.Items(.lstProducts.Items.Count - 1).SubItems.Add(txtprice.Text)
            If MsgBox("Do you want to add new record ?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Sales and Inventory") = MsgBoxResult.Yes Then
                FrmCATEGORYSELECTitem.ShowDialog()
            Else
                Me.Close()
            End If
        End With
    End Sub

    Private Sub FrmADDSUPPLIER_ITEM_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FrmCATEGORYSELECTitem.ShowDialog()
        txtprice.Select()
        txtprice.Focus()
    End Sub
End Class