Public Class frm_Check
    Private Sub btnImportFromExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImportFromExcel.Click
        Try
            Dim oMain As New WMS_LIB.clsLibrary
            oMain.mainFunction()
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub
End Class
