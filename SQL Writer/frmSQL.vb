Public Class frmSQL
    Private Sub frmSQL_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtSQL.Text = ""
    End Sub

    Public Sub LoadText(ByVal thisSQL As String)
        txtSQL.Text = thisSQL
        txtSQL.SelectionStart = txtSQL.Text.Length + 1
    End Sub
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        frmMain.rtxtSelect.Text = txtSQL.Text
        Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Close()
    End Sub

    Private Sub txtSQL_KeyDown(sender As Object, e As KeyEventArgs) Handles txtSQL.KeyDown
        If e.Control And (e.KeyCode = Keys.A) Then

            If sender IsNot Nothing Then
                txtSQL.SelectAll()
            End If
            e.Handled = True
            e.SuppressKeyPress = True
        End If
    End Sub
End Class