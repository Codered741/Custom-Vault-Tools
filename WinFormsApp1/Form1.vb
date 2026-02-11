Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DXFFix()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim box = New frmPgBar
        box.ShowDialog()
        ' Dim BackgroundWorker1 As New System.ComponentModel.BackgroundWorker
        ' BackgroundWorker1.RunWorkerAsync()'
        box.Dispose()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        GetBigDXFs("tr-ltz-vlt01.tait.rocks", "TAIT_TOWERS")
    End Sub

    Private Sub btnFixVault_Click(sender As Object, e As EventArgs) Handles btnFixVault.Click
        FixVault("tr-ltz-vlt01.tait.rocks", "TAIT_TOWERS", "11/18/2025")
    End Sub

    Private Sub btnTest1_Click(sender As Object, e As EventArgs) Handles btnTest1.Click
        FixVaultTest("tr-ltz-vlt01.tait.rocks", "TAIT_TOWERS", "11/18/2025")
    End Sub
End Class

