Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DXFFix()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        GetAcadApp()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        GetBigDXFs("tr-ltz-vlt01.tait.rocks", "TAIT_TOWERS")
        'KillItWithFire("https://vaultsandbox.tait.rocks/", "TAIT_TOWERS", "Administrator", "", "$/Test9_Designs/DISNEY HKDL CSS 2018 - 18A0029/015995-12/215337/")
    End Sub

    Private Sub btnFixVault_Click(sender As Object, e As EventArgs) Handles btnFixVault.Click
        FixVault("tr-ltz-vlt01.tait.rocks", "TAIT_TOWERS")
    End Sub
End Class

