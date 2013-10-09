Imports NetOffice.VBIDEApi
Imports System.Windows.Forms

Friend Class UserControlToolWindow1

    Private _VBE As VBE

    Friend Sub Initialize(ByVal vbe As VBE)

        Me.BackColor = Drawing.Color.Red

        _VBE = vbe

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        MessageBox.Show("Toolwindow shown in VBA editor version " & _VBE.Version)

    End Sub

End Class
