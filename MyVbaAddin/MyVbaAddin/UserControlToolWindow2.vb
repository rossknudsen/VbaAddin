Imports NetOffice.VBIDEApi
Imports System.Windows.Forms

Friend Class UserControlToolWindow2

    Private _VBE As VBE

    Friend Sub Initialize(ByVal vbe As VBE)

        Me.BackColor = Drawing.Color.Blue

        _VBE = vbe

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        MessageBox.Show("Toolwindow shown in VBA editor version " & _VBE.Version)

    End Sub

End Class
