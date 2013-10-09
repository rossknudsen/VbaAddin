Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Drawing

<ComVisible(True), Guid("56948493-b494-4ddb-bed7-2299f3d3bfa1"), ProgId("MyVbaAddin.UserControlHost")> _
Public Class UserControlHost

    Private Class SubClassingWindow
        Inherits System.Windows.Forms.NativeWindow

        Public Event CallBackProc(ByRef m As Message)

        Public Sub New(ByVal handle As IntPtr)
            MyBase.AssignHandle(handle)
        End Sub

        Protected Overrides Sub WndProc(ByRef m As Message)

            Const WM_SIZE As Integer = &H5

            If m.Msg = WM_SIZE Then
                RaiseEvent CallBackProc(m)
            End If

            MyBase.WndProc(m)

        End Sub

        Protected Overrides Sub Finalize()

            Me.ReleaseHandle()

            MyBase.Finalize()

        End Sub

    End Class

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure RECT
        Friend Left As Integer
        Friend Top As Integer
        Friend Right As Integer
        Friend Bottom As Integer
    End Structure

    Private Declare Function GetParent Lib "user32" (ByVal hWnd As IntPtr) As IntPtr
    Private Declare Function GetClientRect Lib "user32" Alias "GetClientRect" (ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Integer

    Private _parentHandle As IntPtr
    Private WithEvents _subClassingWindow As SubClassingWindow

    Friend Sub AddUserControl(ByVal control As UserControl)

        _parentHandle = GetParent(Me.Handle)

        _subClassingWindow = New SubClassingWindow(_parentHandle)

        control.Dock = DockStyle.Fill

        Me.Controls.Add(control)

        AdjustSize()

    End Sub

    Private Sub _subClassingWindow_CallBackProc(ByRef m As System.Windows.Forms.Message) Handles _subClassingWindow.CallBackProc

        AdjustSize()

    End Sub

    Private Sub AdjustSize()

        Dim tRect As RECT

        If GetClientRect(_parentHandle, tRect) <> 0 Then

            Me.Size = New Size(tRect.Right - tRect.Left, tRect.Bottom - tRect.Top)

        End If

    End Sub

    Protected Overrides Function ProcessKeyPreview(ByRef m As System.Windows.Forms.Message) As Boolean

        Const WM_KEYDOWN As Integer = &H100

        Dim result As Boolean = False
        Dim pressedKey As Keys
        Dim hostedUserControl As UserControl
        Dim activeButton As Button

        hostedUserControl = DirectCast(Me.Controls.Item(0), UserControl)

        If m.Msg = WM_KEYDOWN Then

            pressedKey = CType(m.WParam, Keys)

            Select Case pressedKey

                Case Keys.Tab

                    If Control.ModifierKeys = Keys.None Then ' Tab

                        Me.SelectNextControl(hostedUserControl.ActiveControl, True, True, True, True)
                        result = True

                    ElseIf Control.ModifierKeys = Keys.Shift Then ' Shift + Tab

                        Me.SelectNextControl(hostedUserControl.ActiveControl, False, True, True, True)
                        result = True

                    End If

                Case Keys.Return

                    If TypeOf hostedUserControl.ActiveControl Is Button Then

                        activeButton = DirectCast(hostedUserControl.ActiveControl, Button)
                        activeButton.PerformClick()

                    End If

            End Select

        End If

        If result = False Then
            result = MyBase.ProcessKeyPreview(m)
        End If

        Return result

    End Function

End Class