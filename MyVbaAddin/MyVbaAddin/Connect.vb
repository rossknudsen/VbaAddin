Imports NetOffice
Imports NetOffice.Tools
Imports NetOffice.VBIDEApi
Imports NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums

Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Drawing



<ComVisible(True), Guid("dc521227-3325-4f59-9141-ba4716860287"), ProgId("MyVbaAddin.Connect")> _
Public Class Connect
    Implements IDTExtensibility2

    Private _VBE As VBE
    Private _AddIn As AddIn
    Private WithEvents _CommandBarButton1 As CommandBarButton
    Private WithEvents _CommandBarButton2 As CommandBarButton

    Private _toolWindow1 As Window
    Private _toolWindow2 As Window

    Private Sub OnConnection(Application As Object, ConnectMode As ext_ConnectMode, _
       AddInInst As Object, ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection

        Try

            _VBE = New VBE(Nothing, Application)
            _AddIn = New AddIn(Nothing, AddInInst)

            Select Case ConnectMode

                Case ext_ConnectMode.ext_cm_Startup
                    ' OnStartupComplete will be called

                Case ext_ConnectMode.ext_cm_AfterStartup
                    InitializeAddIn()

            End Select

        Catch ex As Exception

            MessageBox.Show(ex.ToString())

        End Try

    End Sub

    Private Sub OnDisconnection(RemoveMode As ext_DisconnectMode, _
                                ByRef custom As System.Array) Implements IDTExtensibility2.OnDisconnection

        If Not _CommandBarButton1 Is Nothing Then

            _CommandBarButton1.Delete()
            _CommandBarButton1 = Nothing

        End If

        If Not _CommandBarButton2 Is Nothing Then

            _CommandBarButton2.Delete()
            _CommandBarButton2 = Nothing

        End If

    End Sub

    Private Sub OnStartupComplete(ByRef custom As System.Array) _
       Implements IDTExtensibility2.OnStartupComplete

        InitializeAddIn()

    End Sub

    Private Sub OnAddInsUpdate(ByRef custom As System.Array) Implements IDTExtensibility2.OnAddInsUpdate

    End Sub

    Private Sub OnBeginShutdown(ByRef custom As System.Array) Implements IDTExtensibility2.OnBeginShutdown

    End Sub

    Private Sub InitializeAddIn()

        Dim standardCommandBar As CommandBar
        Dim commandBarControl As CommandBarControl

        Try

            standardCommandBar = _VBE.CommandBars.Item("Standard")

            commandBarControl = standardCommandBar.Controls.Add(MsoControlType.msoControlButton)
            _CommandBarButton1 = DirectCast(commandBarControl, CommandBarButton)
            _CommandBarButton1.Caption = "Toolwindow 1"
            _CommandBarButton1.FaceId = 59
            _CommandBarButton1.Style = MsoButtonStyle.msoButtonIconAndCaption
            _CommandBarButton1.BeginGroup = True

            commandBarControl = standardCommandBar.Controls.Add(MsoControlType.msoControlButton)
            _CommandBarButton2 = DirectCast(commandBarControl, CommandBarButton)
            _CommandBarButton2.Caption = "Toolwindow 2"
            _CommandBarButton2.FaceId = 59
            _CommandBarButton2.Style = MsoButtonStyle.msoButtonIconAndCaption
            _CommandBarButton2.BeginGroup = True

        Catch ex As Exception

            MessageBox.Show(ex.ToString())

        End Try

    End Sub

    Private Function CreateToolWindow(ByVal toolWindowCaption As String, _
                                      ByVal toolWindowGuid As String, _
                                      ByVal toolWindowUserControl As UserControl) As Window

        Dim userControlObject As Object = Nothing
        Dim userControlHost As UserControlHost
        Dim toolWindow As Window
        Dim progId As String

        ' IMPORTANT: ensure that you use the same ProgId value used in the ProgId attribute of the UserControlHost class
        progId = "MyVbaAddin.UserControlHost"

        toolWindow = _VBE.Windows.CreateToolWindow(_AddIn, progId, toolWindowCaption, toolWindowGuid, userControlObject)
        userControlHost = DirectCast(userControlObject, UserControlHost)

        toolWindow.Visible = True

        userControlHost.AddUserControl(toolWindowUserControl)

        Return toolWindow

    End Function

    Private Sub _CommandBarButton1_Click(Ctrl As CommandBarButton, _
                                         ByRef CancelDefault As Boolean) Handles _CommandBarButton1.ClickEvent

        Dim userControlObject As Object = Nothing
        Dim userControlToolWindow1 As UserControlToolWindow1

        Try

            If _toolWindow1 Is Nothing Then

                userControlToolWindow1 = New UserControlToolWindow1()

                ' TODO: Change the GUID
                _toolWindow1 = CreateToolWindow("My toolwindow 1", "{e80c0630-a44c-44ad-86b5-61d8bf664d42}", userControlToolWindow1)

                userControlToolWindow1.Initialize(_VBE)

            Else

                _toolWindow1.Visible = True

            End If

        Catch ex As Exception

            MessageBox.Show(ex.ToString)

        End Try

    End Sub

    Private Sub _CommandBarButton2_Click(Ctrl As CommandBarButton, _
                                         ByRef CancelDefault As Boolean) Handles _CommandBarButton2.ClickEvent

        Dim userControlObject As Object = Nothing
        Dim userControlToolWindow2 As UserControlToolWindow2

        Try

            If _toolWindow2 Is Nothing Then

                userControlToolWindow2 = New UserControlToolWindow2()

                ' TODO: Change the GUID
                _toolWindow2 = CreateToolWindow("My toolwindow 2", "{ffc9cc65-209e-4caf-ac9c-ee12647e9c9f}", userControlToolWindow2)

                userControlToolWindow2.Initialize(_VBE)

            Else

                _toolWindow2.Visible = True

            End If

        Catch ex As Exception

            MessageBox.Show(ex.ToString)

        End Try

    End Sub

End Class
