'//Before the declaration of the Form Class add the following code
Imports System.Runtime.InteropServices
Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "Form1"

    End Sub

#End Region

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '//and now in your form put this code in the load event 
        '// Adds the icon 
        With uNIF
            .cbSize = Marshal.SizeOf(uNIF)
            .hwnd = Me.Handle
            .uID = 1
            .dwInfoFlags = NIF_ICON Or NIF_MESSAGE
            .uCallbackMessage = New IntPtr(&H500)
            .uVersion = NOTIFYICON_VERSION
            .hIcon = Me.Icon.Handle
        End With
        Result = Shell_NotifyIcon(NIM_ADD, uNIF)
        '// Send a balloon message 
        With uNIF
            .uFlags = NIF_INFO
            .uVersion = 2000
            .szInfoTitle = "Pop-up Tip"
            .szInfo = "Testing Windows 2000+ Ballon Pop-up Tips."
            .dwInfoFlags = NIIF_INFO
        End With
        Result = Shell_NotifyIcon(NIM_MODIFY, uNIF)
        '//if you want to send a balloon message with
        '//the error icon just change de dwInfoFlags to
        '//NIIF_ERROR, or for a warning NIIF_WARNING and
        '//so on, if you want to send messages in other
        '//parts of your project just put the code that
        '//sends the balloon message in the event,
        '//CAUTION: if you don't put the Add icon code
        '//in your main form load event you will not be 
        '//able to receive balloon messages. 
        '//Please send me your comments to drkmouse@prodigy.net.mx
    End Sub
End Class
