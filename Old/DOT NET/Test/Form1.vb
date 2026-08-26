' Windows Balloon
' Rules of Use:
' 1.  Modify as you see fit.
' 2.  Send me an email so I know how far it has gotten.
' 3.  All code remains Open Source, but may be used by anyone for any reason.
' 4.  All comments remain unless they are abusive.
' 5.  Enjoy!
'
' -- Date ---   -- Who ----------- Changes ----------------------------------------
' 30 Mar 2002   D. Robinson        Initial Version.
'
' ---------------------------------------------------------------------------------
' TODO:  1.  Make as a Component
' TODO:  2.  Make background colour settable
' TODO:  3.  Prevent messages from displaying outside of the balloon

Imports System.Drawing
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
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents picIcon As System.Windows.Forms.PictureBox
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.picIcon = New System.Windows.Forms.PictureBox()
        Me.lblMessage = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblTitle
        '
        Me.lblTitle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(225, Byte))
        Me.lblTitle.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.Location = New System.Drawing.Point(32, 8)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(184, 16)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "Title"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(225, Byte))
        Me.LinkLabel1.Location = New System.Drawing.Point(224, 120)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(96, 16)
        Me.LinkLabel1.TabIndex = 3
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "LinkLabel1"
        Me.LinkLabel1.Visible = False
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(32, 144)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(232, 16)
        Me.ProgressBar1.TabIndex = 4
        Me.ProgressBar1.Visible = False
        '
        'ImageList1
        '
        Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth16Bit
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.PictureBox1.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(225, Byte))
        Me.PictureBox1.Location = New System.Drawing.Point(376, 8)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(18, 18)
        Me.PictureBox1.TabIndex = 5
        Me.PictureBox1.TabStop = False
        '
        'picIcon
        '
        Me.picIcon.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(225, Byte))
        Me.picIcon.Location = New System.Drawing.Point(8, 8)
        Me.picIcon.Name = "picIcon"
        Me.picIcon.Size = New System.Drawing.Size(18, 18)
        Me.picIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.picIcon.TabIndex = 6
        Me.picIcon.TabStop = False
        '
        'lblMessage
        '
        Me.lblMessage.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(225, Byte))
        Me.lblMessage.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMessage.Location = New System.Drawing.Point(32, 32)
        Me.lblMessage.Name = "lblMessage"
        Me.lblMessage.Size = New System.Drawing.Size(336, 16)
        Me.lblMessage.TabIndex = 7
        Me.lblMessage.Text = "Message"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(400, 176)
        Me.ControlBox = False
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblMessage, Me.LinkLabel1, Me.picIcon, Me.PictureBox1, Me.ProgressBar1, Me.lblTitle})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form1"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.TopMost = True
        Me.TransparencyKey = System.Drawing.Color.White
        Me.ResumeLayout(False)

    End Sub

#End Region


    Enum ImageType
        img_stop = 0
        img_question = 1
        img_exclamation = 2
        img_information = 3

    End Enum

    Enum OSType
        os_Win16 = 2
        os_Win95 = 6
        os_WinXP = 10
    End Enum

    Dim ImageOffset As Int16 = OSType.os_WinXP
    Dim MinHeight As Integer = 60

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim myGraphics As Graphics
        'Dim myRectangle As Rectangle
        'Dim myPolygon() As Point
        'Dim myPen As New Pen(Color.Black, 1)
        'Dim myBrush As New SolidBrush(Color.FromArgb(252, 255, 255, 225))
        'Dim W, H As Single

        'ReDim myPolygon(45)

        'Me.Height = 80
        'Me.Show()
        'Me.Refresh()
        'W = ActiveForm.Width
        'H = ActiveForm.Height

        'myGraphics = Graphics.FromHwnd(ActiveForm().Handle)

        'myPolygon(0) = New Point(0, 5)
        'myPolygon(1) = New Point(1, 5)
        'myPolygon(2) = New Point(1, 4)
        'myPolygon(3) = New Point(1, 3)
        'myPolygon(4) = New Point(2, 3)
        'myPolygon(5) = New Point(2, 2)
        'myPolygon(6) = New Point(3, 2)
        'myPolygon(7) = New Point(3, 1)
        'myPolygon(8) = New Point(4, 1)
        'myPolygon(9) = New Point(5, 1)
        'myPolygon(10) = New Point(5, 0)
        'myPolygon(11) = New Point(W - 6, 0) ' Top Line
        'myPolygon(12) = New Point(W - 6, 1)
        'myPolygon(13) = New Point(W - 5, 1)
        'myPolygon(14) = New Point(W - 4, 1)
        'myPolygon(15) = New Point(W - 4, 2)
        'myPolygon(16) = New Point(W - 3, 2)
        'myPolygon(17) = New Point(W - 3, 3)
        'myPolygon(18) = New Point(W - 2, 3)
        'myPolygon(19) = New Point(W - 2, 4)
        'myPolygon(20) = New Point(W - 2, 5)
        'myPolygon(21) = New Point(W - 1, 5)
        'myPolygon(22) = New Point(W - 1, H - 36) ' Right Side
        'myPolygon(23) = New Point(W - 2, H - 36)
        'myPolygon(24) = New Point(W - 2, H - 35)
        'myPolygon(25) = New Point(W - 2, H - 34)
        'myPolygon(26) = New Point(W - 3, H - 34)
        'myPolygon(27) = New Point(W - 3, H - 33)
        'myPolygon(28) = New Point(W - 4, H - 33)
        'myPolygon(29) = New Point(W - 5, H - 32)
        'myPolygon(30) = New Point(W - 6, H - 32)
        'myPolygon(31) = New Point(W - 6, H - 31)
        'myPolygon(32) = New Point(W - 30, H - 31) ' Bottom 1
        'myPolygon(33) = New Point(W - 30, H - 1)  ' Callout Right
        'myPolygon(34) = New Point(W - 50, H - 31) ' Callout Left
        'myPolygon(35) = New Point(5, H - 31)      ' Bottom 2
        'myPolygon(36) = New Point(5, H - 32)
        'myPolygon(37) = New Point(4, H - 32)
        'myPolygon(38) = New Point(3, H - 32)
        'myPolygon(39) = New Point(3, H - 33)
        'myPolygon(40) = New Point(2, H - 33)
        'myPolygon(41) = New Point(2, H - 34)
        'myPolygon(42) = New Point(1, H - 34)
        'myPolygon(43) = New Point(1, H - 35)
        'myPolygon(44) = New Point(1, H - 36)
        'myPolygon(45) = New Point(0, H - 36)

        'myGraphics.DrawPolygon(myPen, myPolygon)   ' Complete Polygon (Left Side)
        'myGraphics.FillPolygon(myBrush, myPolygon) ' Fill-in

        '' Ensure messages are inside balloon

        'PictureBox1.Image = ImageList1.Images(0)
        'picIcon.Image = ImageList1.Images(ImageType.img_stop + ImageOffset)
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

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        End
    End Sub

    Private Sub PictureBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove
        PictureBox1.Image = ImageList1.Images(1)
    End Sub

    Private Sub PictureBox1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles PictureBox1.MouseLeave
        PictureBox1.Image = ImageList1.Images(0)
    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        If Me.Height < MinHeight Then Me.Height = MinHeight
    End Sub

End Class

