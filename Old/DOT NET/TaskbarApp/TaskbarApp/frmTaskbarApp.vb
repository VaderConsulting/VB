Public Class frmMain
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
    Friend WithEvents pnlTaskBar As System.Windows.Forms.Panel
    Friend WithEvents cmdStart As System.Windows.Forms.Button
    Friend WithEvents pnlStartMenu As System.Windows.Forms.Panel
    Friend WithEvents TipShutdown As System.Windows.Forms.ToolTip
    Friend WithEvents tipBegin As System.Windows.Forms.ToolTip
    Friend WithEvents cmdShutdown As System.Windows.Forms.Button
    Private components As System.ComponentModel.IContainer

    'Required by the Windows Form Designer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
        Me.cmdShutdown = New System.Windows.Forms.Button()
        Me.pnlStartMenu = New System.Windows.Forms.Panel()
        Me.cmdStart = New System.Windows.Forms.Button()
        Me.pnlTaskBar = New System.Windows.Forms.Panel()
        Me.TipShutdown = New System.Windows.Forms.ToolTip(Me.components)
        Me.tipBegin = New System.Windows.Forms.ToolTip(Me.components)
        Me.pnlStartMenu.SuspendLayout()
        Me.pnlTaskBar.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdShutdown
        '
        Me.cmdShutdown.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdShutdown.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdShutdown.Image = CType(resources.GetObject("cmdShutdown.Image"), System.Drawing.Bitmap)
        Me.cmdShutdown.ImageAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.cmdShutdown.Location = New System.Drawing.Point(-3, 120)
        Me.cmdShutdown.Name = "cmdShutdown"
        Me.cmdShutdown.Size = New System.Drawing.Size(115, 32)
        Me.cmdShutdown.TabIndex = 1
        Me.cmdShutdown.Text = "Shut Down..."
        Me.cmdShutdown.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.TipShutdown.SetToolTip(Me.cmdShutdown, "Shutdown")
        '
        'pnlStartMenu
        '
        Me.pnlStartMenu.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.pnlStartMenu.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.pnlStartMenu.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlStartMenu.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdShutdown})
        Me.pnlStartMenu.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.pnlStartMenu.Location = New System.Drawing.Point(0, 182)
        Me.pnlStartMenu.Name = "pnlStartMenu"
        Me.pnlStartMenu.Size = New System.Drawing.Size(112, 154)
        Me.pnlStartMenu.TabIndex = 1
        Me.pnlStartMenu.Visible = False
        '
        'cmdStart
        '
        Me.cmdStart.BackColor = System.Drawing.Color.Green
        Me.cmdStart.Font = New System.Drawing.Font("Verdana", 8.25!, (System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStart.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdStart.Name = "cmdStart"
        Me.cmdStart.Size = New System.Drawing.Size(64, 24)
        Me.cmdStart.TabIndex = 0
        Me.cmdStart.Text = "Start"
        Me.tipBegin.SetToolTip(Me.cmdStart, "Click here to begin")
        '
        'pnlTaskBar
        '
        Me.pnlTaskBar.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.pnlTaskBar.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdStart})
        Me.pnlTaskBar.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlTaskBar.Location = New System.Drawing.Point(0, 334)
        Me.pnlTaskBar.Name = "pnlTaskBar"
        Me.pnlTaskBar.Size = New System.Drawing.Size(520, 24)
        Me.pnlTaskBar.TabIndex = 0
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(520, 358)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.pnlTaskBar, Me.pnlStartMenu})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Main Control"
        Me.pnlStartMenu.ResumeLayout(False)
        Me.pnlTaskBar.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStart.Click
        pnlStartMenu.Visible = True
    End Sub

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub frmMain_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Click
        pnlStartMenu.Visible = False
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShutdown.Click
        Dim Retval As DialogResult = DialogResult.No
        pnlStartMenu.Visible = False
        Retval = MessageBox.Show("Are you sure you want to exit?", "Shutdown", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.DefaultDesktopOnly)
        If Retval = DialogResult.Yes Then
            End
        End If
    End Sub

    Private Sub cmdShutdown_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdShutdown.MouseLeave
        cmdShutdown.ForeColor = System.Drawing.Color.Black
    End Sub

    Private Sub cmdShutdown_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdShutdown.MouseEnter
        cmdShutdown.ForeColor = System.Drawing.Color.DarkBlue
    End Sub
End Class
