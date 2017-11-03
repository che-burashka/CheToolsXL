Option Strict Off
Option Explicit On

Imports Excel = NetOffice.ExcelApi

Friend Class frmApplicationCaption

    
    Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
    Public Sub New()
        MyBase.New()
        If m_vb6FormDefInstance Is Nothing Then
            If m_InitializingDefInstance Then
                m_vb6FormDefInstance = Me
            Else
                Try
                    'For the start-up form, the first instance created is the default instance.
                    If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
                        m_vb6FormDefInstance = Me
                    End If
                Catch
                End Try
            End If
        End If
        'This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents cmdOK As System.Windows.Forms.Button
    Public WithEvents txtNameInput As System.Windows.Forms.TextBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmApplicationCaption))
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
        Me.ToolTip1.Active = True
        Me.cmdOK = New System.Windows.Forms.Button
        Me.txtNameInput = New System.Windows.Forms.TextBox
        Me.Text = "Name for Excel main window"
        Me.ClientSize = New System.Drawing.Size(312, 213)
        Me.Location = New System.Drawing.Point(4, 23)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
        Me.ControlBox = True
        Me.Enabled = True
        Me.KeyPreview = False
        Me.MaximizeBox = True
        Me.MinimizeBox = True
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = True
        Me.HelpButton = False
        Me.WindowState = System.Windows.Forms.FormWindowState.Normal
        Me.Name = "frmApplicationCaption"
        Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.cmdOK.Text = "OK"
        Me.cmdOK.Size = New System.Drawing.Size(97, 41)
        Me.cmdOK.Location = New System.Drawing.Point(112, 152)
        Me.cmdOK.TabIndex = 1
        Me.cmdOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
        Me.cmdOK.CausesValidation = True
        Me.cmdOK.Enabled = True
        Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdOK.TabStop = True
        Me.cmdOK.Name = "cmdOK"
        Me.txtNameInput.AutoSize = False
        Me.txtNameInput.Size = New System.Drawing.Size(209, 41)
        Me.txtNameInput.Location = New System.Drawing.Point(56, 88)
        Me.txtNameInput.TabIndex = 0
        Me.txtNameInput.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNameInput.AcceptsReturn = True
        Me.txtNameInput.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.txtNameInput.BackColor = System.Drawing.SystemColors.Window
        Me.txtNameInput.CausesValidation = True
        Me.txtNameInput.Enabled = True
        Me.txtNameInput.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNameInput.HideSelection = True
        Me.txtNameInput.ReadOnly = False
        Me.txtNameInput.MaxLength = 0
        Me.txtNameInput.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNameInput.Multiline = False
        Me.txtNameInput.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNameInput.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtNameInput.TabStop = True
        Me.txtNameInput.Visible = True
        Me.txtNameInput.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtNameInput.Name = "txtNameInput"
        Me.Controls.Add(cmdOK)
        Me.Controls.Add(txtNameInput)
    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As frmApplicationCaption
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As frmApplicationCaption
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New frmApplicationCaption
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal Value As frmApplicationCaption)
            m_vb6FormDefInstance = Value
        End Set
    End Property
#End Region

    Public Shared oHostApp As Excel.Application

    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Integer, ByRef lpdwProcessId As Integer) As Integer


    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click

        Dim pid As Integer
        '' requires 'Trust programmatic access to VBA projects'
        Call GetWindowThreadProcessId(oHostApp.VBE.MainWindow.HWnd, pid)
        oHostApp.Caption = "XL" & CStr(pid) & ": " & Me.txtNameInput.Text
        Me.Close()
    End Sub

    Private Sub frmApplicationCaption_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        txtNameInput.Enabled = True
        '' txtNameInput.SetFocus ' crahes here. focus is aleady on this text box ...
    End Sub


    Private Sub txtNameInput_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtNameInput.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        ''
        '' Debug.Print KeyAscii
        If KeyAscii = 13 Then
            cmdOK_Click(cmdOK, New System.EventArgs)
        End If

        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
End Class