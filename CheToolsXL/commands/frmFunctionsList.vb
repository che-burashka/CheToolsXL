Option Strict Off
Option Explicit On
Friend Class frmFunctionsList
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
	Public WithEvents cmdSave As System.Windows.Forms.Button
	Public WithEvents lstFunctionNames As System.Windows.Forms.ListBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmFunctionsList))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.ToolTip1.Active = True
		Me.cmdSave = New System.Windows.Forms.Button
		Me.lstFunctionNames = New System.Windows.Forms.ListBox
		Me.Text = "Functions used in ActiveWorkbook"
		Me.ClientSize = New System.Drawing.Size(343, 467)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
		Me.Name = "frmFunctionsList"
		Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdSave.Text = "Save in file"
		Me.cmdSave.Size = New System.Drawing.Size(113, 25)
		Me.cmdSave.Location = New System.Drawing.Point(120, 432)
		Me.cmdSave.TabIndex = 1
		Me.cmdSave.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
		Me.cmdSave.CausesValidation = True
		Me.cmdSave.Enabled = True
		Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdSave.TabStop = True
		Me.cmdSave.Name = "cmdSave"
		Me.lstFunctionNames.Size = New System.Drawing.Size(289, 410)
		Me.lstFunctionNames.Location = New System.Drawing.Point(24, 16)
		Me.lstFunctionNames.TabIndex = 0
		Me.lstFunctionNames.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstFunctionNames.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstFunctionNames.BackColor = System.Drawing.SystemColors.Window
		Me.lstFunctionNames.CausesValidation = True
		Me.lstFunctionNames.Enabled = True
		Me.lstFunctionNames.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstFunctionNames.IntegralHeight = True
		Me.lstFunctionNames.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstFunctionNames.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstFunctionNames.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstFunctionNames.Sorted = False
		Me.lstFunctionNames.TabStop = True
		Me.lstFunctionNames.Visible = True
		Me.lstFunctionNames.MultiColumn = False
		Me.lstFunctionNames.Name = "lstFunctionNames"
		Me.Controls.Add(cmdSave)
		Me.Controls.Add(lstFunctionNames)
	End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmFunctionsList
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmFunctionsList
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmFunctionsList()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	
	
	Public doSave As Boolean
	
	Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
		
		doSave = True
		Me.Close()
	End Sub
	
	Private Sub frmFunctionsList_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		''
	End Sub
End Class