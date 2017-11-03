Option Strict Off
Option Explicit On




Friend Class frmPasteSpecial
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
	Public WithEvents chkClearClipboard As System.Windows.Forms.CheckBox
	Public WithEvents chkFormulas As System.Windows.Forms.CheckBox
	Public WithEvents chkTranspose As System.Windows.Forms.CheckBox
	Public WithEvents chkFormats As System.Windows.Forms.CheckBox
	Public WithEvents chkValues As System.Windows.Forms.CheckBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdPaste As System.Windows.Forms.Button
	Public WithEvents lblClear As System.Windows.Forms.Label
	Public WithEvents lblFormulas As System.Windows.Forms.Label
	Public WithEvents lblTranspose As System.Windows.Forms.Label
	Public WithEvents lblFormats As System.Windows.Forms.Label
	Public WithEvents lblValues As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkClearClipboard = New System.Windows.Forms.CheckBox
        Me.chkFormulas = New System.Windows.Forms.CheckBox
        Me.chkTranspose = New System.Windows.Forms.CheckBox
        Me.chkFormats = New System.Windows.Forms.CheckBox
        Me.chkValues = New System.Windows.Forms.CheckBox
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdPaste = New System.Windows.Forms.Button
        Me.lblClear = New System.Windows.Forms.Label
        Me.lblFormulas = New System.Windows.Forms.Label
        Me.lblTranspose = New System.Windows.Forms.Label
        Me.lblFormats = New System.Windows.Forms.Label
        Me.lblValues = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'chkClearClipboard
        '
        Me.chkClearClipboard.BackColor = System.Drawing.SystemColors.Control
        Me.chkClearClipboard.Checked = True
        Me.chkClearClipboard.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkClearClipboard.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkClearClipboard.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkClearClipboard.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkClearClipboard.Location = New System.Drawing.Point(200, 96)
        Me.chkClearClipboard.Name = "chkClearClipboard"
        Me.chkClearClipboard.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkClearClipboard.Size = New System.Drawing.Size(17, 17)
        Me.chkClearClipboard.TabIndex = 10
        Me.chkClearClipboard.Text = "Check1"
        Me.chkClearClipboard.UseVisualStyleBackColor = False
        '
        'chkFormulas
        '
        Me.chkFormulas.BackColor = System.Drawing.SystemColors.Control
        Me.chkFormulas.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkFormulas.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFormulas.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkFormulas.Location = New System.Drawing.Point(200, 32)
        Me.chkFormulas.Name = "chkFormulas"
        Me.chkFormulas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkFormulas.Size = New System.Drawing.Size(17, 17)
        Me.chkFormulas.TabIndex = 8
        Me.chkFormulas.Text = "Check1"
        Me.chkFormulas.UseVisualStyleBackColor = False
        '
        'chkTranspose
        '
        Me.chkTranspose.BackColor = System.Drawing.SystemColors.Control
        Me.chkTranspose.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTranspose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkTranspose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTranspose.Location = New System.Drawing.Point(200, 64)
        Me.chkTranspose.Name = "chkTranspose"
        Me.chkTranspose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTranspose.Size = New System.Drawing.Size(17, 17)
        Me.chkTranspose.TabIndex = 6
        Me.chkTranspose.Text = "Check1"
        Me.chkTranspose.UseVisualStyleBackColor = False
        '
        'chkFormats
        '
        Me.chkFormats.BackColor = System.Drawing.SystemColors.Control
        Me.chkFormats.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkFormats.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFormats.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkFormats.Location = New System.Drawing.Point(40, 64)
        Me.chkFormats.Name = "chkFormats"
        Me.chkFormats.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkFormats.Size = New System.Drawing.Size(17, 17)
        Me.chkFormats.TabIndex = 4
        Me.chkFormats.Text = "Check1"
        Me.chkFormats.UseVisualStyleBackColor = False
        '
        'chkValues
        '
        Me.chkValues.BackColor = System.Drawing.SystemColors.Control
        Me.chkValues.Checked = True
        Me.chkValues.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkValues.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkValues.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkValues.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkValues.Location = New System.Drawing.Point(40, 32)
        Me.chkValues.Name = "chkValues"
        Me.chkValues.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkValues.Size = New System.Drawing.Size(17, 17)
        Me.chkValues.TabIndex = 2
        Me.chkValues.Text = "Check1"
        Me.chkValues.UseVisualStyleBackColor = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(184, 179)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(113, 33)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.UseVisualStyleBackColor = False
        '
        'cmdPaste
        '
        Me.cmdPaste.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPaste.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPaste.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPaste.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPaste.Location = New System.Drawing.Point(48, 179)
        Me.cmdPaste.Name = "cmdPaste"
        Me.cmdPaste.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPaste.Size = New System.Drawing.Size(113, 33)
        Me.cmdPaste.TabIndex = 0
        Me.cmdPaste.Text = "Paste"
        Me.cmdPaste.UseVisualStyleBackColor = False
        '
        'lblClear
        '
        Me.lblClear.BackColor = System.Drawing.SystemColors.Control
        Me.lblClear.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblClear.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblClear.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblClear.Location = New System.Drawing.Point(240, 96)
        Me.lblClear.Name = "lblClear"
        Me.lblClear.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblClear.Size = New System.Drawing.Size(73, 17)
        Me.lblClear.TabIndex = 11
        Me.lblClear.Text = "Clear Clibboard"
        '
        'lblFormulas
        '
        Me.lblFormulas.BackColor = System.Drawing.SystemColors.Control
        Me.lblFormulas.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFormulas.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFormulas.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFormulas.Location = New System.Drawing.Point(240, 32)
        Me.lblFormulas.Name = "lblFormulas"
        Me.lblFormulas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFormulas.Size = New System.Drawing.Size(65, 17)
        Me.lblFormulas.TabIndex = 9
        Me.lblFormulas.Text = "Formulas"
        '
        'lblTranspose
        '
        Me.lblTranspose.BackColor = System.Drawing.SystemColors.Control
        Me.lblTranspose.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTranspose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTranspose.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTranspose.Location = New System.Drawing.Point(240, 64)
        Me.lblTranspose.Name = "lblTranspose"
        Me.lblTranspose.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTranspose.Size = New System.Drawing.Size(65, 17)
        Me.lblTranspose.TabIndex = 7
        Me.lblTranspose.Text = "Transpose"
        '
        'lblFormats
        '
        Me.lblFormats.BackColor = System.Drawing.SystemColors.Control
        Me.lblFormats.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFormats.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFormats.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFormats.Location = New System.Drawing.Point(80, 64)
        Me.lblFormats.Name = "lblFormats"
        Me.lblFormats.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFormats.Size = New System.Drawing.Size(65, 17)
        Me.lblFormats.TabIndex = 5
        Me.lblFormats.Text = "Formats"
        '
        'lblValues
        '
        Me.lblValues.BackColor = System.Drawing.SystemColors.Control
        Me.lblValues.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblValues.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblValues.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblValues.Location = New System.Drawing.Point(80, 32)
        Me.lblValues.Name = "lblValues"
        Me.lblValues.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblValues.Size = New System.Drawing.Size(65, 17)
        Me.lblValues.TabIndex = 3
        Me.lblValues.Text = "Values"
        '
        'frmPasteSpecial
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(348, 252)
        Me.Controls.Add(Me.chkClearClipboard)
        Me.Controls.Add(Me.chkFormulas)
        Me.Controls.Add(Me.chkTranspose)
        Me.Controls.Add(Me.chkFormats)
        Me.Controls.Add(Me.chkValues)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdPaste)
        Me.Controls.Add(Me.lblClear)
        Me.Controls.Add(Me.lblFormulas)
        Me.Controls.Add(Me.lblTranspose)
        Me.Controls.Add(Me.lblFormats)
        Me.Controls.Add(Me.lblValues)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(4, 27)
        Me.Name = "frmPasteSpecial"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Paste Special"
        Me.ResumeLayout(False)

    End Sub
#End Region 
#Region "Upgrade Support "
	Private Shared m_vb6FormDefInstance As frmPasteSpecial
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmPasteSpecial
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmPasteSpecial()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 
	
	Public OKCancel As Boolean
	
	Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Integer) As Integer
	
	'UPGRADE_WARNING: Event chkValues.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub chkValues_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkValues.CheckStateChanged
		
		If chkValues.CheckState = 1 Then
			Me.chkFormulas.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
	End Sub
	
	'UPGRADE_WARNING: Event chkFormulas.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub chkFormulas_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkFormulas.CheckStateChanged
		
		If chkFormulas.CheckState = 1 Then
			Me.chkValues.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Hide()
	End Sub
	
	Private Sub cmdPaste_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPaste.Click
		OKCancel = True
		Me.Hide()
	End Sub
	
	'UPGRADE_WARNING: Form event frmPasteSpecial.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub frmPasteSpecial_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		OKCancel = False
		cmdPaste.Focus()
		BringWindowToTop(Me.Handle.ToInt32)
		
	End Sub
End Class