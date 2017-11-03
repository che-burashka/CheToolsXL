Option Strict Off
Option Explicit On

Imports Office = NetOffice.OfficeApi
Imports Excel = NetOffice.ExcelApi

Friend Class frmGlobalFind
    Inherits System.Windows.Forms.Form


    Public g_mostRecentSearchPattern As String
    Public g_recentSearchPatterns As System.Collections.SortedList

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

        g_mostRecentSearchPattern = New String("")
        g_recentSearchPatterns = New System.Collections.SortedList

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
    Public WithEvents cmbFindWhat As System.Windows.Forms.ComboBox
    Public WithEvents lstOutput As System.Windows.Forms.ListBox
    Public WithEvents cmdGoTo As System.Windows.Forms.Button
    Public WithEvents cmdCancel As System.Windows.Forms.Button
    Public WithEvents cmdFind As System.Windows.Forms.Button
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmGlobalFind))
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
        Me.ToolTip1.Active = True
        Me.cmbFindWhat = New System.Windows.Forms.ComboBox
        Me.lstOutput = New System.Windows.Forms.ListBox
        Me.cmdGoTo = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdFind = New System.Windows.Forms.Button
        Me.Text = "Find Globally"
        Me.ClientSize = New System.Drawing.Size(431, 498)
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
        Me.Name = "frmGlobalFind"
        Me.cmbFindWhat.Size = New System.Drawing.Size(393, 21)
        Me.cmbFindWhat.Location = New System.Drawing.Point(16, 16)
        Me.cmbFindWhat.TabIndex = 4
        Me.cmbFindWhat.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbFindWhat.BackColor = System.Drawing.SystemColors.Window
        Me.cmbFindWhat.CausesValidation = True
        Me.cmbFindWhat.Enabled = True
        Me.cmbFindWhat.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cmbFindWhat.IntegralHeight = True
        Me.cmbFindWhat.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbFindWhat.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmbFindWhat.Sorted = False
        Me.cmbFindWhat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
        Me.cmbFindWhat.TabStop = True
        Me.cmbFindWhat.Visible = True
        Me.cmbFindWhat.Name = "cmbFindWhat"
        Me.lstOutput.Size = New System.Drawing.Size(385, 254)
        Me.lstOutput.Location = New System.Drawing.Point(24, 224)
        Me.lstOutput.TabIndex = 3
        Me.lstOutput.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstOutput.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lstOutput.BackColor = System.Drawing.SystemColors.Window
        Me.lstOutput.CausesValidation = True
        Me.lstOutput.Enabled = True
        Me.lstOutput.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lstOutput.IntegralHeight = True
        Me.lstOutput.Cursor = System.Windows.Forms.Cursors.Default
        Me.lstOutput.SelectionMode = System.Windows.Forms.SelectionMode.One
        Me.lstOutput.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lstOutput.Sorted = False
        Me.lstOutput.TabStop = True
        Me.lstOutput.Visible = True
        Me.lstOutput.MultiColumn = False
        Me.lstOutput.Name = "lstOutput"
        Me.cmdGoTo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.cmdGoTo.Text = "Go To"
        Me.cmdGoTo.Size = New System.Drawing.Size(113, 41)
        Me.cmdGoTo.Location = New System.Drawing.Point(152, 144)
        Me.cmdGoTo.TabIndex = 2
        Me.cmdGoTo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdGoTo.BackColor = System.Drawing.SystemColors.Control
        Me.cmdGoTo.CausesValidation = True
        Me.cmdGoTo.Enabled = True
        Me.cmdGoTo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdGoTo.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdGoTo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdGoTo.TabStop = True
        Me.cmdGoTo.Name = "cmdGoTo"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.Size = New System.Drawing.Size(113, 41)
        Me.cmdCancel.Location = New System.Drawing.Point(304, 144)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.CausesValidation = True
        Me.cmdCancel.Enabled = True
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.TabStop = True
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdFind.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.cmdFind.Text = "Find"
        Me.cmdFind.Size = New System.Drawing.Size(113, 41)
        Me.cmdFind.Location = New System.Drawing.Point(16, 144)
        Me.cmdFind.TabIndex = 0
        Me.cmdFind.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdFind.BackColor = System.Drawing.SystemColors.Control
        Me.cmdFind.CausesValidation = True
        Me.cmdFind.Enabled = True
        Me.cmdFind.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdFind.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdFind.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdFind.TabStop = True
        Me.cmdFind.Name = "cmdFind"
        Me.Controls.Add(cmbFindWhat)
        Me.Controls.Add(lstOutput)
        Me.Controls.Add(cmdGoTo)
        Me.Controls.Add(cmdCancel)
        Me.Controls.Add(cmdFind)
    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As frmGlobalFind
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As frmGlobalFind
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New frmGlobalFind()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal value As frmGlobalFind)
            m_vb6FormDefInstance = Value
        End Set
    End Property
#End Region

    Public oHostApp As Excel.Application


    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdFind_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdFind.Click
        Dim wbk As Excel.Workbook
        wbk = oHostApp.ActiveWorkbook
        Dim sht As Excel.Worksheet
        Dim foundAddress As Object

        Dim searchPattern As String

        searchPattern = Me.cmbFindWhat.Text
        g_mostRecentSearchPattern = searchPattern


        If Not g_recentSearchPatterns.Contains(searchPattern) Then
            'Call g_recentSearchPatterns.Add("", searchPattern)
            g_recentSearchPatterns.Add(searchPattern, "")
        End If

        UpdateMRUDisplay()
        Me.lstOutput.Items.Clear()

        For Each sht In wbk.Worksheets
            For Each foundAddress In MatchingCells(Me.cmbFindWhat.Text, sht.Cells)
                'UPGRADE_WARNING: Couldn't resolve default property of object foundAddress(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                Me.lstOutput.Items.Add(foundAddress.Key)
                'Me.txtOutput.Text = Me.txtOutput.Text & foundAddress & vbCrLf
            Next foundAddress
        Next sht

        If Me.lstOutput.Items.Count > 0 Then
            Me.lstOutput.Focus()
            Me.lstOutput.SetSelected(0, True)
            cmdGoTo_Click(cmdGoTo, New System.EventArgs())
        End If

    End Sub


    Private Function MatchingCells(ByVal ptrn As String, ByRef rge As Excel.Range) As System.Collections.SortedList

        Dim ret As System.Collections.SortedList
        ret = New System.Collections.SortedList

        Dim r As Excel.Range
        r = rge.Find(what:=ptrn, after:=Nothing, lookIn:=Excel.Enums.XlFindLookIn.xlFormulas)
        '' r = rge.Find(what:=ptrn, lookIn:=Excel.Enums.XlFindLookIn.xlFormulas, MatchCase:=False)
        Dim firstAddress As String


        If r Is Nothing Then
            MatchingCells = ret
            Exit Function
        End If

        firstAddress = r.Address
        Dim k As String
        'UPGRADE_WARNING: Couldn't resolve default property of object r.Parent.Name. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'

        k = "'" & r.Parent.Name & "'!" & r.Address
        If Not ret.Contains(k) Then ret.Add(k, "")

        Do
            r = rge.FindNext(r)
            'UPGRADE_WARNING: Couldn't resolve default property of object r.Parent.Name. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            '' If Not r Is Nothing Then ret.Add("", "'" & r.Parent.Name & "'!" & r.Address)
            If Not r Is Nothing Then
                k = "'" & r.Parent.Name & "'!" & r.Address
                If Not ret.Contains(k) Then ret.Add(k, "")
            End If

        Loop While Not r Is Nothing And r.Address <> firstAddress

        MatchingCells = ret

    End Function
	
	Private Sub cmdGoTo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGoTo.Click
		On Error Resume Next
		Dim r As Excel.Range
		If Me.lstOutput.SelectedIndex >= 0 Then
			r = oHostApp.Range(Me.lstOutput.Text)
			oHostApp.Goto(Reference:=AdjustReference(r), Scroll:=True)
			r.Select()
		End If
	End Sub
	
	
	Private Function AdjustReference(ByRef r As Excel.Range) As Excel.Range
		Dim ret As Excel.Range
		On Error Resume Next
		Dim c As Integer
		ret = r
		For c = 1 To 6
			ret = ret.Offset(-1, 0)
		Next 
		For c = 1 To 3
			ret = ret.Offset(0, -1)
		Next 
		
ehandler: 
		AdjustReference = ret
	End Function
	
	
	
	'UPGRADE_WARNING: Form event frmGlobalFind.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2065"'
	Private Sub frmGlobalFind_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		Me.cmbFindWhat.Focus()
	End Sub
	
	Private Sub frmGlobalFind_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Me.Text = "Find in " & oHostApp.ActiveWorkbook.Name
		
		UpdateMRUDisplay()
		''Me.cmbFindWhat.SetFocus   '' ???
	End Sub
	
	Private Sub UpdateMRUDisplay()
		cmbFindWhat.Items.Clear()
		cmbFindWhat.Text = g_mostRecentSearchPattern
        Dim sp As Object
		For	Each sp In g_recentSearchPatterns
			'UPGRADE_WARNING: Couldn't resolve default property of object sp(). Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
            cmbFindWhat.Items.Insert(cmbFindWhat.Items.Count, sp.Key)
		Next sp
	End Sub
	
	'UPGRADE_WARNING: Event cmbFindWhat.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2075"'
	Private Sub cmbFindWhat_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbFindWhat.SelectedIndexChanged
		''Me.cmbFindWhat.Text = cmbFindWhat.Text
		'Me.cmbFindWhat.Text = Me.cmbFindWhat.SelText
		cmdFind_Click(cmdFind, New System.EventArgs())
	End Sub
	
	'UPGRADE_WARNING: ComboBox Event cmbFindWhat.DblClick was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2050"'
	Private Sub cmbFindWhat_DblClick()
		''Me.cmbFindWhat.Text = cmbFindWhat.Text
		'Me.cmbFindWhat.Text = Me.cmbFindWhat.SelText
		cmdFind_Click(cmdFind, New System.EventArgs())
	End Sub
	
	Private Sub cmbFindWhat_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmbFindWhat.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		''
		'' Debug.Print KeyAscii
		If KeyAscii = 13 Then
			cmdFind_Click(cmdFind, New System.EventArgs())
		End If
		
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	
	Private Sub lstOutput_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstOutput.DoubleClick
		cmdGoTo_Click(cmdGoTo, New System.EventArgs())
	End Sub
	
	Private Sub lstOutput_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles lstOutput.KeyUp
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		Select Case KeyCode
			Case 40
				cmdGoTo_Click(cmdGoTo, New System.EventArgs())
			Case 38
				cmdGoTo_Click(cmdGoTo, New System.EventArgs())
			Case 13
				cmdGoTo_Click(cmdGoTo, New System.EventArgs())
		End Select
		
	End Sub
	
	'UPGRADE_WARNING: ListBox Event lstOutput.Scroll was not upgraded. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup2050"'
	Private Sub lstOutput_Scroll()
		''
		System.Diagnostics.Debug.WriteLine(Me.lstOutput.SelectedIndex)
	End Sub
End Class