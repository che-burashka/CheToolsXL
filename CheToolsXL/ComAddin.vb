Imports Microsoft.Win32
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

Imports NetOffice
Imports Excel = NetOffice.ExcelApi
Imports NetOffice.ExcelApi.Enums
Imports Office = NetOffice.OfficeApi
Imports NetOffice.OfficeApi.Enums

Imports Extensibility
Imports System.Windows.Forms
Imports System.Reflection


<GuidAttribute("BBF09D96-2350-46FC-A213-5C32E813BA51"), ProgIdAttribute("CheToolsXL.ComAddin"), ComVisible(True)>
Public Class Addin
    Implements IDTExtensibility2

    Private Shared ReadOnly _addinOfficeRegistryKey As String = "Software\\Microsoft\\Office\\Excel\\AddIns\\"
    Private Shared ReadOnly _progId As String = "CheToolsXL.ComAddin"
    Private Shared ReadOnly _addinFriendlyName As String = "Che Tools for Excel"
    Private Shared ReadOnly _addinDescription As String = "Che Tools for Excel"

    ' gui elements
    Private Shared ReadOnly _toolbarName = "Che Toolbar"
    Private Shared ReadOnly _toolbarButtonName As String = "Che ToolbarButton"
    Private Shared ReadOnly _toolbarPopupName As String = "Che ToolbarPopup"
    Private Shared ReadOnly _menuName As String = "Che Menu"
    Private Shared ReadOnly _menuButtonName As String = "Che Button"
    Private Shared ReadOnly _contextName As String = "Che ContextMenu"
    Private Shared ReadOnly _contextMenuButtonName As String = "Che ContextButton"

    Private WithEvents m_host As Excel.Application
    Private m_menuTag As String
    Private m_root As Office.CommandBarPopup
    Private m_commands As New Collection

    Public Sub New()
        MyBase.New()
        m_menuTag = "CheAddinMenuTag"
    End Sub

#Region "IDTExtensibility2 Members"

    Public Sub OnConnection(ByVal Application As Object, ByVal ConnectMode As ext_ConnectMode, ByVal AddInInst As Object, ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection

        Try

            m_host = New Excel.Application(Nothing, Application)

            'm_root = m_host.CommandBars.Item("Worksheet Menu Bar").Controls.Add(MsoControlType.msoControlPopup, m_menuTag, "", True)
            m_root = m_host.CommandBars.Item("Worksheet Menu Bar").Controls.Add(MsoControlType.msoControlPopup)
            m_root.Caption = "Che Addin for XL"
            m_root.Visible = True

            buildMenu()


        Catch ex As Exception

            Dim message As String = String.Format("An error occured.{0}{0}{1}", Environment.NewLine, ex.Message)
            MessageBox.Show(message, _progId, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Public Sub OnDisconnection(ByVal RemoveMode As ext_DisconnectMode, ByRef custom As System.Array) Implements IDTExtensibility2.OnDisconnection

        Try
            removeRec(m_root)
            If (Not IsNothing(m_host)) Then
                m_host.Dispose()
            End If

        Catch ex As Exception

            Dim message As String = String.Format("An error occured.{0}{0}{1}", Environment.NewLine, ex.Message)
            MessageBox.Show(message, _progId, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Public Sub OnStartupComplete(ByRef custom As System.Array) Implements IDTExtensibility2.OnStartupComplete

        Try

            ''CreateTemporaryUserInterface()

        Catch ex As Exception

            Dim message As String = String.Format("An error occured.{0}{0}{1}", Environment.NewLine, ex.Message)
            MessageBox.Show(message, _progId, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements IDTExtensibility2.OnAddInsUpdate

    End Sub

    Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements IDTExtensibility2.OnBeginShutdown

    End Sub

#End Region

#Region "COM Register Functions"

    <ComRegisterFunctionAttribute()>
    Public Shared Sub RegisterFunction(ByVal type As Type)
        Try

            ' add codebase value
            Dim thisAssembly As Assembly = Assembly.GetAssembly(GetType(Addin))
            Dim key As RegistryKey = Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\InprocServer32\1.0.0.0")
            key.SetValue("CodeBase", thisAssembly.CodeBase)
            key.Close()

            Registry.ClassesRoot.CreateSubKey("CLSID\{" + type.GUID.ToString().ToUpper() + "}\Programmable")

            ' add bypass key
            ' http://support.microsoft.com/kb/948461
            key = Registry.ClassesRoot.CreateSubKey("Interface\\{000C0601-0000-0000-C000-000000000046}")
            Dim defaultValue As String = key.GetValue("")
            If (IsNothing(defaultValue)) Then
                key.SetValue("", "Office .NET Framework Lockback Bypass Key")
            End If
            key.Close()

            ' add excel addin key
            Registry.CurrentUser.CreateSubKey(_addinOfficeRegistryKey + _progId)
            Dim rk As RegistryKey = Registry.CurrentUser.OpenSubKey(_addinOfficeRegistryKey + _progId, True)
            rk.SetValue("LoadBehavior", CInt(3))
            rk.SetValue("FriendlyName", _addinFriendlyName)
            rk.SetValue("Description", _addinDescription)
            rk.Close()

        Catch ex As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine)
            MessageBox.Show("An error occured." + details, "Register " + _progId, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    <ComUnregisterFunctionAttribute()>
    Public Shared Sub UnregisterFunction(ByVal type As Type)
        Try

            Registry.ClassesRoot.DeleteSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\Programmable", False)
            Registry.CurrentUser.DeleteSubKey(_addinOfficeRegistryKey + _progId, False)

        Catch throwedException As Exception

            Dim details As String = String.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine)
            MessageBox.Show("An error occured." + details, "Unregister " + _progId, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

#End Region

#Region "UI Methods"

    ''' <summary>
    ''' creates gui elements
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub CreateTemporaryUserInterface()

        ' How to: Add Commands to Shortcut Menus in Excel
        ' http://msdn.microsoft.com/en-us/library/0batekf4.aspx   

        'create commandbar 
        Dim commandBar As Office.CommandBar = m_host.CommandBars.Add(_toolbarName, MsoBarPosition.msoBarTop, System.Type.Missing, True)
        commandBar.Visible = True

        ' add popup to commandbar
        Dim commandBarPop As Office.CommandBarPopup = commandBar.Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarPop.Caption = _toolbarPopupName
        commandBarPop.Tag = _toolbarPopupName

        'add a button to the popup
        Dim commandBarBtn As Office.CommandBarButton = commandBarPop.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
        commandBarBtn.FaceId = 9
        commandBarBtn.Caption = _toolbarButtonName
        commandBarBtn.Tag = _toolbarButtonName
        Dim clickHandler As NetOffice.OfficeApi.CommandBarButton_ClickEventHandler = AddressOf Me.commandBarBtn_ClickEvent
        AddHandler commandBarBtn.ClickEvent, clickHandler

        ' create menu 
        commandBar = m_host.CommandBars("Worksheet Menu Bar")

        ' add popup to menu bar
        commandBarPop = commandBar.Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarPop.Caption = _menuName
        commandBarPop.Tag = _menuName

        ' add a button to the popup
        commandBarBtn = commandBarPop.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
        commandBarBtn.FaceId = 9
        commandBarBtn.Caption = _menuButtonName
        commandBarBtn.Tag = _menuButtonName
        clickHandler = AddressOf Me.commandBarBtn_ClickEvent
        AddHandler commandBarBtn.ClickEvent, clickHandler

        ' create context menu 
        commandBarPop = m_host.CommandBars("Cell").Controls.Add(MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarPop.Caption = _contextName
        commandBarPop.Tag = _contextName

        ' add a button to the popup
        commandBarBtn = commandBarPop.Controls.Add(MsoControlType.msoControlButton, System.Type.Missing, System.Type.Missing, System.Type.Missing, True)
        commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption
        commandBarBtn.Caption = _contextMenuButtonName
        commandBarBtn.Tag = _contextMenuButtonName
        commandBarBtn.FaceId = 9
        clickHandler = AddressOf Me.commandBarBtn_ClickEvent
        AddHandler commandBarBtn.ClickEvent, clickHandler

    End Sub

#End Region

#Region "UI Trigger"

    ''' <summary>
    ''' Click event trigger from created buttons. incoming call comes from word application thread.
    ''' </summary>
    ''' <param name="Ctrl"></param>
    ''' <param name="CancelDefault"></param>
    ''' <remarks></remarks>
    Private Sub commandBarBtn_ClickEvent(ByVal Ctrl As NetOffice.OfficeApi.CommandBarButton, ByRef CancelDefault As Boolean)

        Try

            Dim message As String = String.Format("Click from Button {0}.", Ctrl.Caption)
            MessageBox.Show(message, _progId, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Ctrl.Dispose()

        Catch ex As Exception

            Dim message As String = String.Format("An error occured.{0}{0}{1}", Environment.NewLine, ex.Message)
            MessageBox.Show(message, _progId, MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

#End Region


    Private Sub removeRec(ByVal ctl As Office.CommandBarControl)

        Dim chld As Office.CommandBarControl
        Dim popup As Office.CommandBarPopup

        Try

            If TypeOf (ctl) Is Office.CommandBarPopup Then
                popup = ctl
                For Each chld In popup.Controls
                    removeRec(chld)
                Next
            End If
            ctl.Delete()
        Catch e As Exception
            '
        End Try
    End Sub


    Private Sub addButton(ByVal mnu As Office.CommandBarPopup, ByVal nm As String, ByVal cmd As ICmd, ByVal fwd As Boolean)

        Dim btn As Office.CommandBarButton
        'btn = mnu.Controls.Add(MsoControlType.msoControlButton, "", m_menuTag, "", True)
        btn = mnu.Controls.Add(MsoControlType.msoControlButton)
        btn.Caption = nm
        btn.Visible = True

        cmd.Init(btn, m_host, fwd)
        Me.m_commands.Add(cmd)

    End Sub

    Private Function addSubMenu(ByVal mnu As Office.CommandBarPopup, ByVal nm As String) As Office.CommandBarPopup

        Dim smnu As Office.CommandBarPopup
        'smnu = mnu.Controls.Add(MsoControlType.msoControlPopup, "", m_menuTag, "", True)
        smnu = mnu.Controls.Add(MsoControlType.msoControlPopup)
        smnu.Caption = nm
        smnu.Visible = True
        addSubMenu = smnu

    End Function


    ' add addin menu here:
    Private Sub buildMenu()
        Dim cm As Office.CommandBarPopup
        cm = m_root

        '' Application level
        addButton(cm, "Show &Filenames", New cmdShowFNames, True)
        addButton(cm, "&Normalize Settings", New cmdSetMySettings, True)
        addButton(cm, "Mark &XL Window", New cmdMarkExcelWindow, True)


        '' Workbook level
        cm = addSubMenu(m_root, "This &Workbook")
        addButton(cm, "&Find &Globally in Workbook", New cmdGlobalFind, True)
        addButton(cm, "&List Functions in Workbook", New cmdListFunctions, True)
        addButton(cm, "&Hide Sheets", New cmdHideSheets, True)
        addButton(cm, "&Unhide Sheets", New cmdHideSheets, False)
        addButton(cm, "Highlight Names", New cmdHighlightNamedRanges, True)
        addButton(cm, "Undo Highlight Names", New cmdHighlightNamedRanges, False)

        '' Sheet level
        ''cm = addSubMenu(m_root, "Sheet")
        ''addButton(cm, "Sheet Dependents", New cmdExtSheetDependants, False)

        '' Selection/ Active Cell level
        cm = addSubMenu(m_root, "&Selection")

        addButton(cm, "&Calc", New cmdCalcSelection, True)
        addButton(cm, "&Freeze", New cmdFreeze, True)
        addButton(cm, "&Thaw", New cmdFreeze, False)


        addButton(cm, "&Paste Special", New cmdPasteSpecial, True)


    End Sub


End Class

