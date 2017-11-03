Imports Office = NetOffice.OfficeApi
Imports Excel = NetOffice.ExcelApi

Public Class cmdListFunctions

    Implements ICmd
    Private WithEvents m_btn As Office.CommandBarButton
    Private m_host As Excel.Application

    Public Sub Init(ByVal btn As Office.CommandBarButton, ByVal host As Object, ByVal fwd As Boolean) Implements ICmd.Init
        m_btn = btn
        m_host = host
    End Sub

    Private Sub clickHandler(ByVal b As Office.CommandBarButton, ByRef cd As Boolean) Handles m_btn.ClickEvent
        Run()
    End Sub


    Sub Run()

        On Error GoTo ehandler

        'Dim fcnNames As New System.Collections.Specialized.StringDictionary

        Dim fcnNames As New System.Collections.Generic.SortedDictionary(Of String, Int16)

        fcnNames = listAllFunctionsInWbk((m_host.ActiveWorkbook))

        frmFunctionsList.DefInstance.lstFunctionNames.Items.Clear()
        frmFunctionsList.DefInstance.doSave = False

        Dim f As Object
        For Each f In fcnNames
            frmFunctionsList.DefInstance.lstFunctionNames.Items.Add(f.Key)
        Next f

        frmFunctionsList.DefInstance.Text = "Functions used in " & m_host.ActiveWorkbook.Name
        frmFunctionsList.DefInstance.ShowDialog()

        Dim fname As String
        If frmFunctionsList.DefInstance.doSave Then
            fname = m_host.GetSaveAsFilename(initialFilename:="functions.txt", fileFilter:="Text Files (*.txt), *.txt")
            If VarType(fname) = VariantType.String Then
                '' todo: save to file here
            End If
        End If

        Exit Sub

ehandler:

        MsgBox(Err.Description)

    End Sub

    Function listAllFunctionsInWbk(ByRef wbk As Excel.Workbook) As System.Collections.Generic.SortedDictionary(Of String, Int16)

        Dim ret As New System.Collections.Generic.SortedDictionary(Of String, Int16)
        
        Dim r1 As Excel.Range
        'Dim r2 As Excel.Range
        'Dim r3 As Excel.Range

        Dim formulaString As String
        Dim i1 As Integer : Dim i2 As Integer

        Dim sht As Excel.Worksheet

        For Each sht In wbk.Worksheets

            r1 = CellsWithFormulas(sht)

            If Not r1 Is Nothing Then

                For i1 = 1 To r1.Areas.Count

                    Dim r2 As Excel.Range
                    r2 = r1.Areas.Item(i1)

                    If r2.HasArray Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object r2.CurrentArray.FormulaArray. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                        formulaString = r2.CurrentArray.FormulaArray
                        'If InStr(formulaString, "(") Then Call GetNamesOfFunctions(formulaString, ret)
                    Else
                        Dim r3 As Excel.Range
                        For i2 = 1 To r2.Cells.Count
                            r3 = r2.Cells.Item(i2)
                            'UPGRADE_WARNING: Couldn't resolve default property of object r3.Formula. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
                            formulaString = r3.Formula
                            'If InStr(formulaString, "(") Then Call GetNamesOfFunctions(formulaString, ret)
                        Next
                    End If
                Next i1
            End If
        Next

        listAllFunctionsInWbk = ret

    End Function

    Function CellsWithFormulas(ByRef sht As Excel.Worksheet) As Excel.Range
        On Error GoTo ehandler
        CellsWithFormulas = sht.Cells.SpecialCells(Excel.Enums.XlCellType.xlCellTypeFormulas)
ehandler:
        Err.Clear()
    End Function

End Class
    