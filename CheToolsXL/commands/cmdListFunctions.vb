Imports Office = NetOffice.OfficeApi
Imports Excel = NetOffice.ExcelApi
Imports System.IO

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
            fname = System.IO.Path.ChangeExtension(fname, "txt")

            fname = m_host.GetSaveAsFilename(initialFilename:=fname, fileFilter:="Text
            Files (*.txt), *.txt")
           
			If VarType(fname) = VariantType.String Then
                
                Dim file As System.IO.StreamWriter
                file = My.Computer.FileSystem.OpenTextFileWriter(fname, False)
                For Each ff As String In fcnNames.Keys
                    file.WriteLine(ff)
                Next
                file.Close()
            End If
        End If

        Exit Sub

ehandler:

        MsgBox(Err.Description)

    End Sub

    Sub GetNamesOfFunctions(ByVal formulaString As String, ByRef fcnNames As System.Collections.Generic.SortedDictionary(Of String, Int16))

        Dim pos As Integer
        Dim c As String

        Dim currentFcnName As String
        currentFcnName = ""

        For pos = 1 To Len(formulaString)

            c = Mid(formulaString, pos, 1)

            If Strings.InStr("+-*/ =:;<>&", c) > 0 Then
                currentFcnName = ""
            Else

                Select Case c

                    Case "("
                        If (Len(currentFcnName) > 0) Then
                            fcnNames.Item(LCase(currentFcnName)) = 0
                            currentFcnName = ""
                        End If
                    Case ")"
                        currentFcnName = ""
                    Case ","
                        currentFcnName = ""
                    Case "|"
                        If (Len(currentFcnName) > 0) Then
                            fcnNames.Item(LCase(currentFcnName)) = 0
                            currentFcnName = ""
                        End If
                    Case Else
                        currentFcnName = currentFcnName & c

                End Select

            End If

        Next  ''pos

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

                    Dim ha As Object
                    ha = r2.HasArray
                    If Not TypeOf ha Is Boolean Then
                        ha = False
                    End If
                    If ha Then
                        formulaString = r2.CurrentArray.FormulaArray
                        If InStr(formulaString, "(") Then Call GetNamesOfFunctions(formulaString, ret)

                    Else
                        Dim r3 As Excel.Range
                        For i2 = 1 To r2.Cells.Count
                            r3 = r2.Cells.Item(i2)
                            formulaString = r3.Formula
                            If InStr(formulaString, "(") Then Call GetNamesOfFunctions(formulaString, ret)

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
    