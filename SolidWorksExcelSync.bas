Attribute VB_Name = "SolidWorksExcelSync"
Option Explicit

Private Const xlUp As Long = -4162
Private Const xlOpenXMLWorkbook As Long = 51
Private Const swDocPART As Long = 1
Private Const swSaveAsOptions_Silent As Long = 1

Public Sub main()
    Dim swApp As Object
    Dim swModel As Object
    Dim swAssembly As SldWorks.AssemblyDoc
    Dim excelApp As Object
    Dim workbook As Object
    Dim sheet As Object
    Dim workbookPath As String
    Dim assemblyPath As String
    Dim assemblyFolder As String
    Dim templatePath As String
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim partNumber As String
    Dim partPath As String
    Dim componentName As String
    Dim failureMessage As String
    Dim importedCount As Long
    Dim failedCount As Long
    Dim isNewWorkbook As Boolean
    Dim saveErrors As Long
    Dim saveWarnings As Long

    On Error GoTo Failed

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "Open and save an assembly before running this macro.", vbExclamation
        Exit Sub
    End If

    If swModel.GetType <> 2 Then
        MsgBox "The active SolidWorks document must be an assembly.", vbExclamation
        Exit Sub
    End If
    Set swAssembly = swModel

    assemblyPath = CStr(swModel.GetPathName)
    If Len(assemblyPath) = 0 Then
        MsgBox "Save the assembly before running this macro.", vbExclamation
        Exit Sub
    End If
    assemblyFolder = Left$(assemblyPath, InStrRev(assemblyPath, "\"))

    workbookPath = InputBox( _
        "Enter the full path of the Excel workbook (.xlsx):", _
        "Assembly / Excel sync", _
        assemblyFolder & "AssemblyParts.xlsx")
    If Len(Trim$(workbookPath)) = 0 Then Exit Sub
    If LCase$(Right$(workbookPath, 5)) <> ".xlsx" Then workbookPath = workbookPath & ".xlsx"

    Set excelApp = CreateObject("Excel.Application")
    If Len(Dir$(workbookPath)) > 0 Then
        Set workbook = excelApp.Workbooks.Open(workbookPath)
    Else
        Set workbook = excelApp.Workbooks.Add
        isNewWorkbook = True
    End If
    Set sheet = workbook.Worksheets(1)
    PrepareWorksheet sheet

    ExportComponents swAssembly, sheet
    lastRow = LastPartRow(sheet)

    For rowIndex = 2 To lastRow
        partNumber = CellText(sheet.Cells(rowIndex, 1).Value)
        componentName = CellText(sheet.Cells(rowIndex, 2).Value)

        If Len(partNumber) > 0 And Len(componentName) = 0 Then
            partPath = CellText(sheet.Cells(rowIndex, 3).Value)
            If ImportPart(swApp, swAssembly, sheet, rowIndex, assemblyFolder, _
                          partNumber, partPath, templatePath, failureMessage) Then
                importedCount = importedCount + 1
            Else
                failedCount = failedCount + 1
                Debug.Print "Excel row " & rowIndex & ": " & failureMessage
            End If
        End If
    Next rowIndex

    If importedCount > 0 Then
        swModel.EditRebuild3
        If Not swModel.Save3(swSaveAsOptions_Silent, saveErrors, saveWarnings) Then
            Err.Raise vbObjectError + 1000, "SolidWorksExcelSync", _
                      "SolidWorks could not save the updated assembly."
        End If
    End If

    ExportComponents swAssembly, sheet
    If isNewWorkbook Then
        workbook.SaveAs workbookPath, xlOpenXMLWorkbook
    Else
        workbook.Save
    End If
    excelApp.Visible = True
    workbook.Activate

    If failedCount = 0 Then
        MsgBox "Sync complete. Imported " & importedCount & " part(s).", vbInformation
    Else
        MsgBox "Sync complete. Imported " & importedCount & " part(s); " & _
               failedCount & " row(s) could not be imported. See the VBA Immediate window.", _
               vbExclamation
    End If
    Exit Sub

Failed:
    If Not excelApp Is Nothing Then excelApp.Visible = True
    MsgBox "Sync failed: " & Err.Description, vbCritical
End Sub

Private Sub PrepareWorksheet(ByVal sheet As Object)
    Dim headers As Variant
    Dim columnIndex As Long
    Dim firstCell As String

    headers = Array("Part Number", "Component", "Part File", "Drawing File", "Configuration")
    firstCell = CellText(sheet.Cells(1, 1).Value)

    If Len(firstCell) = 0 Then
        For columnIndex = LBound(headers) To UBound(headers)
            sheet.Cells(1, columnIndex + 1).Value = headers(columnIndex)
        Next columnIndex
        sheet.Rows(1).Font.Bold = True
        Exit Sub
    End If

    For columnIndex = LBound(headers) To UBound(headers)
        If StrComp(CellText(sheet.Cells(1, columnIndex + 1).Value), _
                   CStr(headers(columnIndex)), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 1001, "SolidWorksExcelSync", _
                      "The first worksheet must use the documented column headers in columns A to E."
        End If
    Next columnIndex
End Sub

Private Sub ExportComponents(ByVal model As Object, ByVal sheet As Object)
    Dim components As Variant
    Dim component As Object
    Dim partPath As String
    Dim drawingPath As String
    Dim componentName As String
    Dim partNumber As String
    Dim configuration As String
    Dim targetRow As Long

    components = model.GetComponents(False)
    If IsEmpty(components) Then Exit Sub

    For Each component In components
        partPath = CStr(component.GetPathName)
        If LCase$(Right$(partPath, 7)) = ".sldprt" Then
            componentName = CStr(component.Name2)
            configuration = CStr(component.ReferencedConfiguration)
            partNumber = FileNameWithoutExtension(partPath)
            drawingPath = Left$(partPath, Len(partPath) - 7) & ".SLDDRW"
            If Len(Dir$(drawingPath)) = 0 Then drawingPath = vbNullString

            targetRow = FindComponentRow(sheet, componentName, partPath)
            If targetRow = 0 Then targetRow = LastPartRow(sheet) + 1

            sheet.Cells(targetRow, 1).Value = partNumber
            sheet.Cells(targetRow, 2).Value = componentName
            sheet.Cells(targetRow, 3).Value = partPath
            sheet.Cells(targetRow, 4).Value = drawingPath
            sheet.Cells(targetRow, 5).Value = configuration
        End If
    Next component
End Sub

Private Function ImportPart(ByVal swApp As Object, ByVal model As Object, _
                             ByVal sheet As Object, ByVal rowIndex As Long, _
                             ByVal assemblyFolder As String, ByVal partNumber As String, _
                             ByRef partPath As String, ByRef templatePath As String, _
                             ByRef failureMessage As String) As Boolean
    Dim partDoc As Object
    Dim component As Object
    Dim saved As Boolean
    Dim saveErrors As Long
    Dim saveWarnings As Long

    On Error GoTo ImportFailed

    If Len(partPath) = 0 Then
        partPath = assemblyFolder & SafeFileName(partNumber) & ".SLDPRT"
        If Len(Dir$(partPath)) = 0 Then
            If Len(templatePath) = 0 Then
                templatePath = CStr(swApp.GetDocumentTemplate(swDocPART, "", 0, 0#, 0#))
                If Len(Dir$(templatePath)) = 0 Then
                    templatePath = InputBox( _
                        "Enter the full path of a SolidWorks part template (.prtdot):", _
                        "Part template")
                End If
                If Len(templatePath) = 0 Or Len(Dir$(templatePath)) = 0 Then
                    Err.Raise vbObjectError + 1002, "SolidWorksExcelSync", _
                              "A valid part template is required to create an empty part."
                End If
            End If

            Set partDoc = swApp.NewDocument(templatePath, 0, 0, 0)
            If partDoc Is Nothing Then
                Err.Raise vbObjectError + 1003, "SolidWorksExcelSync", _
                          "SolidWorks could not create a part from the selected template."
            End If
            saved = partDoc.Extension.SaveAs3(partPath, 0, swSaveAsOptions_Silent, _
                                               Nothing, Nothing, saveErrors, saveWarnings)
            swApp.CloseDoc partDoc.GetTitle
            If Not saved Then
                Err.Raise vbObjectError + 1004, "SolidWorksExcelSync", _
                          "SolidWorks could not save the new part: " & partPath
            End If
        End If
    End If

    If LCase$(Right$(partPath, 7)) <> ".sldprt" Or Len(Dir$(partPath)) = 0 Then
        Err.Raise vbObjectError + 1005, "SolidWorksExcelSync", _
                  "The part file does not exist or is not an .SLDPRT file: " & partPath
    End If

    Set component = model.AddComponent5(partPath, 0, "", True, "", 0#, 0#, 0#)
    If component Is Nothing Then
        Err.Raise vbObjectError + 1006, "SolidWorksExcelSync", _
                  "SolidWorks could not insert the part: " & partPath
    End If

    sheet.Cells(rowIndex, 2).Value = CStr(component.Name2)
    sheet.Cells(rowIndex, 3).Value = CStr(component.GetPathName)
    sheet.Cells(rowIndex, 4).Value = DrawingPathFor(partPath)
    sheet.Cells(rowIndex, 5).Value = CStr(component.ReferencedConfiguration)
    ImportPart = True
    Exit Function

ImportFailed:
    failureMessage = Err.Description
End Function

Private Function FindComponentRow(ByVal sheet As Object, ByVal componentName As String, _
                                  ByVal partPath As String) As Long
    Dim rowIndex As Long
    Dim lastRow As Long

    lastRow = LastPartRow(sheet)
    For rowIndex = 2 To lastRow
        If StrComp(CellText(sheet.Cells(rowIndex, 2).Value), componentName, vbTextCompare) = 0 Then
            FindComponentRow = rowIndex
            Exit Function
        End If
    Next rowIndex

    For rowIndex = 2 To lastRow
        If Len(CellText(sheet.Cells(rowIndex, 2).Value)) = 0 And _
           StrComp(CellText(sheet.Cells(rowIndex, 3).Value), partPath, vbTextCompare) = 0 Then
            FindComponentRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function LastPartRow(ByVal sheet As Object) As Long
    LastPartRow = sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
    If LastPartRow < 1 Then LastPartRow = 1
End Function

Private Function CellText(ByVal value As Variant) As String
    If Not IsError(value) And Not IsNull(value) Then
        CellText = Trim$(CStr(value))
    End If
End Function

Private Function FileNameWithoutExtension(ByVal filePath As String) As String
    Dim fileName As String
    Dim separator As Long
    Dim extension As Long

    separator = InStrRev(filePath, "\")
    fileName = Mid$(filePath, separator + 1)
    extension = InStrRev(fileName, ".")
    If extension > 0 Then fileName = Left$(fileName, extension - 1)
    FileNameWithoutExtension = fileName
End Function

Private Function SafeFileName(ByVal fileName As String) As String
    Dim invalidCharacters As Variant
    Dim character As Variant

    invalidCharacters = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each character In invalidCharacters
        fileName = Replace$(fileName, CStr(character), "_")
    Next character
    fileName = Trim$(fileName)
    Do While Len(fileName) > 0 And (Right$(fileName, 1) = "." Or Right$(fileName, 1) = " ")
        fileName = Left$(fileName, Len(fileName) - 1)
    Loop
    If Len(fileName) = 0 Then
        Err.Raise vbObjectError + 1007, "SolidWorksExcelSync", _
                  "The part number does not contain a valid filename."
    End If
    SafeFileName = fileName
End Function

Private Function DrawingPathFor(ByVal partPath As String) As String
    Dim drawingPath As String

    drawingPath = Left$(partPath, Len(partPath) - 7) & ".SLDDRW"
    If Len(Dir$(drawingPath)) > 0 Then DrawingPathFor = drawingPath
End Function
