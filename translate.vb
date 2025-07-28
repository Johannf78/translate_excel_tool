Sub BulkTranslateInTargetWorkbook()
    Dim targetWorkbook As Workbook
    Dim targetSheet As Worksheet
    Dim targetChart As ChartObject
    Dim translationTable As ListObject
    Dim cell As Range
    Dim findText As String
    Dim replaceText As String
    Dim lastRow As Long
    Dim rng As Range
    Dim replacementsMade As Long
    
    ' Split the file path and file name
    Dim targetFileName As String
    Dim targetPath As String
    Dim targetFilePath As String

    ' Get targetPath from named cell
    On Error Resume Next
    targetPath = ThisWorkbook.Names("targetPath").RefersToRange.Value
    If Err.Number <> 0 Then
        MsgBox "Error: Named cell 'targetPath' not found. Please create a named cell called 'targetPath' with the file path.", vbCritical
        Exit Sub
    End If
    
    targetFileName = ThisWorkbook.Names("targetFileName").RefersToRange.Value
    If Err.Number <> 0 Then
        MsgBox "Error: Named cell 'targetFileName' not found. Please create a named cell called 'targetFileName' with the file name.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Combine path and file name
    targetFilePath = targetPath & targetFileName

    ' Show confirmation dialog with file details
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox("Translation will be performed on:" & vbCrLf & vbCrLf & _
                         "File: " & targetFileName & vbCrLf & _
                         "Path: " & targetPath & vbCrLf & vbCrLf & _
                         "Do you want to proceed with the translation?", _
                         vbYesNo + vbQuestion, "Confirm Translation")
    
    ' Check if user wants to cancel
    If userResponse = vbNo Then
        MsgBox "Translation cancelled by user.", vbInformation, "Cancelled"
        Exit Sub
    End If
    
    ' Show additional message about processing time
    MsgBox "Please be patient, it can take up to a few minutes.", vbInformation, "Processing Time"

    ' Optimize Performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Open the target workbook
    Set targetWorkbook = Workbooks.Open(targetFilePath)

    ' Get the translation table by name
    Set translationTable = ThisWorkbook.Sheets("German").ListObjects("Translations_EN_to_DE")

    ' OPTIMIZATION 1: Pre-load all translations into arrays for faster access
    Dim translationData As Variant
    Dim findTexts() As String
    Dim replaceTexts() As String
    Dim translationCount As Long
    Dim i As Long
    
    ' Load all translation data into array
    translationData = translationTable.DataBodyRange.Value
    
    ' Count valid translations
    translationCount = 0
    For i = 1 To UBound(translationData, 1)
        If translationData(i, 1) <> "" And translationData(i, 2) <> "" And _
           Not IsError(translationData(i, 1)) And Not IsError(translationData(i, 2)) Then
            translationCount = translationCount + 1
        End If
    Next i
    
    ' Resize arrays to exact size needed
    ReDim findTexts(1 To translationCount)
    ReDim replaceTexts(1 To translationCount)
    
    ' Fill arrays with valid translations
    Dim arrayIndex As Long
    arrayIndex = 1
    For i = 1 To UBound(translationData, 1)
        If translationData(i, 1) <> "" And translationData(i, 2) <> "" And _
           Not IsError(translationData(i, 1)) And Not IsError(translationData(i, 2)) Then
            findTexts(arrayIndex) = CStr(translationData(i, 1))
            replaceTexts(arrayIndex) = CStr(translationData(i, 2))
            arrayIndex = arrayIndex + 1
        End If
    Next i

    ' Initialize replacement counter
    replacementsMade = 0

    ' OPTIMIZATION 2: Process all translations for each element type at once
    For Each targetSheet In targetWorkbook.Sheets
        ' Translate sheet name first
        Dim originalSheetName As String
        originalSheetName = targetSheet.Name
        Dim newSheetName As String
        newSheetName = originalSheetName
        
        ' Apply all translations to sheet name at once
        For i = 1 To translationCount
            If InStr(1, newSheetName, findTexts(i), vbTextCompare) > 0 Then
                newSheetName = Replace(newSheetName, findTexts(i), replaceTexts(i))
            End If
        Next i
        
        ' Only rename if the name actually changed
        If newSheetName <> originalSheetName Then
            On Error Resume Next
            targetSheet.Name = newSheetName
            If Err.Number = 0 Then
                replacementsMade = replacementsMade + 1
                Application.StatusBar = "Sheet renamed: " & originalSheetName & " -> " & newSheetName
            End If
            On Error GoTo 0
        End If
        
        ' OPTIMIZATION 3: Process cell content with single loop through translations
        Set rng = targetSheet.UsedRange
        If Not rng Is Nothing Then
            ' Check if current range is part of an excluded table
            Dim shouldSkipRange As Boolean
            shouldSkipRange = False
            
            ' List of tables to exclude from translation
            Dim excludedTables As Variant
            excludedTables = Array("Table_relevant_data", "query_table_relavant_data", "Table_critical_data")
            
            ' Check if current range overlaps with any excluded table
            Dim table As ListObject
            For Each table In targetSheet.ListObjects
                Dim tableName As String
                tableName = table.Name
                
                ' Check if this table should be excluded
                Dim j As Long
                For j = LBound(excludedTables) To UBound(excludedTables)
                    If tableName = excludedTables(j) Then
                        ' Check if current range overlaps with excluded table
                        If Not Intersect(rng, table.Range) Is Nothing Then
                            shouldSkipRange = True
                            Application.StatusBar = "Skipping excluded table: " & tableName
                            Exit For
                        End If
                    End If
                Next j
                
                If shouldSkipRange Then Exit For
            Next table
            
            ' Only process range if it's not excluded
            If Not shouldSkipRange Then
                For i = 1 To translationCount
                    On Error Resume Next
                    If rng.Replace(What:=findTexts(i), Replacement:=replaceTexts(i), LookAt:=xlPart) = True Then
                        replacementsMade = replacementsMade + 1
                        Application.StatusBar = "Replacements made: " & replacementsMade
                    End If
                    On Error GoTo 0
                Next i
            End If
        End If
        
        ' OPTIMIZATION 4: Process chart titles efficiently
        If targetSheet.ChartObjects.Count > 0 Then
            For Each targetChart In targetSheet.ChartObjects
                On Error Resume Next
                If targetChart.Chart.HasTitle Then
                    If Not targetChart.Chart.ChartTitle Is Nothing Then
                        Dim chartTitle As String
                        chartTitle = targetChart.Chart.ChartTitle.Text
                        Dim newChartTitle As String
                        newChartTitle = chartTitle
                        
                        ' Apply all translations to chart title at once
                        For i = 1 To translationCount
                            If InStr(1, newChartTitle, findTexts(i), vbTextCompare) > 0 Then
                                newChartTitle = Replace(newChartTitle, findTexts(i), replaceTexts(i))
                            End If
                        Next i
                        
                        If newChartTitle <> chartTitle Then
                            targetChart.Chart.ChartTitle.Text = newChartTitle
                            replacementsMade = replacementsMade + 1
                            Application.StatusBar = "Chart title updated: " & replacementsMade
                        End If
                    End If
                End If
                On Error GoTo 0
            Next targetChart
        End If
        
        ' OPTIMIZATION 5: Process shapes efficiently
        If targetSheet.Shapes.Count > 0 Then
            Dim shape As Shape
            For Each shape In targetSheet.Shapes
                On Error Resume Next
                If shape.TextFrame.HasText Then
                    If Not shape.TextFrame.Characters Is Nothing Then
                        Dim shapeText As String
                        shapeText = shape.TextFrame.Characters.Text
                        Dim newShapeText As String
                        newShapeText = shapeText
                        
                        ' Apply all translations to shape text at once
                        For i = 1 To translationCount
                            If InStr(1, newShapeText, findTexts(i), vbTextCompare) > 0 Then
                                newShapeText = Replace(newShapeText, findTexts(i), replaceTexts(i))
                            End If
                        Next i
                        
                        If newShapeText <> shapeText Then
                            shape.TextFrame.Characters.Text = newShapeText
                            replacementsMade = replacementsMade + 1
                            Application.StatusBar = "Text box updated: " & replacementsMade
                        End If
                    End If
                End If
                On Error GoTo 0
            Next shape
        End If
        
        ' Allow Excel to update UI (reduced frequency)
        If replacementsMade Mod 10 = 0 Then
            DoEvents
        End If
    Next targetSheet

    ' Save & Close the workbook
    targetWorkbook.Save
    targetWorkbook.Close False

    ' Restore Excel Settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False ' Reset status bar

    ' Clean up
    Set targetWorkbook = Nothing
    MsgBox "Translation completed successfully! Total replacements: " & replacementsMade, vbInformation
End Sub
