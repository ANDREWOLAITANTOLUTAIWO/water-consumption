Attribute VB_Name = "Module_Dry"
Option Explicit

' =========================================================================
' CORE CALCULATIONS
' =========================================================================

Public Function CalculateDrySeason(ByVal I As Double, ByVal R As Double, ByVal t As Double, ByVal t As Double, ByVal A As Double, ByVal W As Double, ByVal d As Double, ByVal h As Double) As Double
    ' Dry Season Formula: V = 98.1 + 0.0003I + 5.31R + 1.080T - 2.01t - 0.0003A + 0.0804W + 0.0142d - 0.009h
    
    Dim V As Double
    V = 98.1 + (0.0003 * I) + (5.31 * R) + (1.08 * t) - (2.01 * t) - (0.0003 * A) + (0.0804 * W) + (0.0142 * d) - (0.009 * h)
    
    CalculateDrySeason = V
End Function

Public Function CalculateWetSeason(ByVal I As Double, ByVal S As Double, ByVal R As Double, ByVal t As Double, ByVal t As Double, ByVal A As Double, ByVal W As Double, ByVal d As Double, ByVal h As Double) As Double
    ' Wet Season Formula: V = 15.4 + 0.0003I + 5.24S + 0.108R + 4.43T - 2.03t + 0.0003A + 0.0495W + 0.0012d - 0.007h
    
    Dim V As Double
    V = 15.4 + (0.0003 * I) + (5.24 * S) + (0.108 * R) + (4.43 * t) - (2.03 * t) + (0.0003 * A) + (0.0495 * W) + (0.0012 * d) - (0.007 * h)
    
    CalculateWetSeason = V
End Function

' =========================================================================
' LOCALE-AWARE DATA HANDLING
' =========================================================================

Public Function SafeCDbl(ByVal ValueString As String) As Double
    ' Ensures the string is converted to Double using the dot (.) as the decimal separator,
    ' regardless of the local machine's regional settings.
    ' 1. Replaces commas with dots (in case the source CSV was generated with a comma locale).
    ' 2. Replaces any thousands separators (spaces, usually) with nothing.
    
    Dim CleanedString As String
    
    CleanedString = Replace(ValueString, ",", ".") ' Handle cases where input uses comma decimal
    CleanedString = Replace(CleanedString, " ", "") ' Remove potential thousands spaces

    ' Use CDbl() to convert the standardized string.
    On Error Resume Next
    SafeCDbl = CDbl(CleanedString)
    If Err.Number <> 0 Then
        ' If CDbl fails (e.g., non-numeric data), return 0 and log error.
        Debug.Print "Warning: Failed to convert string to number: " & ValueString & " in line: " & CleanedString
        SafeCDbl = 0
    End If
    On Error GoTo 0
End Function

' =========================================================================
' MAIN BATCH PROCESSOR
' =========================================================================

Public Sub ProcessBatch(ByVal filePath As String)
    
    Dim wsDry As Worksheet, wsWet As Worksheet
    Dim iFile As Integer, TextLine As String, Data() As String
    Dim lRowDry As Long, lRowWet As Long, lDataCount As Long
    
    ' Variable storage
    Dim I As Double, S As Double, R As Double, t As Double, t As Double
    Dim A As Double, W As Double, d As Double, h As Double
    Dim DryVolume As Double, WetVolume As Double
    
    If Dir(filePath) = "" Then
        MsgBox "Error: Input file not found: " & filePath, vbCritical
        Exit Sub
    End If
    
    lDataCount = 9
    
    ' 1. Prepare result sheets (Delete existing ones if present)
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Dry_Season_Results").Delete
    ThisWorkbook.Sheets("Wet_Season_Results").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Create and name the two result sheets
    Set wsDry = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsDry.Name = "Dry_Season_Results"
    Set wsWet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsWet.Name = "Wet_Season_Results"
    
    ' Define the common headers for all 9 inputs
    Dim headers(1 To 9) As String
    headers(1) = "Household Income (I)": headers(2) = "Household Size (S)"
    headers(3) = "Rainfall (R)": headers(4) = "Temperature (T)"
    headers(5) = "Travel Time (t)": headers(6) = "Amount Spent (A)"
    headers(7) = "Willingness To Pay (W)": headers(8) = "Shortest Distance (d)"
    headers(9) = "Height Difference (h)"
    
    ' 2. Setup Headers for BOTH Sheets
    Dim lCol As Long
    For lCol = 1 To 9
        wsDry.Cells(1, lCol).Value = headers(lCol)
        wsWet.Cells(1, lCol).Value = headers(lCol)
    Next lCol
    
    wsDry.Cells(1, 10).Value = "Calculated Dry Volume (Units)"
    wsWet.Cells(1, 10).Value = "Calculated Wet Volume (Units)"
    
    lRowDry = 2
    lRowWet = 2
    
    ' 3. Read the CSV file
    iFile = FreeFile
    On Error GoTo ErrorHandler
    Open filePath For Input As #iFile
    
    ' Skip header line
    Line Input #iFile, TextLine
    
    ' Process data lines
    Do While Not EOF(iFile)
        Line Input #iFile, TextLine
        
        If Trim(TextLine) = "" Then GoTo NextIteration
        
        Data = Split(TextLine, ",")
        
        If UBound(Data) < lDataCount - 1 Then GoTo NextIteration
        
        ' 4. Map and Convert Data using SafeCDbl()
        I = SafeCDbl(Data(0)): S = SafeCDbl(Data(1)): R = SafeCDbl(Data(2)): t = SafeCDbl(Data(3))
        t = SafeCDbl(Data(4)): A = SafeCDbl(Data(5)): W = SafeCDbl(Data(6)): d = SafeCDbl(Data(7)): h = SafeCDbl(Data(8))
        
        ' 5. Perform Calculations
        DryVolume = CalculateDrySeason(I, R, t, t, A, W, d, h)
        WetVolume = CalculateWetSeason(I, S, R, t, t, A, W, d, h)
        
        ' 6. Write Data to BOTH Sheets
        For lCol = 0 To lDataCount - 1
            ' Write inputs to the dry sheet (Data is still the string array)
            wsDry.Cells(lRowDry, lCol + 1).Value = Data(lCol)
            ' Write inputs to the wet sheet
            wsWet.Cells(lRowWet, lCol + 1).Value = Data(lCol)
        Next lCol
        
        ' Write the specific result for each sheet (column 10)
        wsDry.Cells(lRowDry, 10).Value = DryVolume
        wsWet.Cells(lRowWet, 10).Value = WetVolume
        
        lRowDry = lRowDry + 1
        lRowWet = lRowWet + 1
        
NextIteration:
    Loop
    
    ' 7. Cleanup and Formatting
    Close #iFile
    wsDry.Columns("A:J").AutoFit
    wsWet.Columns("A:J").AutoFit
    wsDry.Activate
    
    MsgBox "Batch processing complete! " & (lRowDry - 2) & " records processed. Results are on two sheets: Dry_Season_Results and Wet_Season_Results.", vbInformation
    Exit Sub

ErrorHandler:
    If iFile <> 0 Then Close #iFile
    MsgBox "A critical error occurred during processing at row " & (lRowDry - 1) & ". Check the data format in your CSV file.", vbCritical
    
End Sub

' --- Form Display Subroutine ---
Public Sub ShowBatchForm()
    Batch_Processor_Form.Show
End Sub
