Attribute VB_Name = "Module1"
Option Explicit

' --- Core Calculation Formulas ---
' These functions implement the dry and wet season water volume models.

Public Function CalculateDrySeason(ByVal I As Double, ByVal S As Double, ByVal R As Double, ByVal temp As Double, ByVal t As Double, ByVal A As Double, ByVal W As Double, ByVal d As Double, ByVal h As Double) As Double
    ' Dry Season Formula: V = 98.1 + 0.0003I + 5.31R + 1.080T - 2.01t - 0.0003A + 0.0804W + 0.0142d - 0.009h
    
    Dim V As Double
    V = 98.1 + (0.0003 * I) + (5.31 * R) + (1.08 * temp) - (2.01 * t) - (0.0003 * A) + (0.0804 * W) + (0.0142 * d) - (0.009 * h)
    
    CalculateDrySeason = V
End Function

' --- Main Batch Processing Subroutine ---
Public Sub ProcessBatchDry(ByVal filePath As String)
    
    Dim wsResult As Worksheet ' Worksheet to hold the calculated results
    Dim iFile As Integer      ' File handle
    Dim TextLine As String    ' Single line from CSV
    Dim Data() As String      ' Array of data points from the line
    Dim lRow As Long          ' Current output row
    Dim lDataCount As Long    ' Expected number of input variables (9)
    
    ' Variable storage
    Dim I As Double, S As Double, R As Double, temp As Double, t As Double
    Dim A As Double, W As Double, d As Double, h As Double
    Dim DryVolume As Double, WetVolume As Double
    
    If Dir(filePath) = "" Then
        MsgBox "Error: Input file not found: " & filePath, vbCritical
        Exit Sub
    End If
    
    lDataCount = 9
    
    ' 1. Prepare result sheet (Delete existing one if present)
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Results_Dry_Season").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set wsResult = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsResult.Name = "Results_Dry_Season"
    
    ' 2. Setup Headers
    wsResult.Cells(1, 1).Value = "Household Income (I)"
    wsResult.Cells(1, 2).Value = "Household Size (S)"
    wsResult.Cells(1, 3).Value = "Rainfall (R)"
    wsResult.Cells(1, 4).Value = "Temperature (T)"
    wsResult.Cells(1, 5).Value = "Travel Time (t)"
    wsResult.Cells(1, 6).Value = "Amount Spent (A)"
    wsResult.Cells(1, 7).Value = "Willingness To Pay (W)"
    wsResult.Cells(1, 8).Value = "Shortest Distance (d)"
    wsResult.Cells(1, 9).Value = "Height Difference (h)"
    wsResult.Cells(1, 10).Value = "Calculated Dry Volume (Units)"
    
    lRow = 2
    
    ' 3. Read the CSV file
    iFile = FreeFile
    On Error GoTo ErrorHandler ' Setup error handling for file operations
    Open filePath For Input As #iFile
    
    ' Skip header line
    Line Input #iFile, TextLine
    
    ' Process data lines
    Do While Not EOF(iFile)
        Line Input #iFile, TextLine
        
        If Trim(TextLine) = "" Then GoTo NextIteration
        
        Data = Split(TextLine, ",")
        
        If UBound(Data) < lDataCount - 1 Then GoTo NextIteration ' Skip if not enough data
        
        ' 4. Map and Convert Data (CDbl is critical for non-English locales)
        I = CDbl(Data(0)): S = CDbl(Data(1)): R = CDbl(Data(2)): temp = CDbl(Data(3))
        t = CDbl(Data(4)): A = CDbl(Data(5)): W = CDbl(Data(6)): d = CDbl(Data(7)): h = CDbl(Data(8))
        
        ' 5. Perform Calculations
        DryVolume = CalculateDrySeason(I, S, R, temp, t, A, W, d, h)
        
        ' 6. Write Data and Results
        Dim lCol As Long
        For lCol = 0 To lDataCount - 1
            wsResult.Cells(lRow, lCol + 1).Value = Data(lCol)
        Next lCol
        
        wsResult.Cells(lRow, 10).Value = DryVolume
                
        lRow = lRow + 1
        
NextIteration:
    Loop
    
    ' 7. Cleanup and Formatting
    Close #iFile
    wsResult.Columns("A:K").AutoFit
    wsResult.Activate
    
    MsgBox "Batch processing complete! " & (lRow - 2) & " records processed. Results are on the sheet: " & wsResult.Name, vbInformation
    Exit Sub

ErrorHandler:
    If iFile <> 0 Then Close #iFile
    MsgBox "A critical error occurred during processing at row " & (lRow - 1) & ". This usually means the data in the CSV row is malformed or contains non-numeric text.", vbCritical
    
End Sub

' --- Form Display Subroutine ---
Public Sub ShowBatchForm()
    Batch_Processor_Form_Dry.Show
End Sub
