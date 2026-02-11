Attribute VB_Name = "Module_Dry2"
Option Explicit

' --- Core Calculation Formulas ---
' These functions implement the dry and wet season water volume models.

Public Function CalculateShortestDistance(ByVal hE As Double, ByVal hN As Double, ByVal wE As Double, ByVal wN As Double) As Double
    
    Dim d As Double
    d = Sqr((hE - wE) ^ 2 + (hN - wN) ^ 2)
    
    CalculateShortestDistance = d
End Function

Public Function CalculateHeightDifference(ByVal hH As Double, ByVal wH As Double) As Double
    
    Dim h As Double
    h = Abs(hH - wH)
    
    CalculateHeightDifference = h
End Function

Public Function CalculateDrySeason(ByVal I As Double, ByVal S As Double, ByVal R As Double, ByVal temp As Double, ByVal t As Double, ByVal A As Double, ByVal W As Double, ByVal d As Double, ByVal h As Double) As Double
    ' Dry Season Formula: V = 98.1 + 0.0003I + (5.39 * S) + 5.31R + 1.080T - 2.01t - 0.0003A + 0.0804W + 0.0142d - 0.009h
    
    Dim V As Double
    V = 98.1 + (0.0003 * I) + (5.39 * S) + (0.331 * R) + (1.8 * temp) - (2.01 * t) - (0.0003 * A) + (0.0804 * W) + (0.0142 * d) - (0.009 * h)
    
    CalculateDrySeason = V
End Function

' --- Main Batch Processing Subroutine ---
Public Sub ProcessBatchDry(ByVal filePath As String)
    
    Dim wsResult As Worksheet ' Worksheet to hold the calculated results
    Dim iFile As Integer      ' File handle
    Dim TextLine As String    ' Single line from CSV
    Dim Data() As String      ' Array of data points from the line
    Dim lRow As Long          ' Current output row
    Dim lDataCount As Long    ' Expected number of input variables (13)
    
    ' Variable storage
    Dim hE As Double, hN As Double, wE As Double, wN As Double, hH As Double, wH As Double
    Dim I As Double, S As Double, R As Double, temp As Double, t As Double
    Dim A As Double, W As Double
    Dim d As Double
    Dim h As Double
    Dim DryVolume As Double
    
    If Dir(filePath) = "" Then
        MsgBox "Error: Input file not found: " & filePath, vbCritical
        Exit Sub
    End If
    
    lDataCount = 13
    
    ' 1. Prepare result sheet (Delete existing one if present)
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Results_Dry_Season").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set wsResult = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsResult.Name = "Results_Dry_Season"
    
    ' 2. Setup Headers
    wsResult.Cells(1, 1).Value = "Household Easting (m)"
    wsResult.Cells(1, 2).Value = "Household Northing (m)"
    wsResult.Cells(1, 3).Value = "Waterpoint Easting (m)"
    wsResult.Cells(1, 4).Value = "Waterpoint Northing (m)"
    wsResult.Cells(1, 5).Value = "Household Elevation (m)"
    wsResult.Cells(1, 6).Value = "Waterpoint Elevation (m)"
    wsResult.Cells(1, 7).Value = "Household Income"
    wsResult.Cells(1, 8).Value = "Household Size"
    wsResult.Cells(1, 9).Value = "Rainfall (mm/day)"
    wsResult.Cells(1, 10).Value = "Land Surface Temperature (°C)"
    wsResult.Cells(1, 11).Value = "Travel Time (mins)"
    wsResult.Cells(1, 12).Value = "Amount Spent"
    wsResult.Cells(1, 13).Value = "Willingness To Pay"
    wsResult.Cells(1, 14).Value = "Shortest Distance (m)"
    wsResult.Cells(1, 15).Value = "Height Difference (m)"
    wsResult.Cells(1, 16).Value = "Volume (L)"
    
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
        hE = CDbl(Data(0)): hN = CDbl(Data(1)): wE = CDbl(Data(2)): wN = CDbl(Data(3)): hH = CDbl(Data(4)): wH = CDbl(Data(5))
        I = CDbl(Data(6)): S = CDbl(Data(7)): R = CDbl(Data(8)): temp = CDbl(Data(9))
        t = CDbl(Data(10)): A = CDbl(Data(11)): W = CDbl(Data(12))
        
        ' 5. Perform Calculations
        d = CalculateShortestDistance(hE, hN, wE, wN)
        h = CalculateHeightDifference(hH, wH)
        DryVolume = CalculateDrySeason(I, S, R, temp, t, A, W, d, h)
        
        ' 6. Write Data and Results
        Dim lCol As Long
        For lCol = 0 To lDataCount - 1
            wsResult.Cells(lRow, lCol + 1).Value = Data(lCol)
        Next lCol
        
        wsResult.Cells(lRow, 14).Value = d
        wsResult.Cells(lRow, 15).Value = h
        wsResult.Cells(lRow, 16).Value = DryVolume
                
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
