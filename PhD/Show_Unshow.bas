Attribute VB_Name = "Show_Unshow"
Private Sub ToggleCharts()

Dim ws As Worksheet
Dim co As ChartObject

' Set the worksheet containing your charts
Set ws = ThisWorkbook.Sheets("Dashboard_Dry") ' Change "Sheet1" to your actual sheet name

' Loop through each chart object on the specified worksheet
For Each co In ws.ChartObjects

    ' Toggle the visibility of the chart
    co.Visible = Not co.Visible

Next co

End Sub
