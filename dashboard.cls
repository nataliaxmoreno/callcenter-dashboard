VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
With Application
  .Calculation = xlCalculationManual
  .ScreenUpdating = False
  .DisplayStatusBar = False
  .EnableEvents = False
  .DisplayAlerts = False
End With

Dim pt As PivotTable

For Each pt In Sheet3.PivotTables
    pt.RefreshTable
Next
With Application
  .Calculation = xlCalculationAutomatic
  .ScreenUpdating = True
  .DisplayStatusBar = True
  .EnableEvents = True
  .DisplayAlerts = True
End With

End Sub



Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
Dim mychart As Chart

    ActiveSheet.ChartObjects("support").Activate
    
    ActiveChart.FullSeriesCollection(1).Points(3).DataLabel.Select
    Selection.Left = 116.012
    Selection.Top = 125.012
    ActiveSheet.ChartObjects("billing").Activate
    
    ActiveChart.FullSeriesCollection(1).Points(1).DataLabel.Select
    Selection.Left = 116.012
    Selection.Top = 125.012
    ActiveSheet.ChartObjects("sales").Activate
    
    ActiveChart.FullSeriesCollection(1).Points(2).DataLabel.Select
    Selection.Left = 116.012
    Selection.Top = 125.012
    
    ActiveSheet.ChartObjects("abandoned").Activate
     ActiveChart.FullSeriesCollection(1).Points(2).DataLabel.Select
    Selection.Left = 135.131
    Selection.Top = 95.344
    
    
    Range("h50").Select
    
    




End Sub
