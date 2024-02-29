## Load from site


 Workbooks(1). -- Firts book   
 Worksheets(4) -- four list   

```vba
Sub LoadDataFromWeb()

Set shFirstQtr = Workbooks(1).Worksheets(4)
Set qtQtrResults = shFirstQtr.QueryTables.Add(Connection:="URL;https://tech.com/report", Destination:=shFirstQtr.Cells(1, 1))

With qtQtrResults
 .WebFormatting = xlYes
 .WebSelectionType = xlSpecifiedTables
 .WebTables = "1,2,5,7,8,9,10"
 .Refresh
End With

End Sub
```
