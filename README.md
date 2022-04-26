# Excel-VBA-Data-Preparer-Framework
A simple Excel VBA framework aimed at simplifying data manipulation in Excel.
  - DataGrid: a customized class object which can read and store data from different data sources, perform data manipulation and export data.
  - DataPreparer: a customized class object which can fill templates and interact with file system, fulfilling tasks that are common in workflow automation.

Quick Example: 
```
Private Sub SampleDataGrid()
  Dim pokemonData As New datagrid
  Call pokemonData.loadFromRange(rng:=Selection, rngHasHeader:=True) _
  .filterIn(colName:="Type I", arrayValues:=Array("Electric")) _
  .filterOut(colName:="Spe", arrayValues:=Array("100")) _
  .toRange(Sheet1.Range("AL1"), True)
End Sub
```

![image](https://user-images.githubusercontent.com/103709587/165399963-2d923476-dfb8-4502-9035-48688bca9efe.png)

