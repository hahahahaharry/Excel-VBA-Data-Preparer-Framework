# Excel-VBA-Data-Preparer-Framework
A simple Excel VBA framework aimed at simplifying data manipulation in Excel.
  - DataGrid(WIP): a customized class object which can read and store data from different data sources, perform data manipulation and export data.
  - DataPreparer(WIP): a customized class object which can perform tasks that are common in Excel-related workflow automation projects.

Quick Example: 
Private Sub SampleDataGrid()
    Dim pokemonData As New datagrid
    Call pokemonData.loadFromRange(rng:=Selection, rngHasHeader:=True) _
    .filterIn(colName:="Type I", arrayValues:=Array("Electric")) _
    .filterOut(colName:="Spe", arrayValues:=Array("100")) _
    .toRange(Sheet1.Range("AL1"), True)
End Sub

![image](https://user-images.githubusercontent.com/103709587/165398035-f6ec8540-9ebf-4493-a52f-b16d6062cf4a.png)
