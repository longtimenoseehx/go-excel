# go-excel
this is a excel utils for golang language.

## 1.Get Start

### 1.1 Import excel to struct slice

For Example：

```go
// 1.create mapper and then check & read excel file stream
importMapper := excel.NewImportMapper(excelData, &model.Node{})
excelDataSlice, err := importMapper.CheckAndReadExcel(excelFile, fileHeader)
// 2.run
data, err := importMapper.Run()

```

### 1.2 Export data to excel

For Example：

```go
// 1.create mapper and run
exportMapper := excel.NewExportMapper("node", &model.Node{}, dealData, filter)
excelFile, err := exportMapper.Run()

```

