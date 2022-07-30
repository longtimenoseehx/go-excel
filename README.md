# go-excel
this is a excel utils for golang language.

## 1.Get Start

### 1.1 Import excel to struct slice

For Example：

```go
// 1.call function of ImportExcel2StructSlice()
// file: ptr of struct in pkg:xlsx
// templateStruct: elem type of result:[]templateStruct (support: struct or ptr of struct)
interfaceSlice, err := excel.ImportExcel2StructSlice(file *multipart.File, templateStruct interface{})
if err != nil {
    // ...
}
// 2.parse result to your model slice
var modelSlice []Model
for (interface{} item : interfaceSlice) {
    modelSlice = append(modelSlice, item.(Model))
}
```

### 1.2 Export data to excel

For Example：

```go
// 1.call function of ExportStructSlice2Excel()
// sheetName: the sheetname in excel file to be create
// dataSlice: the source data you want to parse (support: struct, ptr of struct, struct slice, struct array...all like this)
// templateStruct: elem type of previous param (support: struct or ptr of struct)
// filter: the map'key is excel tag'value, and map'value is useless. used to filter the field that you are not want to export
excelFile, err := excel.ExportStructSlice2Excel(sheetName string, dataSlice interface{}, templateStruct interface{}, filter map[string]string)
if err != nil {
    // ...
}
```

