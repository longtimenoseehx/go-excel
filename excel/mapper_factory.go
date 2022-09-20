// Package excel
// @Author Binary.H 2022/9/20 21:54:00
package excel

type abstractMapper interface {
	Run()
}

type importMapper struct {
	templateStruct interface{}
	data           [][]string
}

// NewImportMapper 创建一个excel导入映射器
// params: 源数据切片,模版结构体
func NewImportMapper(data [][]string, templateStruct interface{}) *importMapper {
	return &importMapper{data: data, templateStruct: templateStruct}
}

type exportMapper struct {
	sheet          string
	templateStruct interface{}
	filter         map[string]string
	dataSlice      interface{}
}

// NewExportMapper 创建一个excel导出映射器
// params: sheet名, 模版结构体, 源数据, 过滤列
func NewExportMapper(sheetName string, templateStruct interface{}, sourceData interface{}, filter map[string]string) *exportMapper {
	return &exportMapper{sheet: sheetName, templateStruct: templateStruct, filter: filter, dataSlice: sourceData}
}
