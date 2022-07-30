// Package excel
// @Author Administrator 2022/7/30 14:48:00
package excel

import (
	"fmt"
	"go-excel/common"
	"mime/multipart"
	"reflect"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/tealeg/xlsx"
)

// ImportExcel2StructSlice 导入
// params: file xlsx包下的file结构体指针, templateStruct 要解析到的结构体(需要有excel标签)
// returns:  []templateStruct, err
func ImportExcel2StructSlice(file *multipart.File, templateStruct interface{}) (*[]interface{}, error) {
	xlsx, err := excelize.OpenReader(*file)
	if err != nil {
		return nil, common.ExcelOptErr
	}
	rows := xlsx.GetRows(xlsx.GetSheetMap()[1])

	_, colFieldMap, serveError := getStructTagFieldMap(templateStruct, nil)
	if serveError != nil {
		return nil, serveError
	}

	var result []interface{}
	var fieldSlice []string
	for idx, row := range rows {
		if idx == 0 {
			for _, colName := range row {
				if fieldName, ok := colFieldMap[colName]; ok {
					fieldSlice = append(fieldSlice, fieldName)
				} else {
					return nil, common.ColNotMatchFieldErr
				}
			}
		} else {
			newStruct, parseErr := setRow2StructField(templateStruct, idx+1, row, fieldSlice)
			if parseErr != nil {
				return nil, parseErr
			}
			result = append(result, newStruct)
		}
	}
	return &result, nil
}

// ExportStructSlice2Excel 导出
// params: sheetName导出的sheet名, dataSlice数据源(支持结构体/结构体切片/结构体指针/结构体指针切片),
//		   templateStruct数据源对应的结构体(支持结构体/结构体指针), filter导出时需要过滤掉的字段(key:结构体中excel标签的值(列名),value:不限)
// returns:  file文件指针, err
func ExportStructSlice2Excel(sheetName string, dataSlice interface{}, templateStruct interface{}, filter map[string]string) (*xlsx.File, error) {
	if dataSlice == nil {
		return nil, common.NilParamErr
	}
	file := xlsx.NewFile()
	sheet, err := file.AddSheet(sheetName)
	if err != nil {
		return file, common.ExcelOptErr
	}

	colSlice, colFieldMap, parseStructErr := getStructTagFieldMap(templateStruct, filter)
	if parseStructErr != nil {
		return nil, parseStructErr
	}

	var fieldSlice []string

	titleStyle := getTitleStyle()
	row := sheet.AddRow()
	var cell *xlsx.Cell
	for _, colName := range colSlice {
		cell = row.AddCell()
		cell.Value = colName
		cell.SetStyle(titleStyle)
		if fieldName, ok := colFieldMap[colName]; ok {
			fieldSlice = append(fieldSlice, fieldName)
		}
	}

	getValue := reflect.ValueOf(dataSlice)
	if getValue.Kind() == reflect.Ptr {
		getValue = getValue.Elem()
	}
	if getValue.Kind() == reflect.Slice || getValue.Kind() == reflect.Array {
		for i := 0; i < getValue.Len(); i++ {
			value := getValue.Index(i)
			setErr := setStruct2ExcelRow(sheet, value, fieldSlice)
			if setErr != nil {
				return nil, setErr
			}
		}
	} else if getValue.Kind() == reflect.Struct {
		setErr := setStruct2ExcelRow(sheet, getValue, fieldSlice)
		if setErr != nil {
			return nil, setErr
		}
	} else {
		return nil, common.SourceTypeNotSupportErr
	}

	return file, nil
}

// getStructTagFieldMap 解析结构体中的自定义excel标签
// params: templateStruct目标结构体(支持结构体/结构体指针), filter解析时需要过滤掉的字段(key:结构体中excel标签的值(列名),value:不限)
// returns: 结构体字段excel标签值(excel列名)数组, excel标签->字段名的映射, err
func getStructTagFieldMap(templateStruct interface{}, filter map[string]string) ([]string, map[string]string, error) {
	if templateStruct == nil {
		return nil, nil, common.NilParamErr
	}
	getType := reflect.TypeOf(templateStruct)
	if getType.Kind() == reflect.Ptr {
		getType = getType.Elem()
	}
	if getType.Kind() != reflect.Struct {
		return nil, nil, common.TargetTypeNotSupportErr
	}

	columnsSlice := make([]string, 0)
	colFieldMap := make(map[string]string, getType.NumField())
	for i := 0; i < getType.NumField(); i++ {
		colName, needExport := getType.Field(i).Tag.Lookup(common.ExcelTag)
		if _, exist := filter[colName]; exist || !needExport {
			continue
		}
		columnsSlice = append(columnsSlice, colName)
		colFieldMap[colName] = getType.Field(i).Name
	}
	if len(columnsSlice) == 0 || len(colFieldMap) == 0 {
		return nil, nil, common.TagNotFoundErr
	}
	return columnsSlice, colFieldMap, nil
}

// getTitleStyle 生成excel标题样式
func getTitleStyle() *xlsx.Style {
	centerHAlign := *xlsx.DefaultAlignment()
	centerHAlign.Horizontal = "center"
	centerHAlign.Vertical = "center"
	titleStyle := xlsx.NewStyle()
	font := *xlsx.NewFont(14, "等线")
	font.Bold = true
	titleStyle.Font = font
	titleStyle.Alignment = centerHAlign
	titleStyle.ApplyAlignment = true
	titleStyle.ApplyBorder = true

	return titleStyle
}

// setRow2StructField 导入时使用的工具函数： 解析excel表格中的一行到指定结构体templateStruct
// params: templateStruct目标结构体, rowId行号, row行数据, fieldSlice列名对应的结构体字段名的切片(注意字段名顺序要和列名保持一致)
// returns: 目标结构体, err
func setRow2StructField(templateStruct interface{}, rowId int, row []string, fieldSlice []string) (interface{}, error) {
	if templateStruct == nil || len(row) == 0 || len(fieldSlice) == 0 {
		return nil, common.NilParamErr
	}
	t := reflect.TypeOf(templateStruct)
	if t.Kind() == reflect.Ptr { //指针类型获取真正type需要调用Elem
		t = t.Elem()
	}
	if t.Kind() != reflect.Struct {
		return nil, common.TargetTypeNotSupportErr
	}

	newStruct := reflect.New(t)
	if newStruct.Kind() == reflect.Ptr {
		newStruct = newStruct.Elem()
	}
	if newStruct.Kind() != reflect.Struct {
		return nil, common.TargetTypeNotSupportErr
	}
	for i, cell := range row {
		if len(cell) == 0 {
			continue
		}
		fieldName := fieldSlice[i]
		fieldVal := newStruct.FieldByName(fieldName)
		structField, hasField := t.FieldByName(fieldName)
		if !hasField {
			return nil, common.ColNotMatchFieldErr
		}
		switch structField.Type.Kind() {
		case reflect.String:
			fieldVal.SetString(cell)
		case reflect.Int:
			parseInt, err := strconv.Atoi(cell)
			if err != nil {
				return nil, common.DataParseErrOn(rowId, i, reflect.Int.String())
			}
			fieldVal.SetInt(int64(parseInt))
		case reflect.Int32:
			parseInt, err := strconv.ParseInt(cell, 10, 32)
			if err != nil {
				return nil, common.DataParseErrOn(rowId, i, reflect.Int32.String())
			}
			fieldVal.SetInt(parseInt)
		case reflect.Int64:
			parseInt, err := strconv.ParseInt(cell, 10, 64)
			if err != nil {
				return nil, common.DataParseErrOn(rowId, i, reflect.Int64.String())
			}
			fieldVal.SetInt(parseInt)
		case reflect.Float32:
			float, err := strconv.ParseFloat(cell, 32)
			if err != nil {
				return nil, common.DataParseErrOn(rowId, i, reflect.Float32.String())
			}
			fieldVal.SetFloat(float)
		case reflect.Float64:
			float, err := strconv.ParseFloat(cell, 64)
			if err != nil {
				return nil, common.DataParseErrOn(rowId, i, reflect.Float64.String())
			}
			fieldVal.SetFloat(float)
		case reflect.Bool:
			parseBool, err := strconv.ParseBool(cell)
			if err != nil {
				return nil, common.DataParseErrOn(rowId, i, reflect.Bool.String())
			}
			fieldVal.SetBool(parseBool)
		default:
			return nil, common.TypeNotSupport(structField.Type.Kind().String())
		}
	}
	return newStruct.Interface(), nil
}

// setStruct2ExcelRow 导出时使用的工具函数： 将单个结构体的数据解析到excel的一行上
// params: sheet, value目标结构体反射后得到的Value类型, fieldSlice字段名的切片
func setStruct2ExcelRow(sheet *xlsx.Sheet, value reflect.Value, fieldSlice []string) error {
	if value.Kind() == reflect.Ptr {
		value = value.Elem()
	}
	if value.Kind() != reflect.Struct {
		return common.TargetTypeNotSupportErr
	}
	row := sheet.AddRow()
	var cell *xlsx.Cell
	for _, fieldName := range fieldSlice {
		fieldValue := value.FieldByName(fieldName)
		cell = row.AddCell()
		switch fieldValue.Kind() {
		case reflect.Int:
			fallthrough
		case reflect.Int32:
			fallthrough
		case reflect.Int64:
			cell.Value = strconv.Itoa(int(fieldValue.Int()))
		case reflect.Float32:
			cell.Value = strconv.FormatFloat(fieldValue.Float(), 'E', -1, 32)
		case reflect.Float64:
			cell.Value = strconv.FormatFloat(fieldValue.Float(), 'E', -1, 64)
		case reflect.Bool:
			cell.Value = fmt.Sprintf("%t", fieldValue.Bool())
		case reflect.String:
			fallthrough
		default:
			cell.Value = fieldValue.String()
		}
	}
	return nil
}
