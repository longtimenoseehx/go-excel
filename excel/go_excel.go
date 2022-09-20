// Package excel
// @Author Binary.H 2022/7/30 14:48:00
package excel

import (
	"encoding/csv"
	"fmt"
	"github.com/longtimenoseehx/go-excel/common"
	"github.com/xuri/excelize/v2"
	"mime/multipart"
	"reflect"
	"strconv"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

// Run 导入
// returns:  *[]templateStruct, err
func (mapper *importMapper) Run() (*[]interface{}, error) {
	tagValSlice, colFieldMap, serveError := parseStructExcelFieldInfo(mapper.templateStruct, nil)
	if serveError != nil {
		return nil, serveError
	}

	var result []interface{}
	IdxFieldMap := make(map[int]string, len(tagValSlice))
	for rowIdx, row := range mapper.data {
		if rowIdx == 0 {
			for colIdx, col := range row {
				if len(col) == 0 {
					continue
				}
				// 遍历第一行，构建[][]string列索引 -> 结构体字段名的映射
				if fieldName, ok := colFieldMap[col]; ok {
					IdxFieldMap[colIdx] = fieldName
					delete(colFieldMap, col)
				}
			}
		} else {
			// 检查必填栏，缺失则报错
			if len(colFieldMap) > 0 {
				for _, val := range tagValSlice {
					if _, ok := colFieldMap[val[1:]]; ok && strings.HasPrefix(val, "*") {
						return nil, common.LackRequiredColOf(val[1:])
					}
				}
			}
			newStruct, parseErr := setRow2StructField(mapper.templateStruct, rowIdx, row, IdxFieldMap)
			if parseErr != nil {
				return nil, parseErr
			}
			result = append(result, newStruct)
		}
	}
	return &result, nil
}

// Run 导出
// returns:  file文件指针, err
func (mapper *exportMapper) Run() (*xlsx.File, error) {
	if mapper.dataSlice == nil {
		return nil, common.NilParamErr
	}
	file := xlsx.NewFile()
	sheet, err := file.AddSheet(mapper.sheet)
	if err != nil {
		return file, common.ExcelOptErr
	}

	tagValSlice, colFieldMap, parseStructErr := parseStructExcelFieldInfo(mapper.templateStruct, mapper.filter)
	if parseStructErr != nil {
		return nil, parseStructErr
	}

	var fieldSlice []string

	requireStyle := getTitleStyle(true)
	commonStyle := getTitleStyle(false)
	row := sheet.AddRow()
	var cell *xlsx.Cell
	for _, tagVal := range tagValSlice {
		cell = row.AddCell()
		if strings.HasPrefix(tagVal, "*") {
			tagVal = tagVal[1:]
			cell.SetStyle(requireStyle)
		} else {
			cell.SetStyle(commonStyle)
		}
		cell.Value = tagVal
		if fieldName, ok := colFieldMap[tagVal]; ok {
			fieldSlice = append(fieldSlice, fieldName)
		}
	}

	getValue := reflect.ValueOf(mapper.dataSlice)
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

// parseStructExcelFieldInfo 解析结构体中的自定义excel标签
// params: templateStruct目标结构体(支持结构体/结构体指针), filter解析时需要过滤掉的字段(key:结构体中excel标签的值(列名),value:不限)
// returns: 结构体字段的excel标签值的数组, excel列名->字段名的映射, err
func parseStructExcelFieldInfo(templateStruct interface{}, filter map[string]string) ([]string, map[string]string, error) {
	if templateStruct == nil {
		return nil, nil, common.NilParamErr
	}
	getType := reflect.TypeOf(templateStruct)
	if getType.Kind() == reflect.Ptr {
		getType = getType.Elem()
	}
	if getType.Kind() != reflect.Struct {
		return nil, nil, common.TemplateTypeNotSupportErr
	}

	tagValSlice := make([]string, 0)
	colFieldMap := make(map[string]string, getType.NumField())
	for i := 0; i < getType.NumField(); i++ {
		tagVal, needExport := getType.Field(i).Tag.Lookup(common.ExcelTag)
		if _, exist := filter[tagVal]; exist || !needExport {
			continue
		}
		tagValSlice = append(tagValSlice, tagVal)
		if strings.HasPrefix(tagVal, "*") {
			tagVal = tagVal[1:]
		}
		colFieldMap[tagVal] = getType.Field(i).Name
	}
	if len(colFieldMap) == 0 || len(tagValSlice) == 0 {
		return nil, nil, common.TagNotFoundErr
	}
	return tagValSlice, colFieldMap, nil
}

// getTitleStyle 生成excel标题样式
func getTitleStyle(require bool) *xlsx.Style {
	centerHAlign := *xlsx.DefaultAlignment()
	centerHAlign.Horizontal = "center"
	centerHAlign.Vertical = "center"
	titleStyle := xlsx.NewStyle()
	font := *xlsx.NewFont(14, "等线")
	titleStyle.Font = font
	if require {
		titleStyle.Font.Color = xlsx.RGB_Dark_Red
	}
	titleStyle.Alignment = centerHAlign
	titleStyle.ApplyAlignment = true
	titleStyle.ApplyBorder = true

	return titleStyle
}

// setRow2StructField 导入时使用的工具函数： 解析excel表格中的一行到指定结构体templateStruct
// params: templateStruct目标结构体, rowIdx行号 row行数据, fieldSlice列名对应的结构体字段名的切片(注意字段名顺序要和列名保持一致)
// returns: 目标结构体, err
func setRow2StructField(templateStruct interface{}, rowIdx int, row []string, IdxFieldMap map[int]string) (interface{}, error) {
	if templateStruct == nil || len(row) == 0 || len(IdxFieldMap) == 0 {
		return nil, common.NilParamErr
	}
	t := reflect.TypeOf(templateStruct)
	if t.Kind() == reflect.Ptr { //指针类型获取真正type需要调用Elem
		t = t.Elem()
	}
	if t.Kind() != reflect.Struct {
		return nil, common.TemplateTypeNotSupportErr
	}

	newStruct := reflect.New(t)
	if newStruct.Kind() == reflect.Ptr {
		newStruct = newStruct.Elem()
	}
	if newStruct.Kind() != reflect.Struct {
		return nil, common.TemplateTypeNotSupportErr
	}
	for colIdx, fieldName := range IdxFieldMap {
		fieldVal := newStruct.FieldByName(fieldName)
		structField, hasField := t.FieldByName(fieldName)
		if !hasField {
			return nil, common.ColNotMatchFieldErr
		}
		if !fieldVal.CanSet() {
			return nil, common.NilParamErr
		}

		var cell string
		if rowIdx < len(row) {
			cell = row[colIdx]
		}
		// 如果单元格值为空，查看是否为必填项，是则报错。
		if len(cell) == 0 {
			if tagVal, ok := structField.Tag.Lookup(common.ExcelTag); ok && strings.HasPrefix(tagVal, "*") {
				return nil, common.LackRequiredParamOn(rowIdx, colIdx)
			}
			continue
		}
		// 开始解析数据
		switch structField.Type.Kind() {
		case reflect.String:
			fieldVal.SetString(cell)
		case reflect.Int:
			parseInt, err := strconv.Atoi(cell)
			if err != nil {
				return nil, common.ParamTypeErrOn(rowIdx, colIdx, common.Int)
			}
			fieldVal.SetInt(int64(parseInt))
		case reflect.Int32:
			parseInt, err := strconv.ParseInt(cell, 10, 32)
			if err != nil {
				return nil, common.ParamTypeErrOn(rowIdx, colIdx, common.Int)
			}
			fieldVal.SetInt(parseInt)
		case reflect.Int64:
			parseInt, err := strconv.ParseInt(cell, 10, 64)
			if err != nil {
				return nil, common.ParamTypeErrOn(rowIdx, colIdx, common.Int)
			}
			fieldVal.SetInt(parseInt)
		case reflect.Float32:
			float, err := strconv.ParseFloat(cell, 32)
			if err != nil {
				return nil, common.ParamTypeErrOn(rowIdx, colIdx, common.Float)
			}
			fieldVal.SetFloat(float)
		case reflect.Float64:
			float, err := strconv.ParseFloat(cell, 64)
			if err != nil {
				return nil, common.ParamTypeErrOn(rowIdx, colIdx, common.Float)
			}
			fieldVal.SetFloat(float)
		case reflect.Bool:
			parseBool, err := strconv.ParseBool(cell)
			if err != nil {
				return nil, common.ParamTypeErrOn(rowIdx, colIdx, common.Bool)
			}
			fieldVal.SetBool(parseBool)
		case reflect.Ptr, reflect.Struct:
			if fieldVal.CanInterface() {
				switch fieldVal.Interface().(type) {
				case *time.Time, time.Time:
					parse, parseErr := time.Parse("2006-01-02 15:04:05", cell)
					if parseErr != nil {
						return nil, common.ParamTypeErrOn(rowIdx, colIdx, common.Time)
					}
					if fieldVal.Kind() == reflect.Ptr {
						fieldVal.Set(reflect.ValueOf(&parse))
					} else {
						fieldVal.Set(reflect.ValueOf(parse))
					}
				default:
					return nil, common.TypeNotSupport(fieldName)
				}
			}
		default:
			return nil, common.TypeNotSupport(fieldName)
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
		return common.TemplateTypeNotSupportErr
	}
	row := sheet.AddRow()
	var cell *xlsx.Cell
	for _, fieldName := range fieldSlice {
		fieldValue := value.FieldByName(fieldName)
		cell = row.AddCell()
		if !fieldValue.CanAddr() {
			continue
		}
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
		case reflect.Ptr:
			if fieldValue.IsNil() {
				continue
			}
			fieldValue.Elem()
			fallthrough
		case reflect.Struct:
			if !fieldValue.IsZero() && fieldValue.CanInterface() {
				interfaceObj := fieldValue.Interface()
				switch interfaceObj.(type) {
				case *time.Time:
					ptr := interfaceObj.(*time.Time)
					cell.Value = ptr.Format("2006-01-02 15:04:05")
				case time.Time:
					obj := interfaceObj.(time.Time)
					cell.Value = obj.Format("2006-01-02 15:04:05")
				default:
					cell.Value = fieldValue.String()
				}
			}
		case reflect.String:
			fallthrough
		default:
			cell.Value = fieldValue.String()
		}
	}
	return nil
}

// CheckAndReadExcel 检查导入的excel文件格式, 并返回读取出的数据切片
func (mapper *importMapper) CheckAndReadExcel(file *multipart.File, fileHeader *multipart.FileHeader) ([][]string, error) {
	var recordSlice [][]string
	var readErr error
	fileNameArr := strings.Split(fileHeader.Filename, ".")
	switch fileNameArr[len(fileNameArr)-1] {
	case common.Xls:
		fallthrough
	case common.Xlsx:
		recordSlice, readErr = readXlsx(file)
	case common.Csv:
		recordSlice, readErr = readCsv(file)
	default:
		return nil, common.ExcelFormatErr
	}
	if readErr != nil {
		return nil, common.ExcelOptErr
	}
	return recordSlice, nil
}

func readXlsx(file *multipart.File) ([][]string, error) {
	xlsx, err := excelize.OpenReader(*file)
	if err != nil {
		return nil, err
	}
	rows, getRowErr := xlsx.GetRows(xlsx.GetSheetMap()[1])
	if getRowErr != nil {
		return nil, getRowErr
	}
	return rows, nil
}

func readCsv(file *multipart.File) ([][]string, error) {
	reader := csv.NewReader(*file)
	reader.FieldsPerRecord = -1
	reader.LazyQuotes = true
	records, readErr := reader.ReadAll()
	if readErr != nil {
		return nil, readErr
	}
	return records, nil
}
