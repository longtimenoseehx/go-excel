// Package common
// @Author Administrator 2022/7/30 14:51:00
package common

import (
	"fmt"
	"strconv"
)

const (
	ExcelOptErrCOde   = -1
	NoneMatchTypeCode = -2
	ParamErrCode      = -3
)

const (
	Int   = "整数"
	Float = "浮点数"
	Bool  = "布尔值"
)

type ExcelError struct {
	ErrCode int
	ErrMsg  string
}

func NewExcelErr(errCode int, errMsg string) error {
	err := ExcelError{
		ErrCode: errCode,
		ErrMsg:  errMsg,
	}

	return err
}

func TypeNotSupport(errMsg string) error {
	err := ExcelError{
		ErrCode: NoneMatchTypeCode,
		ErrMsg:  fmt.Sprintf("暂不支持此类型字段：%s", errMsg),
	}

	return err
}

func ParamTypeErrOn(row, col int, kind string) error {
	err := ExcelError{
		ErrCode: NoneMatchTypeCode,
		ErrMsg:  fmt.Sprintf("[%d,%d]处的参数必须为%s！请检查", row, col, kind),
	}
	return err
}

func LackRequiredColOf(colName string) error {
	err := ExcelError{
		ErrCode: NoneMatchTypeCode,
		ErrMsg:  fmt.Sprintf("缺失必填列【%s】, 请核对！", colName),
	}
	return err
}

func LackRequiredParamOn(row, col int) error {
	err := ExcelError{
		ErrCode: NoneMatchTypeCode,
		ErrMsg:  fmt.Sprintf("[%d,%d]处为必填参数！请检查", row, col),
	}
	return err
}

func (e ExcelError) Error() string {
	if e.ErrMsg == "" {
		return strconv.Itoa(e.ErrCode)
	}
	return e.ErrMsg
}

func (e ExcelError) Code() int {
	return e.ErrCode
}

var ExcelOptErr = NewExcelErr(ExcelOptErrCOde, "Excel文件流操作异常")
var NilParamErr = NewExcelErr(ParamErrCode, "参数不能为空")
var TagNotFoundErr = NewExcelErr(ParamErrCode, "未识别到具有Excel标签的字段")
var TemplateTypeNotSupportErr = NewExcelErr(ParamErrCode, "非法的模板结构体类型：必须为结构体指针/结构体")
var SourceTypeNotSupportErr = NewExcelErr(ParamErrCode, "非法的数据源类型：必须为结构体/结构体切片/结构体数组")
var ColNotMatchFieldErr = NewExcelErr(ParamErrCode, "Excel列名与结构体字段Tag值不匹配")
var ExcelFormatErr = NewExcelErr(ParamErrCode, "文件格式错误！仅支持xls xlsx csv格式")
