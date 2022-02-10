package xlsxexport

import (
	"bytes"
	"encoding/json"
	"fmt"
	"strings"
	"time"

	"github.com/tidwall/gjson"
	"github.com/xuri/excelize/v2"
)

type Sheet struct {
	Name   string  `json:"name" yaml:"name"`
	Fields []Field `json:"fields" yaml:"fields"`
}

type Field struct {
	Title  string `json:"title" yaml:"title"`
	Field  string `json:"field" yaml:"field"`
	Format string `json:"format" yaml:"format"`
	Func   string `json:"func" yaml:"func"`
}

type Loader func(index int) ([]interface{}, error)

func Render(ss []Sheet, loader Loader) ([]byte, error) {
	f := excelize.NewFile()

	for _, v := range ss {
		if err := renderSheet(f, v, loader); err != nil {
			return nil, err
		}
	}

	f.DeleteSheet("Sheet1")
	buf := bytes.NewBufferString("")
	if err := f.Write(buf); err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}

func renderSheet(f *excelize.File, s Sheet, loader Loader) error {
	f.NewSheet(s.Name)
	// set title
	cnm := map[int]string{}
	// 存储字段渲染
	fieldsM := map[string]RenderF{}

	for i, v := range s.Fields {
		cn, err := excelize.ColumnNumberToName(i + 1)
		if err != nil {
			return err
		}
		cnm[i] = cn
		f.SetCellStr(s.Name, fmt.Sprintf("%s1", cn), v.Title)

		fs := strings.Split(v.Format, ":")
		if format, ok := formatM[fs[0]]; ok {
			fieldsM[v.Format] = format(v)
		} else {
			return fmt.Errorf("Format %s not supported", v.Format)
		}
	}
	row := 2
	skip := 0

	for {
		datas, err := loader(skip)
		if err != nil {
			return err
		}
		if len(datas) == 0 {
			return nil
		}
		j, err := json.Marshal(datas)
		if err != nil {
			return err
		}
		objs := gjson.Parse(string(j)).Array()

		for _, obj := range objs {
			for i, field := range s.Fields {
				o := obj.Get(field.Field)
				if !o.Exists() {
					continue
				}
				fieldsM[field.Format](f, s.Name, fmt.Sprintf("%s%d", cnm[i], row), o)
			}
			row++
		}
		skip++
	}
}

func Regist(name string, r func(field Field) RenderF) {
	formatM[name] = r
}

type RenderF = func(f *excelize.File, sheet, axis string, obj gjson.Result)

var formatM = map[string]func(field Field) RenderF{
	"string": stringF,
	"time":   timeF,
	"amount": amountF,
	"enum":   enumF,
}

func stringF(field Field) func(f *excelize.File, sheet, axis string, obj gjson.Result) {
	return func(f *excelize.File, sheet, axis string, obj gjson.Result) {
		f.SetCellStr(sheet, axis, obj.String())
	}
}

// timeF time:2006-01-02 15:04:05
func timeF(field Field) func(f *excelize.File, sheet, axis string, obj gjson.Result) {
	fs := strings.Split(field.Format, ":")
	format := ""
	if len(fs) > 2 {
		format = strings.Join(fs[1:], ":")
	}
	return func(f *excelize.File, sheet, axis string, obj gjson.Result) {
		if obj.Int() != 0 {
			f.SetCellStr(sheet, axis, time.Unix(obj.Int(), 0).Format(format))
		}
	}
}

func amountF(field Field) func(f *excelize.File, sheet, axis string, obj gjson.Result) {
	return func(f *excelize.File, sheet, axis string, obj gjson.Result) {
		f.SetCellFloat(sheet, axis, obj.Float()/100, 2, 64)
	}
}

// enumF enum:A,a;B,b
func enumF(field Field) func(f *excelize.File, sheet, axis string, obj gjson.Result) {
	m := map[string]string{}
	fs := strings.Split(field.Format, ":")
	if len(fs) == 2 {
		es := strings.Split(fs[1], ";")
		for _, v := range es {
			vs := strings.Split(v, ",")
			if len(vs) == 2 {
				m[vs[0]] = vs[1]
			}
		}
	}

	return func(f *excelize.File, sheet, axis string, obj gjson.Result) {
		f.SetCellStr(sheet, axis, m[obj.String()])
	}
}
