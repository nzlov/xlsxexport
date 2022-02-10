package main

import (
	"os"

	"github.com/nzlov/xlsxexport"
)

type People struct {
	Name string `json:"name"`
	Age  int    `json:"age"`
}

func main() {
	s := []xlsxexport.Sheet{
		{
			Name: "人员",
			Fields: []xlsxexport.Field{
				{
					Title:  "姓名",
					Field:  "obj.name",
					Format: "string",
				},
				{
					Title:  "年龄",
					Field:  "obj.age",
					Format: "string",
				},
			},
		},
	}

	datas := []interface{}{
		map[string]interface{}{
			"obj": People{Name: "a", Age: 10},
		},
		map[string]interface{}{
			"obj": People{Name: "b", Age: 32},
		},
		map[string]interface{}{
			"obj": People{Name: "c", Age: 20},
		},
	}

	data, err := xlsxexport.Render(s, func(index int) ([]interface{}, error) {
		if index != 0 {
			return []interface{}{}, nil
		}
		return datas, nil
	})
	if err != nil {
		panic(err)
	}

	os.WriteFile("a.xlsx", data, os.ModePerm)
}
