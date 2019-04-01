package main

import (
	"flag"
	"fmt"
	"os"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func usage() {
	fmt.Fprintf(os.Stderr, "usage: excel2lua [inputfile]\n")
	flag.PrintDefaults()
	os.Exit(2)
}

func assert(cond bool, errmsg string) {
	if !cond {
		panic(errmsg)
	}
}

type field struct {
	_index int
	_name  string
	_type  string // bool,int,string,array,dict,comment
	_attr  string
}

func parseFields(fieldsName []string, fieldsDesc []string) []*field {
	fmt.Println("parseFields", fieldsName, fieldsDesc)
	size := len(fieldsName)
	assert(size == len(fieldsDesc), "size: fieldsName != fieldsDesc")
	fieldList := make([]*field, size, size)
	for i := 0; i < size; i++ {
		desc := strings.Split(fieldsDesc[i], "|")
		f := &field{
			_index: i,
			_name:  strings.TrimSpace(fieldsName[i]),
			_type:  strings.TrimSpace(desc[0]),
		}
		if len(desc) >= 2 {
			f._attr = strings.TrimSpace(desc[1])
		} else {
			f._attr = "nil"
		}
		fieldList[i] = f
	}
	return fieldList
}

func getDefaultValue(_type string) string {
	switch _type {
	case "bool":
		return "false"
	case "int":
		return "0"
	case "string":
		return "\"\""
	case "array":
		return ""
	case "dict":
		return ""
	}
	fmt.Printf("invalid type[%v]", _type)
	panic("don't have default value in this type:" + _type)
}

func cellWrapper(f *field, cell string) string {
	fmt.Println("cellWrapper", f, cell)
	if cell == "" {
		cell = getDefaultValue(f._type)
	}
	var value string
	switch f._type {
	case "bool":
		value = strings.ToLower(cell)
	case "int":
		value = cell
	case "string":
		value = fmt.Sprintf("\"%s\"", cell)
	case "array":
		value = fmt.Sprintf("{%s}", cell)
	case "dict":
		// input: "1",10007,5|"2",10007,5
		// output: {["1"]={10007,5},["2"]={10007,5}}
		if cell == "" {
			value = "{}"
		} else {
			value = "{"
			for _, item := range strings.Split(cell, "|") {
				// item: "1",10007,5
				itemCells := strings.Split(item, ",")
				fmt.Println("item, itemCells", item, len(item), itemCells, len(itemCells), cell)
				size := len(itemCells)
				if size < 2 {
					panic(fmt.Sprintf("Invalid dict format: %s", cell))
				} else if size == 2 {
					value += fmt.Sprintf("[%v]=%v,", itemCells[0], itemCells[1])
				} else {
					value += fmt.Sprintf("[%v]={%v},", itemCells[0], strings.Join(itemCells[1:], ","))
				}
			}
			value += "}"
		}
	case "comment":
	}
	return fmt.Sprintf("--[[%s]]%v,", f._name, value)
}

func processSheet(xlsx *excelize.File, sheetName string) {
	fmt.Println("Process sheet", sheetName)
	rows, err := xlsx.GetRows(sheetName)
	if err != nil {
		fmt.Println(err)
		return
	}
	if len(rows) <= 2 {
		fmt.Printf("invalid header")
		return
	}
	luaCode := "return {\n"
	fields := parseFields(rows[0], rows[1])
	fmt.Printf("fields:%v\n", fields)

	for rowIndex, row := range rows {
		row = row[:len(row)-1]
		fmt.Println("row:", rowIndex, row, len(row))
		if rowIndex < 3 {
			continue
		}
		// line := fmt.Sprintf("  --[[%d]]{", rowIndex-2)
		line := "  {"
		for colIndex, colCell := range row {
			fmt.Println("col:", colIndex, colCell)
			f := fields[colIndex]
			assert(f != nil, "field not found")
			if f._type == "comment" {
				continue // skip comment
			}
			line += cellWrapper(f, colCell)
		}
		line += "},\n"
		luaCode += line
	}
	luaCode += "}\n\n"
	fmt.Println(sheetName, luaCode)
}

func main() {
	flag.Usage = usage
	flag.Parse()

	args := flag.Args()
	if len(args) < 1 {
		fmt.Println("Input file is missing.")
		os.Exit(1)
	}

	xlsx, err := excelize.OpenFile(args[0])
	if err != nil {
		fmt.Println(err)
		return
	}
	// Get all sheets
	for _, name := range xlsx.GetSheetMap() {
		processSheet(xlsx, name)
	}
}
