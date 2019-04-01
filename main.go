package main

import (
	"flag"
	"fmt"
	"io/ioutil"
	"os"
	"strings"
	"sync"

	"github.com/360EntSecGroup-Skylar/excelize"
)

// func usage() {
// 	fmt.Fprintf(os.Stderr, "usage: excel2lua [inputfile]\n")
// 	flag.PrintDefaults()
// 	os.Exit(2)
// }

func assert(cond bool, errmsg string) {
	if !cond {
		panic(errmsg)
	}
}

type field struct {
	_index int
	_name  string
	_type  string // bool,int,string,array,dict,comment
	_attr  string // nil,unique,client,server
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
		return ""
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
		if cell == "1" {
			value = "true"
		} else {
			value = "false"
		}
		// value = strings.ToLower(cell)
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

func processSheet(xlsx *excelize.File, fileName string, sheetName string, wg *sync.WaitGroup, isClient bool, outputPath string) {
	defer wg.Done()
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
	// luaCode := "return {\n"

	sheetNameUpper := strings.ToUpper(sheetName)
	luaCode := fmt.Sprintf(`--[[
* @file        : %s.lua
* @author      : Steve_Lhf
* @sour        : excel/%s
* @sheet name  : %s
* brief:       : this file was create by tools, DO NOT modify it!
* Copyright(C) 2017 ONEMT, All rights reserved
--]]`, sheetName, fileName, sheetName)
	luaCode += fmt.Sprintf("\n\nlocal %s =\n{\n", sheetNameUpper)

	fields := parseFields(rows[0], rows[1])
	fmt.Printf("fields:%v\n", fields)

	for rowIndex, row := range rows {
		row = row[:len(row)-1]
		fmt.Println("row:", rowIndex, row, len(row))
		if rowIndex < 3 {
			continue
		}
		// line := fmt.Sprintf("   {", rowIndex-2)
		line := "  {"
		for colIndex, colCell := range row {
			fmt.Println("col:", colIndex, colCell)
			f := fields[colIndex]
			assert(f != nil, "field not found")
			if f._type == "comment" {
				continue // skip comment
			}
			if (isClient && f._attr == "server") || (!isClient && f._attr == "client") {
				continue
			}
			line += cellWrapper(f, colCell)
		}
		line += "},\n"
		luaCode += line
	}
	// luaCode += "}\n"
	luaCode += fmt.Sprintf("};\nreturn %s\n", strings.ToUpper(sheetName))
	ioutil.WriteFile("output/"+sheetName+".lua", []byte(luaCode), 0644)
}

func main() {
	// flag.Usage = usage
	isClient := flag.Bool("client", true, "true=>client, false=>server")
	outputPath := flag.String("output", "output/", "The path you want to output.")
	intputFile := flag.String("input", "input.xlsx", "The xlsx file you want to convert.")
	flag.Parse()

	if len(*intputFile) == 0 {
		fmt.Println("Input file is missing.")
		os.Exit(1)
	}

	xlsx, err := excelize.OpenFile(*intputFile)
	if err != nil {
		fmt.Println(err)
		return
	}
	// Get all sheets
	var wg sync.WaitGroup
	for _, sheetName := range xlsx.GetSheetMap() {
		wg.Add(1)
		go processSheet(xlsx, *intputFile, sheetName, &wg, *isClient, *outputPath)
	}
	wg.Wait()
}
