package main

import (
	"fmt"
	"strconv"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	f, err := excelize.OpenFile("test1.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}

	sheet := "Table"

	err = f.RemoveCol(sheet, "A")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.RemoveCol(sheet, "C")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.RemoveCol(sheet, "D")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.RemoveCol(sheet, "D")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.RemoveCol(sheet, "D")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.RemoveCol(sheet, "D")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.RemoveCol(sheet, "D")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.RemoveCol(sheet, "G")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.RemoveCol(sheet, "H")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.RemoveCol(sheet, "H")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.RemoveCol(sheet, "H")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.RemoveCol(sheet, "H")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.RemoveCol(sheet, "H")
	if err != nil {
		fmt.Println(err)
		return
	}

	err = f.RemoveCol(sheet, "H")
	if err != nil {
		fmt.Println(err)
		return
	}

	f.SetCellValue(sheet, "H1", "Notes")

	rows, err := f.GetRows(sheet)
	if err != nil {
		fmt.Println(err)
	}

	maxLen := len(rows)

	averageLocation := "D" + strconv.Itoa(maxLen+2)

	lastRowNum := "D" + strconv.Itoa(maxLen)
	fmt.Println(lastRowNum)

	noteAverage := "AVERAGE(D2:" + lastRowNum + ")"
	fmt.Println(noteAverage)

	err = f.SetCellFormula(sheet, averageLocation, noteAverage)
	if err != nil {
		fmt.Println(err)
		return
	}

	now := time.Now()
	test := now.Format("2006-01-02")
	fileName := "DWestTickets" + test + ".xlsx"

	err = f.SaveAs(fileName)
	if err != nil {
		fmt.Println(err)
		return
	}
}
