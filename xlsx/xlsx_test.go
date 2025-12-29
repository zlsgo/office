package xlsx_test

import (
	"fmt"
	"math/rand"
	"os"
	"strings"
	"testing"

	"github.com/sohaha/zlsgo"
	"github.com/sohaha/zlsgo/ztype"
	"github.com/xuri/excelize/v2"
	"github.com/zlsgo/office/xlsx"
)

func TestBase(t *testing.T) {
	tt := zlsgo.NewTest(t)

	testFile := "./testdata/test_base.xlsx"
	f := excelize.NewFile()
	sheet := "TestBase"
	f.NewSheet(sheet)

	headers := []string{"Date", "Title", "Other"}
	_ = f.SetSheetRow(sheet, "A1", &headers)

	rows := [][]interface{}{
		{"2025-01-01", "Hello", "X"},
		{"2025-01-02", "World", "Y"},
	}
	for i, row := range rows {
		_ = f.SetSheetRow(sheet, fmt.Sprintf("A%d", i+2), &row)
	}

	err := f.SaveAs(testFile)
	tt.NoError(err)
	defer os.Remove(testFile)

	data, err := xlsx.Read(testFile, func(ro *xlsx.ReadOptions) {
		ro.Sheet = sheet
		ro.Handler = func(row int, data ztype.Map) ztype.Map {
			return data
		}
		ro.HeaderHandler = func(index string, col string) string {
			tt.Log(index, col)
			return col
		}
	})
	tt.NoError(err)
	tt.Equal(2, len(data))
	tt.Log(data)

	b, err := xlsx.Write(data, func(wo *xlsx.WriteOptions) {
		wo.First = []string{"Date"}
		wo.Last = []string{"Title"}
		wo.Sheet = "Test"
	})
	tt.NoError(err)
	tt.EqualTrue(len(b) > 0)

	outFile := "./testdata/test2.xlsx"
	wf, err := xlsx.Open("")
	tt.NoError(err)
	err = wf.WriteFile(outFile, data, func(wo *xlsx.WriteOptions) {
		// wo.Sheet = "Test"
		wo.CellHandler = func(sheet string, cell string, value interface{}) ([]xlsx.RichText, int) {
			color := fmt.Sprintf("%06x", rand.Intn(0xFFFFFF))
			size := 12.0
			style := 0
			if strings.HasSuffix(cell, "1") {
				size = 14.0
				style, _ = wf.NewStyle(&excelize.Style{
					Fill: excelize.Fill{Type: "pattern", Color: []string{"7030A0"}, Pattern: 1},
					Font: &excelize.Font{
						Size:      size,
						Color:     "FFFFFF",
						VertAlign: "baseline",
					},
				})
				return nil, style
			}
			return []xlsx.RichText{
				{
					Text: ztype.ToString(value),
					Font: &excelize.Font{
						Size:      size,
						Color:     color,
						VertAlign: "baseline",
					},
				},
			}, style
		}
	})
	tt.NoError(err)
	_ = os.Remove(outFile)
}

func TestOffset(t *testing.T) {
	tt := zlsgo.NewTest(t)

	testFile := "./testdata/test_offset.xlsx"
	f := excelize.NewFile()
	defer os.Remove(testFile)

	sheet1 := "TestOffset"
	f.NewSheet(sheet1)

	dataRows := [][]interface{}{
		{"", "", "", "", ""},
		{"", "", "", "", ""},
		{"", "", "", "", ""},
		{"", "Col1", "Col2", "Col3", "Col4"},
		{"", "A1", "B1", "C1", "D1"},
		{"", "A2", "B2", "C2", "D2"},
		{"", "A3", "B3", "C3", "D3"},
	}

	for i, row := range dataRows {
		f.SetSheetRow(sheet1, fmt.Sprintf("A%d", i+1), &row)
	}

	err := f.SaveAs(testFile)
	tt.NoError(err)

	data, err := xlsx.Read(testFile, func(ro *xlsx.ReadOptions) {
		ro.Sheet = sheet1
		ro.OffsetY = 3
	})
	tt.Log(data)
	tt.NoError(err)
	tt.Equal(3, len(data))

	tt.Equal("A1", data[0].Get("Col1").String())
	tt.Equal("B1", data[0].Get("Col2").String())
	tt.Equal("C1", data[0].Get("Col3").String())
	tt.Equal("A2", data[1].Get("Col1").String())
	tt.Equal("B2", data[1].Get("Col2").String())
	tt.Equal("C2", data[1].Get("Col3").String())
	tt.Equal("A3", data[2].Get("Col1").String())
	tt.Equal("B3", data[2].Get("Col2").String())
	tt.Equal("C3", data[2].Get("Col3").String())

	data, err = xlsx.Read(testFile, func(ro *xlsx.ReadOptions) {
		ro.Sheet = sheet1
		ro.OffsetX = 1
		ro.OffsetY = 3
	})
	tt.Log(data)
	tt.NoError(err)
	tt.Equal(3, len(data))

	tt.Equal("A1", data[0].Get("Col1").String())
	tt.Equal("B1", data[0].Get("Col2").String())
	tt.Equal("C1", data[0].Get("Col3").String())
	tt.Equal("A2", data[1].Get("Col1").String())
	tt.Equal("B2", data[1].Get("Col2").String())
	tt.Equal("C2", data[1].Get("Col3").String())
	tt.Equal("A3", data[2].Get("Col1").String())
	tt.Equal("B3", data[2].Get("Col2").String())
	tt.Equal("C3", data[2].Get("Col3").String())

	_, err = xlsx.Read(testFile, func(ro *xlsx.ReadOptions) {
		ro.Sheet = sheet1
		ro.OffsetX = 10
		ro.OffsetY = 10
	})
	tt.EqualTrue(err != nil)
	tt.Equal("no data", err.Error())

	data, err = xlsx.Read(testFile, func(ro *xlsx.ReadOptions) {
		ro.Sheet = sheet1
		ro.NoHeaderRow = true
	})
	tt.Log(data)
	tt.NoError(err)
	tt.Equal(7, len(data))

	tt.Equal("Col1", data[3].Get("B").String())
	tt.Equal("Col2", data[3].Get("C").String())
	tt.Equal("A1", data[4].Get("B").String())
	tt.Equal("B1", data[4].Get("C").String())
	tt.Equal("A2", data[5].Get("B").String())
	tt.Equal("B2", data[5].Get("C").String())

	data, err = xlsx.Read(testFile, func(ro *xlsx.ReadOptions) {
		ro.Sheet = sheet1
		ro.OffsetY = 100
	})
	tt.EqualTrue(err != nil)
	tt.Equal("no data", err.Error())
	tt.Equal(0, len(data))
}

func TestRawCellValueFields(t *testing.T) {
	tt := zlsgo.NewTest(t)

	testFile := "./testdata/test_raw_cell_value.xlsx"
	defer os.Remove(testFile)

	sheet := "RawCellValue"
	f := excelize.NewFile()
	f.NewSheet(sheet)

	headers := []string{"Name", "Formula", "Value", "AnotherFormula"}
	_ = f.SetSheetRow(sheet, "A1", &headers)

	_ = f.SetCellValue(sheet, "A2", "Row1")
	_ = f.SetCellValue(sheet, "C2", 100)
	_ = f.SetCellFormula(sheet, "B2", "=SUM(10,20)")
	_ = f.SetCellFormula(sheet, "D2", "=A2&\"_test\"")

	_ = f.SetCellValue(sheet, "A3", "Row2")
	_ = f.SetCellValue(sheet, "C3", 200)
	_ = f.SetCellFormula(sheet, "B3", "=A2&\"_suffix\"")
	_ = f.SetCellFormula(sheet, "D3", "=CONCATENATE(A2,\"_x\")")

	_ = f.SetCellValue(sheet, "A4", "Row3")
	_ = f.SetCellValue(sheet, "C4", 300)
	_ = f.SetCellFormula(sheet, "B4", "=UPPER(\"hello\")")
	_ = f.SetCellFormula(sheet, "D4", "=B3")

	err := f.SaveAs(testFile)
	tt.NoError(err)

	data, err := xlsx.Read(testFile, func(ro *xlsx.ReadOptions) {
		ro.Sheet = sheet
	})
	tt.NoError(err)
	tt.Equal(3, len(data))
	tt.Equal("30", data[0].Get("Formula").String())
	tt.Equal("Row1_suffix", data[1].Get("Formula").String())
	tt.Equal("HELLO", data[2].Get("Formula").String())

	data, err = xlsx.Read(testFile, func(ro *xlsx.ReadOptions) {
		ro.Sheet = sheet
		ro.RawCellValueFields = []string{"Formula", "AnotherFormula"}
	})
	tt.NoError(err)
	tt.Equal(3, len(data))
	tt.Equal("=SUM(10,20)", data[0].Get("Formula").String())
	tt.Equal("=A2&\"_suffix\"", data[1].Get("Formula").String())
	tt.Equal("=UPPER(\"hello\")", data[2].Get("Formula").String())
	tt.Equal("=A2&\"_test\"", data[0].Get("AnotherFormula").String())
	tt.Equal("=CONCATENATE(A2,\"_x\")", data[1].Get("AnotherFormula").String())
	tt.Equal("=B3", data[2].Get("AnotherFormula").String())
	tt.Equal(100, data[0].Get("Value").Int())
	tt.Equal(200, data[1].Get("Value").Int())

	data, err = xlsx.Read(testFile, func(ro *xlsx.ReadOptions) {
		ro.Sheet = sheet
		ro.RawCellValue = true
	})
	tt.NoError(err)
	tt.Equal(3, len(data))
	tt.Equal("=SUM(10,20)", data[0].Get("Formula").String())
	tt.Equal("=A2&\"_suffix\"", data[1].Get("Formula").String())
	tt.Equal("=UPPER(\"hello\")", data[2].Get("Formula").String())
	tt.Equal("=A2&\"_test\"", data[0].Get("AnotherFormula").String())
	tt.Equal("=CONCATENATE(A2,\"_x\")", data[1].Get("AnotherFormula").String())
	tt.Equal("=B3", data[2].Get("AnotherFormula").String())
	tt.Equal(100, data[0].Get("Value").Int())
	tt.Equal(200, data[1].Get("Value").Int())
}

func TestRawCellValueFieldsNoHeaderOffset(t *testing.T) {
	tt := zlsgo.NewTest(t)

	testFile := "./testdata/test_raw_cell_value_no_header.xlsx"
	defer os.Remove(testFile)

	sheet := "RawNoHeader"
	f := excelize.NewFile()
	f.NewSheet(sheet)

	_ = f.SetCellValue(sheet, "A1", "Skip")
	_ = f.SetCellValue(sheet, "B1", "Skip")

	_ = f.SetCellValue(sheet, "A2", "Row1")
	_ = f.SetCellFormula(sheet, "B2", "=1+1")

	_ = f.SetCellValue(sheet, "A3", "Row2")
	_ = f.SetCellFormula(sheet, "B3", "=2+2")

	err := f.SaveAs(testFile)
	tt.NoError(err)

	data, err := xlsx.Read(testFile, func(ro *xlsx.ReadOptions) {
		ro.Sheet = sheet
		ro.NoHeaderRow = true
		ro.OffsetY = 1
		ro.RawCellValueFields = []string{"B"}
	})
	tt.NoError(err)
	tt.Equal(2, len(data))
	tt.Equal("=1+1", data[0].Get("B").String())
	tt.Equal("=2+2", data[1].Get("B").String())
}

func TestRawCellValueFieldsReverse(t *testing.T) {
	tt := zlsgo.NewTest(t)

	testFile := "./testdata/test_raw_cell_value_reverse.xlsx"
	defer os.Remove(testFile)

	sheet := "RawReverse"
	f := excelize.NewFile()
	f.NewSheet(sheet)

	headers := []string{"Name", "Formula"}
	_ = f.SetSheetRow(sheet, "A1", &headers)

	_ = f.SetCellValue(sheet, "A2", "R1")
	_ = f.SetCellFormula(sheet, "B2", "=1+1")
	_ = f.SetCellValue(sheet, "A3", "R2")
	_ = f.SetCellFormula(sheet, "B3", "=2+2")
	_ = f.SetCellValue(sheet, "A4", "R3")
	_ = f.SetCellFormula(sheet, "B4", "=3+3")

	err := f.SaveAs(testFile)
	tt.NoError(err)

	data, err := xlsx.Read(testFile, func(ro *xlsx.ReadOptions) {
		ro.Sheet = sheet
		ro.Reverse = true
		ro.RawCellValueFields = []string{"Formula"}
	})
	tt.NoError(err)
	tt.Equal(3, len(data))
	tt.Equal("R3", data[0].Get("Name").String())
	tt.Equal("=3+3", data[0].Get("Formula").String())
	tt.Equal("R2", data[1].Get("Name").String())
	tt.Equal("=2+2", data[1].Get("Formula").String())
	tt.Equal("R1", data[2].Get("Name").String())
	tt.Equal("=1+1", data[2].Get("Formula").String())
}

func TestCalcCellValueFields(t *testing.T) {
	tt := zlsgo.NewTest(t)

	testFile := "./testdata/test_calc_cell_value.xlsx"
	defer os.Remove(testFile)

	sheet := "CalcFields"
	f := excelize.NewFile()
	f.NewSheet(sheet)

	headers := []string{"Name", "Formula1", "Value", "Formula2"}
	_ = f.SetSheetRow(sheet, "A1", &headers)

	_ = f.SetCellValue(sheet, "A2", "Row1")
	_ = f.SetCellFormula(sheet, "B2", "=10+20")
	_ = f.SetCellValue(sheet, "C2", 100)
	_ = f.SetCellFormula(sheet, "D2", "=5*5")

	_ = f.SetCellValue(sheet, "A3", "Row2")
	_ = f.SetCellFormula(sheet, "B3", "=30+40")
	_ = f.SetCellValue(sheet, "C3", 200)
	_ = f.SetCellFormula(sheet, "D3", "=6*6")

	err := f.SaveAs(testFile)
	tt.NoError(err)

	data, err := xlsx.Read(testFile, func(ro *xlsx.ReadOptions) {
		ro.Sheet = sheet
		ro.RawCellValue = true
		ro.CalcCellValueFields = []string{"Formula1", "Value"}
	})
	tt.NoError(err)
	tt.Equal(2, len(data))

	tt.Equal("30", data[0].Get("Formula1").String())
	tt.Equal("70", data[1].Get("Formula1").String())
	tt.Equal(100, data[0].Get("Value").Int())
	tt.Equal(200, data[1].Get("Value").Int())

	tt.Equal("=5*5", data[0].Get("Formula2").String())
	tt.Equal("=6*6", data[1].Get("Formula2").String())
}
