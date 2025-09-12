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
	wf, err := xlsx.Open(outFile)
	tt.NoError(err)
	err = wf.WriteFile(data, func(wo *xlsx.WriteOptions) {
		// wo.Sheet = "Test"
		wo.CellHandler = func(sheet string, cell string, value interface{}) ([]xlsx.RichText, int) {
			color := fmt.Sprintf("%06x", rand.Intn(0xFFFFFF))
			size := 12.0
			style := 0
			if strings.HasSuffix(cell, "1") {
				size = 14.0
				style, _ = wf.NewStyle(&excelize.Style{
					Fill: excelize.Fill{Type: "pattern", Color: []string{"7030A0"}, Shading: 1},
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
	// _ = os.Remove(outFile)
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
