package xlsx

import (
	"errors"
	"strconv"

	"github.com/sohaha/zlsgo/zarray"
	"github.com/sohaha/zlsgo/zfile"
	"github.com/sohaha/zlsgo/ztype"
	"github.com/sohaha/zlsgo/zutil"
	"github.com/xuri/excelize/v2"
)

type ReadOptions struct {
	Handler        func(row int, data ztype.Map) ztype.Map
	HeaderHandler  func(index string, col string) string
	Sheet          string
	Fields         []string
	Reverse        bool
	Parallel       uint
	OffsetX        int
	OffsetY        int
	NoHeaderRow    bool
	MaxRows        int
	RemoveEmptyRow bool
	TrimSpace      bool
	HeaderMaps     map[string]string

	excelize.Options
}

// Read read xlsx file
func Read(path string, opt ...func(*ReadOptions)) (ztype.Maps, error) {
	f, err := Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	return f.Read(opt...)
}

type (
	RichText     excelize.RichTextRun
	WriteOptions struct {
		Sheet       string
		First       []string
		Last        []string
		CellHandler func(sheet string, cell string, value interface{}) ([]RichText, int)
	}
)

// Write write xlsx file
func write(f *excelize.File, data ztype.Maps, opt ...func(*WriteOptions)) error {
	if len(data) == 0 {
		return errors.New("no data")
	}

	o := zutil.Optional(WriteOptions{Sheet: "Sheet1"}, opt...)
	header := zarray.SortWithPriority(zarray.Keys(data[0]), o.First, o.Last)
	headerSize := len(header)

	index, err := f.NewSheet(o.Sheet)
	if err != nil {
		return err
	}
	f.SetActiveSheet(index)

	err = f.SetSheetRow(o.Sheet, "A1", &header)
	if err != nil {
		return err
	}

	if o.CellHandler != nil {
		for i := range header {
			cell := ToCol(i) + "1"
			richTextRuns, styleID := o.CellHandler(o.Sheet, cell, header[i])
			if styleID > 0 {
				_ = f.SetCellStyle(o.Sheet, cell, cell, styleID)
			}
			if richTextRuns == nil {
				continue
			}
			excelizeRuns := make([]excelize.RichTextRun, len(richTextRuns))
			for i, rt := range richTextRuns {
				excelizeRuns[i] = excelize.RichTextRun(rt)
			}
			f.SetCellRichText(o.Sheet, cell, excelizeRuns)
		}
	}

	for i := range data {
		value := make([]interface{}, 0, headerSize)
		for j := range header {
			value = append(value, data[i][header[j]])
		}
		f.SetSheetRow(o.Sheet, "A"+strconv.Itoa(i+2), &value)
		if o.CellHandler != nil {
			for j := range value {
				cell := ToCol(j) + strconv.Itoa(i+2)
				richTextRuns, styleID := o.CellHandler(o.Sheet, cell, value[j])
				if styleID > 0 {
					_ = f.SetCellStyle(o.Sheet, cell, cell, styleID)
				}
				if richTextRuns == nil {
					continue
				}
				excelizeRuns := make([]excelize.RichTextRun, len(richTextRuns))
				for i, rt := range richTextRuns {
					excelizeRuns[i] = excelize.RichTextRun(rt)
				}
				f.SetCellRichText(o.Sheet, cell, excelizeRuns)
			}
		}
	}

	return nil
}

func (x *Xlsx) Write(data ztype.Maps, opt ...func(*WriteOptions)) ([]byte, error) {
	err := write(x.f, data, opt...)
	if err != nil {
		return nil, err
	}

	b, err := x.f.WriteToBuffer()
	if err != nil {
		return nil, err
	}

	return b.Bytes(), nil
}

func (x *Xlsx) WriteFile(data ztype.Maps, opt ...func(*WriteOptions)) error {
	b, err := x.Write(data, opt...)
	if err != nil {
		return err
	}
	if x.path != "" {
		return zfile.WriteFile(x.path, b)
	}

	return x.f.Save()
}

// WriteFile write xlsx file
func WriteFile(path string, data ztype.Maps, opt ...func(*WriteOptions)) error {
	b, err := Write(data, opt...)
	if err != nil {
		return err
	}
	return zfile.WriteFile(path, b)
}

func Write(data ztype.Maps, opt ...func(*WriteOptions)) ([]byte, error) {
	f := excelize.NewFile()
	defer f.Close()
	err := write(f, data, opt...)
	if err != nil {
		return nil, err
	}

	b, err := f.WriteToBuffer()
	if err != nil {
		return []byte{}, err
	}

	return b.Bytes(), nil
}

func (x *Xlsx) NewStyle(style *excelize.Style) (int, error) {
	return x.f.NewStyle(style)
}
