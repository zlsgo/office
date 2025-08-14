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
	Handler        func(index int, data ztype.Map) ztype.Map
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

type WriteOptions struct {
	Sheet string
	First []string
	Last  []string
}

// Write write xlsx file
func Write(data ztype.Maps, opt ...func(*WriteOptions)) ([]byte, error) {
	f := excelize.NewFile()
	defer f.Close()

	if len(data) == 0 {
		return []byte{}, errors.New("no data")
	}

	o := zutil.Optional(WriteOptions{Sheet: "Sheet1"}, opt...)
	header := zarray.SortWithPriority(zarray.Keys(data[0]), o.First, o.Last)
	headerSize := len(header)

	index, err := f.NewSheet(o.Sheet)
	if err != nil {
		return []byte{}, err
	}
	f.SetActiveSheet(index)

	err = f.SetSheetRow(o.Sheet, "A1", &header)
	if err != nil {
		return []byte{}, err
	}

	for i := range data {
		value := make([]interface{}, 0, headerSize)
		for j := range header {
			value = append(value, data[i][header[j]])
		}
		f.SetSheetRow(o.Sheet, "A"+strconv.Itoa(i+2), &value)
	}

	b, err := f.WriteToBuffer()
	if err != nil {
		return []byte{}, err
	}

	return b.Bytes(), nil
}

// WriteFile write xlsx file
func WriteFile(path string, data ztype.Maps, opt ...func(*WriteOptions)) error {
	b, err := Write(data, opt...)
	if err != nil {
		return err
	}
	return zfile.WriteFile(path, b)
}
