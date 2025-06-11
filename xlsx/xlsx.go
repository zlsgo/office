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
	Handler func(index int, data ztype.Map) ztype.Map
	Sheet   string
	Fields  []string
	Reverse bool
}

// Read read xlsx file
func Read(path string, opt ...func(*ReadOptions)) (ztype.Maps, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	o := zutil.Optional(ReadOptions{Sheet: "Sheet1"}, opt...)
	rows, err := f.GetRows(o.Sheet)
	if err != nil {
		return nil, err
	}

	if len(rows) < 2 {
		return ztype.Maps{}, errors.New("no data")
	}

	cols := rows[0]
	rows = rows[1:]

	if o.Reverse {
		rows = zarray.Reverse(rows)
	}

	parallel := uint(len(rows) / 3000)

	return zarray.Filter(zarray.Map(rows, func(index int, row []string) ztype.Map {
		data := make(ztype.Map, len(row))

		isEmptyRow := true
		for j := range row {
			if j >= len(cols) {
				continue
			}

			if len(o.Fields) > 0 && !zarray.Contains(o.Fields, cols[j]) {
				continue
			}

			data[cols[j]] = row[j]
			if isEmptyRow && row[j] != "" {
				isEmptyRow = false
			}
		}

		if isEmptyRow {
			return ztype.Map{}
		}

		if o.Handler != nil {
			return o.Handler(index, data)
		}

		return data
	}, parallel), func(index int, item ztype.Map) bool {
		return len(item) > 0
	}), nil
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
