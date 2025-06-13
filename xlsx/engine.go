package xlsx

import (
	"errors"

	"github.com/sohaha/zlsgo/zarray"
	"github.com/sohaha/zlsgo/ztype"
	"github.com/sohaha/zlsgo/zutil"
	"github.com/xuri/excelize/v2"
)

type Xlsx struct {
	f *excelize.File
}

func Open(path string) (*Xlsx, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	return &Xlsx{f: f}, nil
}

func (x *Xlsx) Close() error {
	return x.f.Close()
}

func (x *Xlsx) Engine() *excelize.File {
	return x.f
}

func (x *Xlsx) Read(opt ...func(*ReadOptions)) (ztype.Maps, error) {
	o := zutil.Optional(ReadOptions{Sheet: "Sheet1"}, opt...)
	rows, err := x.f.GetRows(o.Sheet)
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
