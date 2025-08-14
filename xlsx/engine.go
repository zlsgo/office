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
	o := zutil.Optional(ReadOptions{}, opt...)
	if o.Sheet == "" {
		sheets := x.f.GetSheetList()
		if len(sheets) == 0 {
			return nil, errors.New("no sheet")
		}
		o.Sheet = sheets[0]
	}

	rows, err := x.f.GetRows(o.Sheet)
	if err != nil {
		return nil, err
	}

	minRows := o.OffsetY + 2
	if o.NoHeaderRow {
		minRows = o.OffsetY + 1
	}
	if len(rows) < minRows {
		return ztype.Maps{}, errors.New("no data")
	}

	toCol := func(i int) string {
		s := ""
		for i >= 0 {
			s = string(rune('A'+(i%26))) + s
			i = i/26 - 1
		}
		return s
	}

	var cols []string

	if o.NoHeaderRow {
		headerRow := rows[o.OffsetY]
		cols = make([]string, len(headerRow))
		for i := range headerRow {
			cols[i] = toCol(i)
		}
		rows = rows[o.OffsetY:]
	} else {
		cols = rows[o.OffsetY]
		rows = rows[o.OffsetY+1:]
	}

	if o.OffsetX > 0 {
		if o.OffsetX < len(cols) {
			cols = cols[o.OffsetX:]
		} else {
			cols = []string{}
		}
	}

	if o.Reverse {
		rows = zarray.Reverse(rows)
	}

	if o.MaxRows > 0 {
		if len(rows) > o.MaxRows {
			rows = rows[:o.MaxRows]
		}
	}

	parallel := o.Parallel
	if parallel == 0 {
		parallel = uint(len(rows) / 3000)
	}

	result := zarray.Map(rows, func(index int, row []string) ztype.Map {
		data := make(ztype.Map, len(row))

		isEmptyRow := true
		rowEffective := row
		if o.OffsetX > 0 {
			if o.OffsetX < len(row) {
				rowEffective = row[o.OffsetX:]
			} else {
				rowEffective = []string{}
			}
		}

		for j := range rowEffective {
			if j >= len(cols) {
				if o.NoHeaderRow {
					key := toCol(o.OffsetX + j)
					if len(o.Fields) > 0 && !zarray.Contains(o.Fields, key) {
						continue
					}
					data[key] = rowEffective[j]
					if isEmptyRow && rowEffective[j] != "" {
						isEmptyRow = false
					}
				}
				continue
			}

			if len(o.Fields) > 0 && !zarray.Contains(o.Fields, cols[j]) {
				continue
			}

			data[cols[j]] = rowEffective[j]
			if isEmptyRow && rowEffective[j] != "" {
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
	}, parallel)

	if o.RemoveEmptyRow {
		result = zarray.Filter(result, func(_ int, v ztype.Map) bool {
			return len(v) > 0
		})
	}

	return result, nil
}
