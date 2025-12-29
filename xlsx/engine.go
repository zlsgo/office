package xlsx

import (
	"errors"
	"strconv"
	"strings"

	"github.com/sohaha/zlsgo/zarray"
	"github.com/sohaha/zlsgo/zfile"
	"github.com/sohaha/zlsgo/ztype"
	"github.com/sohaha/zlsgo/zutil"
	"github.com/xuri/excelize/v2"
)

type Xlsx struct {
	f    *excelize.File
	path string
}

func Open(path string) (*Xlsx, error) {
	if path != "" {
		path = zfile.RealPath(path)
		f, err := excelize.OpenFile(path)
		if err != nil {
			if !strings.Contains(err.Error(), "no such file") {
				return nil, err
			}

			f = excelize.NewFile()
		}
		return &Xlsx{f: f, path: path}, nil
	}
	return &Xlsx{f: excelize.NewFile(), path: ""}, nil
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

	rawRows := [][]string{}
	if len(o.RawCellValueFields) > 0 && !o.RawCellValue {
		rawOpt := o.Options
		rawOpt.RawCellValue = true
		rawRows, _ = x.f.GetRows(o.Sheet, rawOpt)
		if !o.NoHeaderRow && len(rawRows) > o.OffsetY {
			rawRows = rawRows[o.OffsetY+1:]
		} else if o.NoHeaderRow && len(rawRows) > o.OffsetY {
			rawRows = rawRows[o.OffsetY:]
		}
	}

	rows, err := x.f.GetRows(o.Sheet, o.Options)
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

	var (
		cols      []string
		headerRow = rows[o.OffsetY]
	)

	colsIndex := make([]string, len(headerRow))
	for i := range headerRow {
		colsIndex[i] = ToCol(i)
	}

	if o.NoHeaderRow {
		cols = make([]string, len(headerRow))
		copy(cols, colsIndex)
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

	if o.TrimSpace {
		for i := range cols {
			cols[i] = strings.TrimSpace(cols[i])
		}
		cols = zarray.Filter(cols, func(_ int, v string) bool {
			return v != ""
		})
	}

	if o.HeaderHandler != nil {
		for i := range cols {
			cols[i] = o.HeaderHandler(colsIndex[o.OffsetX+i], cols[i])
		}
	}

	if len(o.HeaderMaps) > 0 {
		for i := range cols {
			if mapped, ok := o.HeaderMaps[cols[i]]; ok {
				cols[i] = mapped
			}
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
		rawRowEffective := []string{}
		if len(rawRows) > index {
			rawRow := rawRows[index]
			if o.OffsetX < len(rawRow) {
				rawRowEffective = rawRow[o.OffsetX:]
			}
		}

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
					key := ToCol(o.OffsetX + j)
					if len(o.Fields) > 0 && !zarray.Contains(o.Fields, key) {
						continue
					}

					value := rowEffective[j]
					// Determine if we need raw value (formula)
					needRawValue := o.RawCellValue || (len(rawRowEffective) > j && zarray.Contains(o.RawCellValueFields, key))
					if needRawValue {
						// Try rawRows first (if available), otherwise get formula
						if len(rawRowEffective) > j && rawRowEffective[j] != "" {
							value = rawRowEffective[j]
						}
						if value == "" {
							cellAddr := ToCol(o.OffsetX+j) + strconv.Itoa(index+o.OffsetY+2)
							if formula, err := x.f.GetCellFormula(o.Sheet, cellAddr); err == nil && formula != "" {
								value = formula
							}
						}
					} else if value == "" || strings.HasPrefix(value, "=") {
						// Use CalcCellValue for computed values
						cellAddr := ToCol(o.OffsetX+j) + strconv.Itoa(index+o.OffsetY+2)
						if calcVal, err := x.f.CalcCellValue(o.Sheet, cellAddr); err == nil && calcVal != "" {
							value = calcVal
						}
					}
					if o.TrimSpace {
						value = strings.TrimSpace(value)
					}

					data[key] = value
					if isEmptyRow && rowEffective[j] != "" {
						isEmptyRow = false
					}
				}
				continue
			}

			if len(o.Fields) > 0 && !zarray.Contains(o.Fields, cols[j]) {
				continue
			}

			value := rowEffective[j]
			// Determine if we need raw value (formula)
			needRawValue := o.RawCellValue || (len(rawRowEffective) > j && zarray.Contains(o.RawCellValueFields, cols[j]))
			if needRawValue {
				// Try rawRows first (if available), otherwise get formula
				if len(rawRowEffective) > j && rawRowEffective[j] != "" {
					value = rawRowEffective[j]
				}
				if value == "" {
					cellAddr := ToCol(o.OffsetX+j) + strconv.Itoa(index+o.OffsetY+2)
					if formula, err := x.f.GetCellFormula(o.Sheet, cellAddr); err == nil && formula != "" {
						value = formula
					}
				}
			} else if value == "" || strings.HasPrefix(value, "=") {
				// Use CalcCellValue for computed values
				cellAddr := ToCol(o.OffsetX+j) + strconv.Itoa(index+o.OffsetY+2)
				if calcVal, err := x.f.CalcCellValue(o.Sheet, cellAddr); err == nil && calcVal != "" {
					value = calcVal
				}
			}
			if o.TrimSpace {
				value = strings.TrimSpace(value)
			}

			data[cols[j]] = value
			if isEmptyRow && rowEffective[j] != "" {
				isEmptyRow = false
			}
		}

		if isEmptyRow {
			return ztype.Map{}
		}
		if !o.NoHeaderRow {
			if len(o.Fields) > 0 {
				for _, k := range o.Fields {
					if _, ok := data[k]; !ok {
						data[k] = nil
					}
				}
			} else {
				for _, k := range cols {
					if _, ok := data[k]; !ok {
						data[k] = nil
					}
				}
			}
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
