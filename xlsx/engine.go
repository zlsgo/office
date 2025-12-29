package xlsx

import (
	"errors"
	"strconv"
	"strings"
	"sync"

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
	var err error
	if len(o.RawCellValueFields) > 0 && !o.RawCellValue {
		rawOpt := o.Options
		rawOpt.RawCellValue = true
		rawRows, err = x.f.GetRows(o.Sheet, rawOpt)
		if err != nil {
			return nil, err
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

	rawStart := o.OffsetY
	if !o.NoHeaderRow {
		rawStart = o.OffsetY + 1
	}
	if rawStart < len(rawRows) {
		rawRows = rawRows[rawStart:]
	} else {
		rawRows = [][]string{}
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
		for i := range cols {
			if cols[i] == "" {
				cols[i] = colsIndex[o.OffsetX+i]
			}
		}
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

	dataStartRowNum := o.OffsetY + 1
	if !o.NoHeaderRow {
		dataStartRowNum = o.OffsetY + 2
	}

	type rowMeta struct {
		row    []string
		rawRow []string
		rowNum int
	}

	rowsMeta := make([]rowMeta, len(rows))
	for i, row := range rows {
		rowNum := dataStartRowNum + i
		var rawRow []string
		if i < len(rawRows) {
			rawRow = rawRows[i]
		}
		rowsMeta[i] = rowMeta{row: row, rawRow: rawRow, rowNum: rowNum}
	}

	if o.Reverse {
		rowsMeta = zarray.Reverse(rowsMeta)
	}

	if o.MaxRows > 0 {
		if len(rowsMeta) > o.MaxRows {
			rowsMeta = rowsMeta[:o.MaxRows]
		}
	}

	parallel := o.Parallel
	if parallel == 0 {
		parallel = 1
		if len(rowsMeta) > 3000 {
			parallel = uint(len(rowsMeta) / 3000)
			if parallel == 0 {
				parallel = 1
			}
		}
	}

	var formulaMu sync.Mutex

	result := zarray.Map(rowsMeta, func(index int, meta rowMeta) ztype.Map {
		row := meta.row
		data := make(ztype.Map, len(row))

		isEmptyRow := true
		rowEffective := row
		rawRowEffective := []string{}
		if o.OffsetX < len(meta.rawRow) {
			rawRowEffective = meta.rawRow[o.OffsetX:]
		}
		rowNum := meta.rowNum

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
					needRawValue := (o.RawCellValue || zarray.Contains(o.RawCellValueFields, key)) && !zarray.Contains(o.CalcCellValueFields, key)
					if needRawValue {
						if len(rawRowEffective) > j && rawRowEffective[j] != "" {
							value = rawRowEffective[j]
						}
						if value == "" {
							cellAddr := ToCol(o.OffsetX+j) + strconv.Itoa(rowNum)
							formulaMu.Lock()
							formula, err := x.f.GetCellFormula(o.Sheet, cellAddr)
							formulaMu.Unlock()
							if err == nil && formula != "" {
								value = formula
							}
						}
					} else if value == "" || strings.HasPrefix(value, "=") {
						cellAddr := ToCol(o.OffsetX+j) + strconv.Itoa(rowNum)
						formulaMu.Lock()
						calcVal, err := x.f.CalcCellValue(o.Sheet, cellAddr)
						formulaMu.Unlock()
						if err == nil && calcVal != "" {
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
			needRawValue := (o.RawCellValue || zarray.Contains(o.RawCellValueFields, cols[j])) && !zarray.Contains(o.CalcCellValueFields, cols[j])
			if needRawValue {
				if len(rawRowEffective) > j && rawRowEffective[j] != "" {
					value = rawRowEffective[j]
				}
				if value == "" {
					cellAddr := ToCol(o.OffsetX+j) + strconv.Itoa(rowNum)
					formulaMu.Lock()
					formula, err := x.f.GetCellFormula(o.Sheet, cellAddr)
					formulaMu.Unlock()
					if err == nil && formula != "" {
						value = formula
					}
				}
			} else if value == "" || strings.HasPrefix(value, "=") {
				cellAddr := ToCol(o.OffsetX+j) + strconv.Itoa(rowNum)
				formulaMu.Lock()
				calcVal, err := x.f.CalcCellValue(o.Sheet, cellAddr)
				formulaMu.Unlock()
				if err == nil && calcVal != "" {
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
	rowsMeta = nil

	if o.RemoveEmptyRow {
		result = zarray.Filter(result, func(_ int, v ztype.Map) bool {
			return len(v) > 0
		})
	}

	return result, nil
}
