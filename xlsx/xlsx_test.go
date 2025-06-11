package xlsx_test

import (
	"os"
	"testing"

	"github.com/sohaha/zlsgo"
	"github.com/sohaha/zlsgo/ztype"
	"github.com/zlsgo/office/xlsx"
)

func TestNew(t *testing.T) {
	tt := zlsgo.NewTest(t)

	data, err := xlsx.Read("./testdata/test.xlsx", func(ro *xlsx.ReadOptions) {
		ro.Handler = func(index int, data ztype.Map) ztype.Map {
			return data
		}
	})
	tt.NoError(err)
	tt.Log(data)

	b, err := xlsx.Write(data, func(wo *xlsx.WriteOptions) {
		wo.First = []string{"Date"}
		wo.Last = []string{"Title"}
	})
	tt.NoError(err)
	tt.EqualTrue(len(b) > 0)

	err = xlsx.WriteFile("./testdata/test2.xlsx", data, func(wo *xlsx.WriteOptions) {
		wo.Sheet = "Test"
	})
	tt.NoError(err)
	_ = os.Remove("./testdata/test2.xlsx")
}
