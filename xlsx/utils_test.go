package xlsx

import (
	"strings"
	"testing"
)

func TestToCol(t *testing.T) {
	for i := 0; i < 100; i++ {
		t.Log(ToCol(i))
	}
}

func TestToColIndex(t *testing.T) {
	cases := map[string]int{
		"A":   0,
		"Z":   25,
		"AA":  26,
		"AB":  27,
		"AZ":  51,
		"BA":  52,
		"ZZ":  701,
		"AAA": 702,
		"XFD": 16383, // Excel 最大列（2007+）
	}

	for s, want := range cases {
		got := ToColIndex(s)
		if got != want {
			t.Fatalf("ToColIndex(%q) = %d, want %d", s, got, want)
		}
		back := ToCol(got)
		if back != strings.ToUpper(s) {
			t.Fatalf("ToCol(ToColIndex(%q)) = %q, want %q", s, back, strings.ToUpper(s))
		}
	}

	for i := 0; i < 5000; i++ {
		s := ToCol(i)
		if got := ToColIndex(s); got != i {
			t.Fatalf("ToColIndex(ToCol(%d)) = %d, want %d (s=%q)", i, got, i, s)
		}
	}

	bad := []string{"", "A0", "-A", "!", "中文", "A B"}
	for _, s := range bad {
		if got := ToColIndex(s); got != -1 {
			t.Fatalf("ToColIndex(%q) = %d, want -1", s, got)
		}
	}
}
