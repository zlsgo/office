package xlsx

import "strings"

// ToCol converts a column index to Excel-style column name (e.g., 0 -> A, 25 -> Z, 26 -> AA)
func ToCol(i int) string {
	s := ""
	for i >= 0 {
		s = string(rune('A'+(i%26))) + s
		i = i/26 - 1
	}
	return s
}

// ToColIndex converts an Excel-style column name to 0-based index (e.g., A -> 0, Z -> 25, AA -> 26)
func ToColIndex(col string) int {
	col = strings.TrimSpace(col)
	if col == "" {
		return -1
	}
	col = strings.ToUpper(col)
	n := 0
	for _, c := range col {
		if c < 'A' || c > 'Z' {
			return -1
		}
		n = n*26 + int(c-'A'+1)
	}
	return n - 1
}
