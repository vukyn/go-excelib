package config

const (
	DEFAULT_TITLE       = "Export Example Excel Title"
	DEFAULT_SHEET_NAME  = "Sheet1"
	DEFAULT_TABLE_NAME  = "Table1"
	DEFAULT_FILE_NAME   = "Excel1"
	DEFAULT_FILE_PATH   = "./tmp"
	DEFAULT_INDEX_NAME  = "No."
	DEFAULT_TABLE_STYLE = "TableStyleLight9"
	OFFSET              = 6
	MAX_ROW             = 1000000
)

var (
	EXCEL_COLUMN = []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ"}
)

type ExportType uint

const (
	XLSX ExportType = iota + 1
	CSV
)
