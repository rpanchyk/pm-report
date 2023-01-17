package models

type ExcelContext struct {
	NameColumn       string
	ManagerColumn    string
	PositionColumn   string
	RateColumn       string
	TotalHoursColumn string
	TotalCostColumn  string

	FirstDateColumnIndex int
	LastDateColumnIndex  int

	ColsCount    int
	LastRowIndex int

	ProjectKeyToColor map[string]string
}
