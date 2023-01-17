package services

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"math"
	"math/big"
	"os"
	"pm-report/models"
	"strconv"
	"strings"
	"time"
)

const (
	effortDateFormat = "2006-01-02"
	columnDateFormat = "%02d/%02d"
)

type ExcelService struct {
	filePath string
}

func NewExcelService(filePath string) *ExcelService {
	return &ExcelService{filePath: filePath}
}

func (s *ExcelService) Save(report *models.Report) error {
	f, err := s.createOrOpenFile()
	if err != nil {
		return err
	}

	sheet := s.getSheetName(report)
	context := s.createContext(report)

	err = s.createSheet(f, sheet)
	if err != nil {
		return err
	}

	err = s.fillHeader(f, sheet, context, report)
	if err != nil {
		return err
	}

	err = s.fillBody(f, sheet, context, report)
	if err != nil {
		return err
	}

	err = f.SaveAs(s.filePath)
	if err != nil {
		return err
	}

	err = f.Close()
	if err != nil {
		return err
	}

	return nil
}

func (s *ExcelService) createOrOpenFile() (*excelize.File, error) {
	var f *excelize.File

	_, err := os.Stat(s.filePath)
	if err != nil && errors.Is(err, os.ErrNotExist) {
		f = excelize.NewFile()
	} else {
		f, err = excelize.OpenFile(s.filePath)
		if err != nil {
			return nil, err
		}
	}

	return f, nil
}

func (s *ExcelService) getSheetName(report *models.Report) string {
	dateFrom := report.DateFrom
	dateTo := report.DateTo

	if dateFrom.Month() == dateTo.Month() {
		return dateFrom.Format("January")
	}

	return dateFrom.Format("January") + " - " + dateTo.Format("January")
}

func (s *ExcelService) createContext(report *models.Report) *models.ExcelContext {
	return &models.ExcelContext{
		NameColumn:       "A",
		ManagerColumn:    "B",
		PositionColumn:   "C",
		RateColumn:       "D",
		TotalHoursColumn: "E",
		TotalCostColumn:  "F",

		LastRowIndex: 1,

		ProjectKeyToColor: s.getProjectKeyToColor(report.Projects),
	}
}

func (s *ExcelService) getProjectKeyToColor(projects []models.Project) map[string]string {
	// list of hardcoded colors for projects in report
	projectColors := []string{"#46bdc6", "#92d050", "#b6d7a8", "#d9ead3"}

	managerToColor := map[string]string{}
	colorIndex := 0
	for _, project := range projects {
		if len(project.Manager) == 0 {
			continue
		}

		if _, ok := managerToColor[project.Manager]; !ok {
			managerToColor[project.Manager] = projectColors[colorIndex]

			colorIndex++
			if colorIndex >= len(projectColors) {
				colorIndex = 0
			}
		}
	}

	projectKeyToColor := map[string]string{}

	for _, project := range projects {
		if color, ok := managerToColor[project.Manager]; ok {
			projectKeyToColor[project.Key] = color
		}
	}

	return projectKeyToColor
}

func (s *ExcelService) createSheet(f *excelize.File, sheet string) error {
	sheetIndex := f.NewSheet(sheet)
	f.SetActiveSheet(sheetIndex)

	if f.GetSheetIndex("Sheet1") != -1 {
		f.DeleteSheet("Sheet1")
	}

	err := f.SetPanes(sheet, `{"freeze": true, "x_split": 6, "y_split": 1}`)
	if err != nil {
		return err
	}

	return nil
}

func (s *ExcelService) fillHeader(f *excelize.File, sheet string, context *models.ExcelContext, report *models.Report) error {
	rowIndex := strconv.Itoa(context.LastRowIndex)

	// name
	err := f.SetCellValue(sheet, context.NameColumn+rowIndex, "Name")
	if err != nil {
		return err
	}
	err = f.SetColWidth(sheet, context.NameColumn, context.NameColumn, 25)
	if err != nil {
		return err
	}

	// manager
	err = f.SetCellValue(sheet, context.ManagerColumn+rowIndex, "Manager")
	if err != nil {
		return err
	}
	err = f.SetColWidth(sheet, context.ManagerColumn, context.ManagerColumn, 25)
	if err != nil {
		return err
	}

	// position
	err = f.SetCellValue(sheet, context.PositionColumn+rowIndex, "Position")
	if err != nil {
		return err
	}
	err = f.SetColWidth(sheet, context.PositionColumn, context.PositionColumn, 25)
	if err != nil {
		return err
	}

	// rate
	err = f.SetCellValue(sheet, context.RateColumn+rowIndex, "Rate")
	if err != nil {
		return err
	}
	err = f.SetColWidth(sheet, context.RateColumn, context.RateColumn, 9)
	if err != nil {
		return err
	}

	// total hours
	err = f.SetCellValue(sheet, context.TotalHoursColumn+rowIndex, "Total hours")
	if err != nil {
		return err
	}
	err = f.SetColWidth(sheet, context.TotalHoursColumn, context.TotalHoursColumn, 12)
	if err != nil {
		return err
	}

	// total cost
	err = f.SetCellValue(sheet, context.TotalCostColumn+rowIndex, "Total cost")
	if err != nil {
		return err
	}
	err = f.SetColWidth(sheet, context.TotalCostColumn, context.TotalCostColumn, 12)
	if err != nil {
		return err
	}

	// common
	alignment := excelize.Alignment{Horizontal: "center"}
	font := excelize.Font{Bold: true}
	style, err := f.NewStyle(&excelize.Style{Alignment: &alignment, Font: &font})
	if err != nil {
		return err
	}
	err = f.SetRowStyle(sheet, context.LastRowIndex, context.LastRowIndex, style)
	if err != nil {
		return err
	}

	colsCount, err := s.getColsCountInRow(f, sheet, context.LastRowIndex)
	if err != nil {
		return err
	}

	// dates
	colIndex := *colsCount
	context.FirstDateColumnIndex = colIndex + 1

	for date := report.DateFrom; !date.After(report.DateTo); date = date.AddDate(0, 0, 1) {
		colIndex++

		cell, err := excelize.CoordinatesToCellName(colIndex, context.LastRowIndex)
		if err != nil {
			return err
		}

		err = f.SetCellValue(sheet, cell, fmt.Sprintf(columnDateFormat, date.Month(), date.Day()))
		if err != nil {
			return err
		}

		col := strings.TrimRight(cell, rowIndex)

		err = f.SetColWidth(sheet, col, col, 6)
		if err != nil {
			return err
		}

		style, err = f.NewStyle(&excelize.Style{Alignment: &alignment, Font: &font})
		if err != nil {
			return err
		}
		err = f.SetCellStyle(sheet, col+rowIndex, col+rowIndex, style)
		if err != nil {
			return err
		}

		weekday := date.Weekday()
		if weekday == 0 || weekday == 6 {
			fill := excelize.Fill{Color: []string{"#FEC7CE"}, Type: "pattern", Pattern: 3}
			style, err = f.NewStyle(&excelize.Style{Alignment: &alignment, Font: &font, Fill: fill})
			if err != nil {
				return err
			}
			err = f.SetCellStyle(sheet, col+rowIndex, col+rowIndex, style)
			if err != nil {
				return err
			}
		}
	}
	context.LastDateColumnIndex = colIndex

	colsCount, err = s.getColsCountInRow(f, sheet, context.LastRowIndex)
	if err != nil {
		return err
	}
	context.ColsCount = *colsCount

	return nil
}

func (s *ExcelService) getColsCountInRow(f *excelize.File, sheet string, rowIndex int) (*int, error) {
	rows, err := f.Rows(sheet)
	if err != nil {
		return nil, err
	}

	if !rows.Next() {
		return nil, errors.New("no rows")
	}

	rowColumns, err := rows.Columns()
	if err != nil {
		return nil, err
	}

	if err = rows.Close(); err != nil {
		return nil, err
	}

	result := len(rowColumns)
	return &result, nil
}

func (s *ExcelService) fillBody(f *excelize.File, sheet string, context *models.ExcelContext, report *models.Report) error {
	for _, project := range report.Projects {
		log.Println("Creating report for project:", project.Key)

		err := s.fillProjectRow(f, sheet, &project, context)
		if err != nil {
			return err
		}

		err = s.prepareUserRows(f, sheet, &project, context)
		if err != nil {
			return err
		}

		for _, user := range project.Users {
			log.Println("Processing efforts for user:", user.Name, fmt.Sprintf("(%d)", len(user.Issues)))

			err = s.fillUserRow(f, sheet, &user, project.Key, context)
			if err != nil {
				return err
			}
		}
	}

	return nil
}

func (s *ExcelService) fillProjectRow(f *excelize.File, sheet string, project *models.Project, context *models.ExcelContext) error {
	context.LastRowIndex++
	noneBorders := []excelize.Border{
		{Type: "left", Color: "#FFFFFF", Style: 1},
	}
	style, err := f.NewStyle(&excelize.Style{Border: noneBorders})
	if err != nil {
		return err
	}
	cell, err := excelize.CoordinatesToCellName(context.ColsCount, context.LastRowIndex)
	if err != nil {
		return err
	}
	err = f.SetCellStyle(sheet, context.NameColumn+strconv.Itoa(context.LastRowIndex), cell, style)
	if err != nil {
		return err
	}

	context.LastRowIndex++
	rowIndex := strconv.Itoa(context.LastRowIndex)

	font := excelize.Font{Bold: true}
	style, err = f.NewStyle(&excelize.Style{Font: &font})
	if err != nil {
		return err
	}
	ownerStyle, err := f.NewStyle(&excelize.Style{Font: &font, Alignment: &excelize.Alignment{Horizontal: "center"}})
	if err != nil {
		return err
	}
	totalCostStyle, err := f.NewStyle(&excelize.Style{Font: &font, NumFmt: 177})
	if err != nil {
		return err
	}

	if color, ok := context.ProjectKeyToColor[project.Key]; ok {
		borderColor := s.shadeColor(color, 15)
		borders := []excelize.Border{
			{Type: "top", Color: borderColor, Style: 1},
			{Type: "left", Color: borderColor, Style: 1},
			{Type: "bottom", Color: "#444444", Style: 1},
			{Type: "right", Color: borderColor, Style: 1},
		}
		shadedColor := s.shadeColor(color, 10)
		fill := excelize.Fill{Color: []string{shadedColor}, Type: "pattern", Pattern: 1}
		style, err = f.NewStyle(&excelize.Style{Border: borders, Font: &font, Fill: fill})
		if err != nil {
			return err
		}
		ownerStyle, err = f.NewStyle(&excelize.Style{Border: borders, Font: &font, Fill: fill, Alignment: &excelize.Alignment{Horizontal: "center"}})
		if err != nil {
			return err
		}
		totalCostStyle, err = f.NewStyle(&excelize.Style{Border: borders, Font: &font, Fill: fill, NumFmt: 177})
		if err != nil {
			return err
		}
	}
	cell, err = excelize.CoordinatesToCellName(context.ColsCount, context.LastRowIndex)
	if err != nil {
		return err
	}
	err = f.SetCellStyle(sheet, context.NameColumn+rowIndex, cell, style)
	if err != nil {
		return err
	}

	err = f.SetCellStyle(sheet, context.TotalCostColumn+rowIndex, context.TotalCostColumn+rowIndex, totalCostStyle)
	if err != nil {
		return err
	}

	name := project.Key
	if len(project.DisplayName) > 0 {
		name = project.DisplayName
	}
	err = f.SetCellValue(sheet, context.NameColumn+rowIndex, name)
	if err != nil {
		return err
	}

	//conditionalFormat, err := s.createZeroValueConditionalFormat(f)
	//if err != nil {
	//	return err
	//}
	//err = f.SetConditionalFormat(sheet, context.ManagerColumn+rowIndex, *conditionalFormat)
	//if err != nil {
	//	return err
	//}
	err = f.SetCellStyle(sheet, context.ManagerColumn+rowIndex, context.ManagerColumn+rowIndex, ownerStyle)
	if err != nil {
		return err
	}
	err = f.SetCellValue(sheet, context.ManagerColumn+rowIndex, project.Manager)
	if err != nil {
		return err
	}
	//err = f.SetConditionalFormat(sheet, context.PositionColumn+rowIndex, *conditionalFormat)
	//if err != nil {
	//	return err
	//}
	err = f.MergeCell(sheet, context.PositionColumn+rowIndex, context.RateColumn+rowIndex)
	if err != nil {
		return err
	}
	err = f.SetCellStyle(sheet, context.PositionColumn+rowIndex, context.PositionColumn+rowIndex, ownerStyle)
	if err != nil {
		return err
	}
	err = f.SetCellValue(sheet, context.PositionColumn+rowIndex, project.Owner)
	if err != nil {
		return err
	}

	firstRowIndex := strconv.Itoa(context.LastRowIndex + 1)
	lastRowIndex := strconv.Itoa(context.LastRowIndex + len(project.Users))

	formula := "sum(" + context.TotalHoursColumn + firstRowIndex + ":" + context.TotalHoursColumn + lastRowIndex + ")"
	err = f.SetCellFormula(sheet, context.TotalHoursColumn+rowIndex, formula)
	if err != nil {
		return err
	}

	formula = "sum(" + context.TotalCostColumn + firstRowIndex + ":" + context.TotalCostColumn + lastRowIndex + ")"
	err = f.SetCellFormula(sheet, context.TotalCostColumn+rowIndex, formula)
	if err != nil {
		return err
	}

	return nil
}

func (s *ExcelService) prepareUserRows(f *excelize.File, sheet string, project *models.Project, context *models.ExcelContext) error {
	rowIndex := strconv.Itoa(context.LastRowIndex + 1)
	lastUserRowIndex := strconv.Itoa(context.LastRowIndex + len(project.Users))

	var borders []excelize.Border
	var fill excelize.Fill

	if color, ok := context.ProjectKeyToColor[project.Key]; ok {
		borderColor := s.shadeColor(color, 15)

		borders = []excelize.Border{
			{Type: "top", Color: borderColor, Style: 1},
			{Type: "left", Color: borderColor, Style: 1},
			{Type: "bottom", Color: borderColor, Style: 1},
			{Type: "right", Color: borderColor, Style: 1},
		}
		fill = excelize.Fill{Color: []string{color}, Type: "pattern", Pattern: 1}
	}

	style, err := f.NewStyle(&excelize.Style{Border: borders, Fill: fill})
	if err != nil {
		return err
	}
	cell, err := excelize.CoordinatesToCellName(context.ColsCount, context.LastRowIndex+len(project.Users))
	if err != nil {
		return err
	}
	err = f.SetCellStyle(sheet, context.NameColumn+rowIndex, cell, style)
	if err != nil {
		return err
	}

	alignment := excelize.Alignment{Horizontal: "center"}
	style, err = f.NewStyle(&excelize.Style{Border: borders, Fill: fill, Alignment: &alignment})
	if err != nil {
		return err
	}
	firstDateCell, err := excelize.CoordinatesToCellName(context.FirstDateColumnIndex, context.LastRowIndex+1)
	if err != nil {
		return err
	}
	err = f.SetCellStyle(sheet, firstDateCell, cell, style)
	if err != nil {
		return err
	}

	style, err = f.NewStyle(&excelize.Style{Border: borders, Fill: fill, NumFmt: 177})
	if err != nil {
		return err
	}
	err = f.SetCellStyle(sheet, context.RateColumn+rowIndex, context.RateColumn+lastUserRowIndex, style)
	if err != nil {
		return err
	}
	err = f.SetCellStyle(sheet, context.TotalCostColumn+rowIndex, context.TotalCostColumn+lastUserRowIndex, style)
	if err != nil {
		return err
	}

	return nil
}

func (s *ExcelService) shadeColor(color string, percent int64) string {
	if len(color) == 0 {
		return color
	}

	r := s.shadeColorPart(color[1:3], percent)
	g := s.shadeColorPart(color[3:5], percent)
	b := s.shadeColorPart(color[5:7], percent)

	return "#" + r + g + b
}

// https://stackoverflow.com/questions/5560248/programmatically-lighten-or-darken-a-hex-color-or-rgb-and-blend-colors
func (s *ExcelService) shadeColorPart(colorPart string, percent int64) string {
	number := new(big.Int)
	number.SetString(colorPart, 16)

	shaded := number.Int64() * (100 - percent) / 100
	if shaded <= 0 {
		shaded = 0
	}

	result := fmt.Sprintf("%x", shaded)
	if len(result) == 1 {
		result = "0" + result
	}
	return result
}

func (s *ExcelService) createZeroValueConditionalFormat(f *excelize.File) (*string, error) {
	format, err := f.NewConditionalStyle(`{
		"font": {
			"color": "#9A0511"
		},
		"fill": {
			"type": "pattern",
			"color": ["#FEC7CE"],
			"pattern": 1
		}
	}`)
	if err != nil {
		return nil, err
	}

	formatSet := fmt.Sprintf(`[{ "format": %d, "type": "cell", "criteria": "=", "value": "0" }]`, format)
	return &formatSet, nil
}

func (s *ExcelService) fillUserRow(f *excelize.File, sheet string, user *models.User, projectKey string, context *models.ExcelContext) error {
	context.LastRowIndex++
	rowIndex := strconv.Itoa(context.LastRowIndex)

	err := f.SetCellValue(sheet, context.NameColumn+rowIndex, user.Name)
	if err != nil {
		return err
	}

	err = f.SetCellValue(sheet, context.PositionColumn+rowIndex, user.Position)
	if err != nil {
		return err
	}

	conditionalFormat, err := s.createZeroValueConditionalFormat(f)
	if err != nil {
		return err
	}
	err = f.SetConditionalFormat(sheet, context.RateColumn+rowIndex, *conditionalFormat)
	if err != nil {
		return err
	}
	err = f.SetCellValue(sheet, context.RateColumn+rowIndex, user.Rate)
	if err != nil {
		return err
	}

	// user total hours
	firstDateColumnName, err := excelize.ColumnNumberToName(context.FirstDateColumnIndex)
	if err != nil {
		return err
	}

	lastDateColumnName, err := excelize.ColumnNumberToName(context.LastDateColumnIndex)
	if err != nil {
		return err
	}

	formula := "sum(" + firstDateColumnName + rowIndex + ":" + lastDateColumnName + rowIndex + ")"
	err = f.SetCellFormula(sheet, context.TotalHoursColumn+rowIndex, formula)
	if err != nil {
		return err
	}

	// user total cost
	formula = context.RateColumn + rowIndex + "*" + context.TotalHoursColumn + rowIndex
	err = f.SetCellFormula(sheet, context.TotalCostColumn+rowIndex, formula)
	if err != nil {
		return err
	}

	dateToTimeSpentSeconds := map[string]int{}

	for _, issue := range user.Issues {
		for _, effort := range issue.Efforts {
			date := effort.Date

			if _, ok := dateToTimeSpentSeconds[date]; ok {
				dateToTimeSpentSeconds[date] += effort.TimeSpentSeconds
			} else {
				dateToTimeSpentSeconds[date] = effort.TimeSpentSeconds
			}
		}
	}

	for date, timeSpentSeconds := range dateToTimeSpentSeconds {
		parsedDate, err := time.Parse(effortDateFormat, date)
		if err != nil {
			return err
		}

		formattedDate := fmt.Sprintf(columnDateFormat, parsedDate.Month(), parsedDate.Day())

		for i := context.FirstDateColumnIndex; i <= context.LastDateColumnIndex; i++ {
			col, err := excelize.ColumnNumberToName(i)
			if err != nil {
				return err
			}

			headerDate, err := f.GetCellValue(sheet, col+"1")
			if err != nil {
				return err
			}

			if headerDate == formattedDate {
				err = f.SetCellValue(sheet, col+rowIndex, s.convertSecondsToHours(timeSpentSeconds))
				if err != nil {
					return err
				}
				break
			}
		}
	}

	return nil
}

func (s *ExcelService) convertSecondsToHours(seconds int) float64 {
	value := float64(seconds) / 3600
	return math.Round(value*100) / 100
}
