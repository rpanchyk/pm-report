package services

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"pm-report/models"
	"pm-report/utils"
	"sort"
	"strconv"
)

type ProjectConfigService struct {
	filePath string
}

func NewProjectConfigService(filePath string) *ProjectConfigService {
	return &ProjectConfigService{
		filePath: filePath,
	}
}

func (s *ProjectConfigService) Get() (*models.ProjectConfigWrapper, error) {
	_, err := os.Stat(s.filePath)
	if err != nil && errors.Is(err, os.ErrNotExist) {
		return &models.ProjectConfigWrapper{}, nil
	}

	f, err := excelize.OpenFile(s.filePath)
	if err != nil {
		return nil, err
	}

	context := s.createContext()
	var projectConfigs []models.ProjectConfig

	for _, sheet := range f.GetSheetList() {
		projectConfig, err := s.getProjectConfig(f, sheet, context)
		if err != nil {
			return nil, err
		}
		projectConfigs = append(projectConfigs, *projectConfig)
	}

	err = f.Close()
	if err != nil {
		return nil, err
	}

	projectConfigWrapper := &models.ProjectConfigWrapper{ProjectConfigs: projectConfigs}

	log.Println("Parsed", s.filePath, utils.ToPrettyString("config", projectConfigWrapper))

	return projectConfigWrapper, nil
}

func (s *ProjectConfigService) getProjectConfig(f *excelize.File, sheet string, context *models.ProjectConfigContext) (*models.ProjectConfig, error) {
	displayName, err := f.GetCellValue(sheet, context.Project.DisplayNameValueCell)
	if err != nil {
		return nil, err
	}

	owner, err := f.GetCellValue(sheet, context.Project.OwnerValueCell)
	if err != nil {
		return nil, err
	}

	manager, err := f.GetCellValue(sheet, context.Project.ManagerValueCell)
	if err != nil {
		return nil, err
	}

	userConfigs, err := s.getUserConfigs(f, sheet, context)
	if err != nil {
		return nil, err
	}

	projectConfig := models.ProjectConfig{
		Key:              sheet,
		DisplayName:      displayName,
		Owner:            owner,
		Manager:          manager,
		UserNameToConfig: userConfigs,
	}
	return &projectConfig, nil
}

func (s *ProjectConfigService) getUserConfigs(f *excelize.File, sheet string, context *models.ProjectConfigContext) (map[string]models.UserConfig, error) {
	userConfigs := map[string]models.UserConfig{}

	rows, err := f.Rows(sheet)
	if err != nil {
		return nil, err
	}

	rowIndex := 0
	for rows.Next() {
		rowIndex++
		if rowIndex <= context.User.HeaderRowIndex {
			continue // skip
		}

		rowCols, err := rows.Columns()
		if err != nil {
			return nil, err
		}

		if rowCols == nil {
			continue
		}

		userName := ""
		if len(rowCols) > 0 {
			userName = rowCols[0]
		}
		if len(userName) == 0 { // required
			continue
		}

		position := ""
		if len(rowCols) > 1 {
			position = rowCols[1]
		}

		rate := 0
		if len(rowCols) > 2 {
			rate, err = strconv.Atoi(rowCols[2])
			if err != nil {
				return nil, err
			}
		}

		userConfigs[userName] = models.UserConfig{
			Position: position,
			Rate:     rate,
		}
	}

	if err = rows.Close(); err != nil {
		return nil, err
	}

	return userConfigs, nil
}

func (s *ProjectConfigService) Save(projectConfigWrapper *models.ProjectConfigWrapper, report *models.Report) error {
	updatedProjectConfigWrapper := s.updateProjectConfigs(projectConfigWrapper, report)

	err := s.saveProjectConfigs(updatedProjectConfigWrapper)
	if err != nil {
		return err
	}

	log.Println("Synchronized", s.filePath, utils.ToPrettyString("config", updatedProjectConfigWrapper))

	return nil
}

func (s *ProjectConfigService) updateProjectConfigs(projectConfigWrapper *models.ProjectConfigWrapper, report *models.Report) *models.ProjectConfigWrapper {
	var projectConfigs []models.ProjectConfig

	for _, reportProject := range report.Projects {
		updatedUserConfigs := map[string]models.UserConfig{}

		// actual users from report
		for _, reportUser := range reportProject.Users {
			updatedUserConfigs[reportUser.Name] = models.UserConfig{
				Position: reportUser.Position,
				Rate:     reportUser.Rate,
			}
		}

		// missing users from config
		projectConfig := projectConfigWrapper.Get(reportProject.Key)
		if projectConfig != nil {
			for userName, userConfig := range projectConfig.UserNameToConfig {
				if _, ok := updatedUserConfigs[userName]; !ok {
					updatedUserConfigs[userName] = models.UserConfig{
						Position: userConfig.Position,
						Rate:     userConfig.Rate,
					}
				}
			}
		}

		updatedProjectConfig := models.ProjectConfig{
			Key:              reportProject.Key,
			DisplayName:      reportProject.DisplayName,
			Owner:            reportProject.Owner,
			Manager:          reportProject.Manager,
			UserNameToConfig: updatedUserConfigs,
		}
		projectConfigs = append(projectConfigs, updatedProjectConfig)
	}

	return &models.ProjectConfigWrapper{ProjectConfigs: projectConfigs}
}

func (s *ProjectConfigService) saveProjectConfigs(projectConfigWrapper *models.ProjectConfigWrapper) error {
	f := excelize.NewFile()
	context := s.createContext()

	for _, projectConfig := range projectConfigWrapper.ProjectConfigs {
		s.createSheet(f, projectConfig.Key)

		err := s.fillProjectInfo(f, projectConfig.Key, context, &projectConfig)
		if err != nil {
			return err
		}

		err = s.fillUsersHeader(f, projectConfig.Key, context, &projectConfig)
		if err != nil {
			return err
		}

		err = s.fillUsersBody(f, projectConfig.Key, context, &projectConfig)
		if err != nil {
			return err
		}
	}

	err := f.SaveAs(s.filePath)
	if err != nil {
		return err
	}

	err = f.Close()
	if err != nil {
		return err
	}

	return nil
}

func (s *ProjectConfigService) createSheet(f *excelize.File, sheet string) {
	f.NewSheet(sheet)

	if f.GetSheetIndex("Sheet1") != -1 {
		f.DeleteSheet("Sheet1")
	}
}

func (s *ProjectConfigService) fillProjectInfo(f *excelize.File, sheet string, context *models.ProjectConfigContext, projectConfig *models.ProjectConfig) error {
	err := f.MergeCell(sheet, "A1", "C1")
	if err != nil {
		return err
	}
	alignment := excelize.Alignment{Horizontal: "center"}
	borders := []excelize.Border{
		{Type: "top", Color: "#000000", Style: 1},
		{Type: "left", Color: "#000000", Style: 1},
		{Type: "bottom", Color: "#000000", Style: 1},
		{Type: "right", Color: "#000000", Style: 1},
	}
	headerFont := excelize.Font{Color: "#ffffff", Bold: true}
	fill := excelize.Fill{Color: []string{"#009a00"}, Type: "pattern", Pattern: 1}
	style, err := f.NewStyle(&excelize.Style{Alignment: &alignment, Font: &headerFont, Border: borders, Fill: fill})
	if err != nil {
		return err
	}
	err = f.SetCellStyle(sheet, context.Project.HeaderCell, context.Project.HeaderCell, style)
	if err != nil {
		return err
	}
	err = f.SetCellValue(sheet, context.Project.HeaderCell, "Project Info")
	if err != nil {
		return err
	}

	col, row, err := excelize.CellNameToCoordinates(context.Project.KeyValueCell)
	if err != nil {
		return err
	}
	cell, err := excelize.CoordinatesToCellName(col+1, row)
	if err != nil {
		return err
	}
	err = f.MergeCell(sheet, context.Project.KeyValueCell, cell)
	if err != nil {
		return err
	}
	err = f.SetCellValue(sheet, context.Project.KeyTitleCell, "Key")
	if err != nil {
		return err
	}
	err = f.SetCellValue(sheet, context.Project.KeyValueCell, sheet)
	if err != nil {
		return err
	}

	col, row, err = excelize.CellNameToCoordinates(context.Project.DisplayNameValueCell)
	if err != nil {
		return err
	}
	cell, err = excelize.CoordinatesToCellName(col+1, row)
	if err != nil {
		return err
	}
	err = f.MergeCell(sheet, context.Project.DisplayNameValueCell, cell)
	if err != nil {
		return err
	}
	err = f.SetCellValue(sheet, context.Project.DisplayNameTitleCell, "Display Name")
	if err != nil {
		return err
	}
	err = f.SetCellValue(sheet, context.Project.DisplayNameValueCell, projectConfig.DisplayName)
	if err != nil {
		return err
	}

	col, row, err = excelize.CellNameToCoordinates(context.Project.OwnerValueCell)
	if err != nil {
		return err
	}
	cell, err = excelize.CoordinatesToCellName(col+1, row)
	if err != nil {
		return err
	}
	err = f.MergeCell(sheet, context.Project.OwnerValueCell, cell)
	if err != nil {
		return err
	}
	err = f.SetCellValue(sheet, context.Project.OwnerTitleCell, "Owner")
	if err != nil {
		return err
	}
	err = f.SetCellValue(sheet, context.Project.OwnerValueCell, projectConfig.Owner)
	if err != nil {
		return err
	}

	col, row, err = excelize.CellNameToCoordinates(context.Project.ManagerValueCell)
	if err != nil {
		return err
	}
	cell, err = excelize.CoordinatesToCellName(col+1, row)
	if err != nil {
		return err
	}
	err = f.MergeCell(sheet, context.Project.ManagerValueCell, cell)
	if err != nil {
		return err
	}
	err = f.SetCellValue(sheet, context.Project.ManagerTitleCell, "Manager")
	if err != nil {
		return err
	}
	err = f.SetCellValue(sheet, context.Project.ManagerValueCell, projectConfig.Manager)
	if err != nil {
		return err
	}

	return nil
}

func (s *ProjectConfigService) fillUsersHeader(f *excelize.File, sheet string, context *models.ProjectConfigContext, projectConfig *models.ProjectConfig) error {
	err := f.SetColWidth(sheet, context.User.NameColumn, context.User.NameColumn, 30)
	if err != nil {
		return err
	}

	err = f.SetColWidth(sheet, context.User.PositionColumn, context.User.PositionColumn, 30)
	if err != nil {
		return err
	}

	err = f.SetColWidth(sheet, context.User.RateColumn, context.User.RateColumn, 10)
	if err != nil {
		return err
	}

	style, err := f.NewStyle(&excelize.Style{NumFmt: 177})
	if err != nil {
		return err
	}
	err = f.SetColStyle(sheet, context.User.RateColumn, style)
	if err != nil {
		return err
	}

	rowIndex := strconv.Itoa(context.User.HeaderRowIndex - 1)

	err = f.MergeCell(sheet, context.User.NameColumn+rowIndex, context.User.RateColumn+rowIndex)
	if err != nil {
		return err
	}

	rowIndex = strconv.Itoa(context.User.HeaderRowIndex)

	alignment := excelize.Alignment{Horizontal: "center"}
	borders := []excelize.Border{
		{Type: "top", Color: "#000000", Style: 1},
		{Type: "left", Color: "#000000", Style: 1},
		{Type: "bottom", Color: "#000000", Style: 1},
		{Type: "right", Color: "#000000", Style: 1},
	}
	headerFont := excelize.Font{Color: "#ffffff", Bold: true}
	fill := excelize.Fill{Color: []string{"#009a00"}, Type: "pattern", Pattern: 1}
	style, err = f.NewStyle(&excelize.Style{Alignment: &alignment, Font: &headerFont, Border: borders, Fill: fill})
	if err != nil {
		return err
	}
	err = f.SetCellStyle(sheet, context.User.NameColumn+rowIndex, context.User.RateColumn+rowIndex, style)
	if err != nil {
		return err
	}

	err = f.SetCellValue(sheet, context.User.NameColumn+rowIndex, "Name")
	if err != nil {
		return err
	}

	err = f.SetCellValue(sheet, context.User.PositionColumn+rowIndex, "Position")
	if err != nil {
		return err
	}

	err = f.SetCellValue(sheet, context.User.RateColumn+rowIndex, "Rate")
	if err != nil {
		return err
	}

	return nil
}

func (s *ProjectConfigService) fillUsersBody(f *excelize.File, sheet string, context *models.ProjectConfigContext, projectConfig *models.ProjectConfig) error {
	lastRowIndex := context.User.HeaderRowIndex
	users := projectConfig.UserNameToConfig

	userNames := make([]string, 0, len(users))
	for userName := range users {
		userNames = append(userNames, userName)
	}
	sort.Strings(userNames)

	conditionalFormat, err := s.getConditionalFormat(f)
	if err != nil {
		return err
	}

	for _, userName := range userNames {
		lastRowIndex++
		rowIndex := strconv.Itoa(lastRowIndex)

		err = f.SetConditionalFormat(sheet, context.User.RateColumn+rowIndex, *conditionalFormat)
		if err != nil {
			return err
		}

		err = f.SetCellValue(sheet, context.User.NameColumn+rowIndex, userName)
		if err != nil {
			return err
		}

		err = f.SetCellValue(sheet, context.User.PositionColumn+rowIndex, users[userName].Position)
		if err != nil {
			return err
		}

		err = f.SetCellValue(sheet, context.User.RateColumn+rowIndex, users[userName].Rate)
		if err != nil {
			return err
		}
	}

	return nil
}

func (s *ProjectConfigService) createContext() *models.ProjectConfigContext {
	return &models.ProjectConfigContext{
		Project: models.InfoProjectConfigContext{
			HeaderCell: "A1",

			KeyTitleCell: "A2",
			KeyValueCell: "B2",

			DisplayNameTitleCell: "A3",
			DisplayNameValueCell: "B3",

			OwnerTitleCell: "A4",
			OwnerValueCell: "B4",

			ManagerTitleCell: "A5",
			ManagerValueCell: "B5",
		},
		User: models.UserProjectConfigContext{
			HeaderRowIndex: 7,
			NameColumn:     "A",
			PositionColumn: "B",
			RateColumn:     "C",
		},
	}
}

func (s *ProjectConfigService) getConditionalFormat(f *excelize.File) (*string, error) {
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

	formatSet := fmt.Sprintf(`[{ "type": "cell", "criteria": "=", "format": %d, "value": "0" }]`, format)
	return &formatSet, nil
}
