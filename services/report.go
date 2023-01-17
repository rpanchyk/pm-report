package services

import (
	"log"
	"pm-report/models"
	"pm-report/utils"
	"sort"
	"strings"
	"time"
)

type ReportService struct {
	projectConfigService *ProjectConfigService
	tempoService         *TempoService
	tokens               []models.TokenTempoAppConfig
}

func NewReportService(projectConfigService *ProjectConfigService, tempoService *TempoService, tokens []models.TokenTempoAppConfig) *ReportService {
	return &ReportService{
		projectConfigService: projectConfigService,
		tempoService:         tempoService,
		tokens:               tokens,
	}
}

func (s *ReportService) Create(dateFrom, dateTo time.Time) (*models.Report, error) {
	projectConfigWrapper, err := s.projectConfigService.Get()
	if err != nil {
		return nil, err
	}

	var projects []models.Project

	log.Println("Getting report started")

	for _, token := range s.tokens {
		for _, projectKey := range utils.ToList(token.Projects) {
			projectConfig := projectConfigWrapper.Get(projectKey)
			if projectConfig == nil {
				projectConfig = &models.ProjectConfig{Key: projectKey}
			}

			project, err := s.getProject(token.Token, projectConfig, dateFrom, dateTo)
			if err != nil {
				return nil, err
			}
			projects = append(projects, *project)
		}
	}

	log.Println("Getting report finished")

	report := &models.Report{
		DateFrom: dateFrom,
		DateTo:   dateTo,
		Projects: projects,
	}

	err = s.projectConfigService.Save(projectConfigWrapper, report)
	if err != nil {
		return nil, err
	}

	return report, nil
}

func (s *ReportService) getProject(token string, projectConfig *models.ProjectConfig, dateFrom, dateTo time.Time) (*models.Project, error) {
	tempoResults, err := s.tempoService.GetTempoWorklogs(token, projectConfig.Key, dateFrom, dateTo)
	if err != nil {
		return nil, err
	}

	users, err := s.getUsers(tempoResults, projectConfig)
	if err != nil {
		return nil, err
	}

	project := models.Project{
		Key:         projectConfig.Key,
		DisplayName: projectConfig.DisplayName,
		Owner:       projectConfig.Owner,
		Manager:     projectConfig.Manager,
		Users:       users,
	}
	return &project, nil
}

func (s *ReportService) getUsers(results []models.TempoResult, projectConfig *models.ProjectConfig) ([]models.User, error) {
	userIdToTempoResult := map[string][]models.TempoResult{} // group tempo results by account id

	for _, result := range results {
		accountId := result.Author.AccountId
		if _, ok := userIdToTempoResult[accountId]; ok {
			userIdToTempoResult[accountId] = append(userIdToTempoResult[accountId], result)
		} else {
			userIdToTempoResult[accountId] = []models.TempoResult{result}
		}
	}

	var users []models.User

	for _, userResults := range userIdToTempoResult {
		issues, err := s.getIssues(userResults)
		if err != nil {
			return nil, err
		}

		author := userResults[0].Author
		userConfig := projectConfig.UserNameToConfig[author.DisplayName]

		user := models.User{
			AccountId: author.AccountId,
			Name:      author.DisplayName,
			Position:  userConfig.Position,
			Rate:      userConfig.Rate,
			Issues:    issues,
		}
		users = append(users, user)
	}

	sort.Slice(users, func(i, j int) bool {
		return strings.ToLower(users[i].Name) < strings.ToLower(users[j].Name)
	})

	return users, nil
}

func (s *ReportService) getIssues(results []models.TempoResult) ([]models.Issue, error) {
	issueKeyToResults := map[string][]models.TempoResult{} // group tempo results by issue id

	for _, result := range results {
		issueKey := result.Issue.Key

		if _, ok := issueKeyToResults[issueKey]; ok {
			issueKeyToResults[issueKey] = append(issueKeyToResults[issueKey], result)
		} else {
			issueKeyToResults[issueKey] = []models.TempoResult{result}
		}
	}

	var issues []models.Issue

	for issueKey, results := range issueKeyToResults {
		efforts, err := s.getEfforts(results)
		if err != nil {
			return nil, err
		}

		issue := models.Issue{
			Key:     issueKey,
			Efforts: efforts,
		}

		issues = append(issues, issue)
	}

	return issues, nil
}

func (s *ReportService) getEfforts(results []models.TempoResult) ([]models.Effort, error) {
	dateToEffort := map[string]models.Effort{}

	for _, result := range results {
		date := result.StartDate

		var timeSpentSeconds int
		if _, ok := dateToEffort[date]; ok {
			timeSpentSeconds = dateToEffort[date].TimeSpentSeconds + result.TimeSpentSeconds
		} else {
			timeSpentSeconds = result.TimeSpentSeconds
		}

		dateToEffort[date] = models.Effort{
			Date:             result.StartDate,
			TimeSpentSeconds: timeSpentSeconds,
		}
	}

	var efforts []models.Effort

	for _, effort := range dateToEffort {
		efforts = append(efforts, effort)
	}

	return efforts, nil
}
