package services

import (
	"encoding/json"
	"errors"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"pm-report/models"
	"time"
)

const (
	dateFormat = "2006-01-02"
)

type TempoService struct {
	worklogsUrlTemplate string
}

func NewTempoService(url string) *TempoService {
	return &TempoService{
		worklogsUrlTemplate: url + "/core/3/worklogs?project=%s&from=%s&to=%s&offset=%d&limit=%d",
	}
}

func (s *TempoService) GetTempoWorklogs(token, projectKey string, dateFrom, dateTo time.Time) ([]models.TempoResult, error) {
	var tempoResults []models.TempoResult
	offset := 0
	limit := 100

	for {
		response, err := s.fetchTempoResponse(token, projectKey, dateFrom, dateTo, offset, limit)
		if err != nil {
			return nil, err
		}
		tempoResults = append(tempoResults, response.Results...)

		log.Println("Fetched tempo report for", projectKey, "project:", response.Metadata.Count, "records")

		if response.Metadata.Count < response.Metadata.Limit {
			break
		}
		offset = response.Metadata.Count + response.Metadata.Offset
	}

	return tempoResults, nil
}

func (s *TempoService) fetchTempoResponse(token, projectKey string, dateFrom, dateTo time.Time, offset, limit int) (*models.TempoResponse, error) {
	client := http.Client{Timeout: time.Second * 60}

	url := fmt.Sprintf(s.worklogsUrlTemplate,
		projectKey,
		dateFrom.Format(dateFormat),
		dateTo.Format(dateFormat),
		offset,
		limit)

	request, err := http.NewRequest(http.MethodGet, url, nil)
	if err != nil {
		return nil, err
	}

	request.Header.Set("Authorization", "Bearer "+token)
	request.Header.Set("Content-Type", "application/json")

	response, err := client.Do(request)
	if err != nil {
		return nil, err
	}

	if response.Body != nil {
		defer func() {
			err := response.Body.Close()
			if err != nil {
				log.Println(err)
				return
			}
		}()
	}

	if response.StatusCode != 200 {
		return nil, errors.New("Tempo error: " + response.Status)
	}

	body, err := ioutil.ReadAll(response.Body)
	if err != nil {
		return nil, err
	}

	tempoResponse := &models.TempoResponse{}
	err = json.Unmarshal(body, tempoResponse)
	if err != nil {
		return nil, err
	}

	return tempoResponse, nil
}
