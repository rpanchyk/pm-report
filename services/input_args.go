package services

import (
	"errors"
	"golang.org/x/text/cases"
	"golang.org/x/text/language"
	"log"
	"os"
	"pm-report/models"
	"pm-report/utils"
	"strconv"
	"strings"
	"time"
)

const (
	monthNumberDateFormat      = "01"
	monthShortStringDateFormat = "Jan"
	monthLongStringDateFormat  = "January"
	yearDateFormat             = "2006"
)

type InputArgsService struct {
}

func NewInputArgsService() *InputArgsService {
	return &InputArgsService{}
}

func (s *InputArgsService) Parse(args []string) (*models.InputArgs, error) {
	if len(args) < 1 {
		return nil, errors.New("error: not enough input arguments")
	}

	// 1st (required)
	monthArg := strings.Trim(args[0], " ")
	monthTime, err := s.tryParseMonthAsNumber(monthArg)
	if err != nil {
		monthTime, err = s.tryParseMonthAsString(monthArg)
		if err != nil {
			return nil, errors.New("error: month as argument is not recognized: " + monthArg)
		}
	}
	log.Println("Month input argument is accepted:", monthTime.Month())

	// 2nd (optional, default: current year)
	yearTime := time.Now()
	if len(args) >= 2 {
		yearArg := strings.Trim(args[1], " ")

		yearTime, err = time.Parse(yearDateFormat, yearArg)
		if err != nil {
			return nil, errors.New("error: year as argument is not recognized: " + yearArg)
		}
		log.Println("Year input argument is accepted:", yearTime.Year())
	} else {
		log.Println("Year input argument is default:", yearTime.Year())
	}

	// 3rd (optional, default: file name)
	appConfig := "AppConfig.yaml"
	if len(args) >= 3 {
		appConfig = args[2]
		_, err := os.Stat(appConfig)
		if err != nil {
			return nil, err
		}
		log.Println("App config file is accepted:", appConfig)
	} else {
		log.Println("App config file is default:", appConfig)
	}

	dateFrom, dateTo, err := s.createDateRange(*monthTime, yearTime)
	if err != nil {
		return nil, err
	}

	inputArgs := &models.InputArgs{
		DateFrom:  *dateFrom,
		DateTo:    *dateTo,
		AppConfig: appConfig,
	}
	log.Println("Parsed", utils.ToPrettyString("input args", inputArgs))

	return inputArgs, nil
}

func (s *InputArgsService) tryParseMonthAsNumber(month string) (*time.Time, error) {
	candidate := month
	if _, err := strconv.ParseInt(month, 10, 64); err == nil {
		if len(month) == 1 {
			candidate = "0" + month
		}
	}

	monthTime, err := time.Parse(monthNumberDateFormat, candidate)
	if err != nil {
		return nil, err
	}

	return &monthTime, nil
}

func (s *InputArgsService) tryParseMonthAsString(month string) (*time.Time, error) {
	candidate := cases.Title(language.English).String(month)

	monthTime, err := time.Parse(monthShortStringDateFormat, candidate)
	if err != nil {
		monthTime, err = time.Parse(monthLongStringDateFormat, candidate)
		if err != nil {
			return nil, err
		}
	}

	return &monthTime, nil
}

func (s *InputArgsService) createDateRange(monthTime, yearTime time.Time) (*time.Time, *time.Time, error) {
	candidate := strconv.Itoa(yearTime.Year()) + "-" + monthTime.Month().String()

	dateFrom, err := time.Parse("2006-January", candidate)
	if err != nil {
		return nil, nil, err
	}

	dateTo := dateFrom.AddDate(0, 1, -1)

	return &dateFrom, &dateTo, nil
}
