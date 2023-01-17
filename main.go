package main

import (
	"log"
	"os"
	"pm-report/services"
)

func main() {
	log.Println("Report creating started")

	// args
	inputArgsService := services.NewInputArgsService()

	inputArgs, err := inputArgsService.Parse(os.Args[1:])
	if err != nil {
		log.Fatal(err)
		return
	}

	// config
	appConfigService := services.NewAppConfigService(inputArgs.AppConfig)

	appConfig, err := appConfigService.Get()
	if err != nil {
		log.Fatal(err)
		return
	}

	// get data
	reportService := services.NewReportService(
		services.NewProjectConfigService(appConfig.Files.ProjectConfigFile),
		services.NewTempoService(appConfig.Tempo.Url),
		appConfig.Tempo.Tokens)

	report, err := reportService.Create(inputArgs.DateFrom, inputArgs.DateTo)
	if err != nil {
		log.Fatal(err)
		return
	}

	// save data
	excelService := services.NewExcelService(appConfig.Files.ReportFile)

	err = excelService.Save(report)
	if err != nil {
		log.Fatal(err)
		return
	}

	log.Println("Report creating finished successfully")
}
