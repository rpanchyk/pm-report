package models

import "time"

type Report struct {
	DateFrom time.Time
	DateTo   time.Time
	Projects []Project
}

type Project struct {
	Key         string
	DisplayName string
	Owner       string
	Manager     string
	Users       []User
}

type User struct {
	AccountId string
	Name      string
	Position  string
	Rate      int
	Issues    []Issue
}

type Issue struct {
	Key     string
	Efforts []Effort
}

type Effort struct {
	Date             string
	TimeSpentSeconds int
}
