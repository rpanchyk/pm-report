package models

import "time"

type InputArgs struct {
	DateFrom  time.Time
	DateTo    time.Time
	AppConfig string
}
