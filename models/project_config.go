package models

type ProjectConfigWrapper struct {
	ProjectConfigs []ProjectConfig
}

type ProjectConfig struct {
	Key              string
	DisplayName      string
	Owner            string
	Manager          string
	UserNameToConfig map[string]UserConfig
}

type UserConfig struct {
	Position string
	Rate     int
}

func (s *ProjectConfigWrapper) Get(projectKey string) *ProjectConfig {
	for _, config := range s.ProjectConfigs {
		if config.Key == projectKey {
			return &config
		}
	}
	return nil
}

type ProjectConfigContext struct {
	Project      InfoProjectConfigContext
	User         UserProjectConfigContext
	LastRowIndex int
}

type InfoProjectConfigContext struct {
	HeaderCell string

	KeyTitleCell string
	KeyValueCell string

	DisplayNameTitleCell string
	DisplayNameValueCell string

	OwnerTitleCell string
	OwnerValueCell string

	ManagerTitleCell string
	ManagerValueCell string
}

type UserProjectConfigContext struct {
	HeaderRowIndex int

	NameColumn     string
	PositionColumn string
	RateColumn     string
}
