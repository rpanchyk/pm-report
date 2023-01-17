package models

type AppConfig struct {
	Files FilesAppConfig `mapstructure:"files"`
	Tempo TempoAppConfig `mapstructure:"tempo"`
}

type FilesAppConfig struct {
	ProjectConfigFile string `mapstructure:"project_config"`
	ReportFile        string `mapstructure:"report"`
}

type TempoAppConfig struct {
	Url    string                `mapstructure:"url"`
	Tokens []TokenTempoAppConfig `mapstructure:"tokens"`
}

type TokenTempoAppConfig struct {
	Token    string `mapstructure:"token"`
	Projects string `mapstructure:"projects"`
}
