define make_build
    GOOS=$(1) GOARCH=$(2) go build -o builds/$(1)/$(2)/$(3)
	cp -f AppConfig.yaml builds/$(1)/$(2)/
	cd builds/$(1)/$(2) && rm -f pm-report.zip && zip --recurse-paths --move pm-report.zip . && cd -
endef

# Batch build
build: deps build-linux build-macosx build-windows

# Dependencies
deps:
	go mod tidy && go mod vendor

# Linux
build-linux:
	$(call make_build,linux,amd64,pm-report)

# MacOSX
build-macosx:
	$(call make_build,darwin,amd64,pm-report)
	$(call make_build,darwin,arm64,pm-report)

# Windows
build-windows:
	$(call make_build,windows,amd64,pm-report.exe)
