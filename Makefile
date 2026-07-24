.PHONY: audit build check coverage eval format format-check install release-check test

install:
	go mod download

format:
	gofmt -w cmd internal evals

format-check:
	test -z "$$(gofmt -l cmd internal evals)"

test:
	go test ./...

coverage:
	go test -coverprofile=coverage.out ./cmd/... ./internal/commands ./internal/format
	bash scripts/check-go-coverage.sh coverage.out 80

# Evaluate the binary from this checkout, not whatever xlsx-tool happens to be
# on PATH.
eval: build
	XLSX_TOOL_BIN=./xlsx-tool go run ./evals

build:
	go build ./cmd/xlsx-tool

audit:
	go mod verify
	GOTOOLCHAIN=$$(go env GOVERSION) go run golang.org/x/vuln/cmd/govulncheck@v1.1.4 ./...

check: format-check
	go vet ./...
	go test -race ./...
	$(MAKE) coverage
	$(MAKE) eval

release-check: check
	printf '%s\n' "$${GITHUB_REF_NAME}" | grep -Eq '^v[0-9]+\.[0-9]+\.[0-9]+([+-][0-9A-Za-z.-]+)?$$'
