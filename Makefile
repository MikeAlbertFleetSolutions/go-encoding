.PHONY: default get codetest fmt lint vet test

default: fmt codetest

get:
	go get -v ./...
	go mod tidy

codetest: lint vet

fmt:
	go fmt ./...

lint:
	$(shell go env GOPATH)/bin/golangci-lint run --fix

vet:
	go vet -all .

test:
	go test -v -cover ./...