.PHONY: all build build-32 clear

all: build build-32

build:
	GOOS=windows GOARCH=amd64 go build -o fapiao2excel64.exe

build-32:
	GOOS=windows GOARCH=386 go build -o fapiao2excel32.exe

clear:
	for name in $$(ls *.exe); do rm $$name; done