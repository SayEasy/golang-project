

# main.go相同目录运行go build 生成main.exe   windows
go build

#linux 运行一下命令，生成data-arrange-cli （和main.exe运行包）
CGO_ENABLED=1 GOOS=linux GOARCH=amd64 go build -tags static_all -v -o  ./bin/data-arrange-cli  ./pkg/main/main.go