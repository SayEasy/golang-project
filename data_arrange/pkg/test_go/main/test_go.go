package main

import (
	"fmt"
	"strings"
)

func main() {
	str := ""
	ss := strings.Split(str,"")
	for i := 0; i < len(ss); i++ {
		fmt.Println(fmt.Sprintf("i = %v, str = %v",i,ss[i]))
	}
}
