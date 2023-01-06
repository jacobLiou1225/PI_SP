package main

import (
	"PI/pkg"
	"fmt"
)

func main() {
	fileName := "13"
	fmt.Println(pkg.BuildSp(fileName, "pdf"))
	//fmt.Println(pkg.Api(fileName))

}
