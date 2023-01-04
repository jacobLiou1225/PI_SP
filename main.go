package main

import (
	"PI/pkg"
	"fmt"
)

func main() {
	fileName := "225"
	fmt.Println(pkg.BuildSp(fileName, "pdf"))
	//fmt.Println(pkg.Api(fileName))

}
