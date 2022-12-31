package main

import (
	"PI/pkg"
	"fmt"
)

func main() {
	fileName := "225"
	fmt.Println(pkg.BuildSpPdf(fileName))
	fmt.Println(pkg.Api(fileName))

}
