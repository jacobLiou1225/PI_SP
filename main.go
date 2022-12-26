package main

import (
	"PI/pkg"
	"fmt"
)

func main() {
	fileName := "請在這裡輸入檔案名稱"
	fmt.Println(pkg.BuildPi(fileName))
	fmt.Println(pkg.Api(fileName))
	//pkg.BuildSp()
}
