package pkg

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"

	"github.com/xuri/excelize/v2"
)

func BuildSp() {

	f, err := excelize.OpenFile("piModle.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	res, err := http.Get("https://api.testing.eirc.app/meglobe/v1.0/order/pisp/b4e71c02-ed05-4a7c-bdfe-132b1d36800f")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer res.Body.Close()

	body, err := ioutil.ReadAll(res.Body)
	if err != nil {
		fmt.Println(err)
		return
	}
	var readPiContent Read_Pi_content
	json.Unmarshal(body, &readPiContent)
	//SP部分------------------------------------------分隔線---------------------------------------------
	//manuOrderCount :=len(readPiContent.Body.Sp.ManufacturerOrder)
	var howManyManufacture int = len(readPiContent.Body.Sp.ManufacturerOrder)

	var howManySpItem int = len(readPiContent.Body.Sp.SpItems)

	var theArrayMoreThan5 int = len(readPiContent.Body.Sp.SpItems) - 5

	var manufacturerOrderArray [10]string
	for i := 0; i < howManyManufacture; i++ {
		manufacturerOrderArray[i], _ = excelize.CoordinatesToCellName(3, 8+i)
	}

	var manufacturerContractArray [10]string
	for i := 0; i < howManyManufacture; i++ {
		manufacturerContractArray[i], _ = excelize.CoordinatesToCellName(8, 8+i)
	}
	var manufacturerNameAtC19 [10]string
	for i := 0; i < howManyManufacture; i++ {
		manufacturerNameAtC19[i], _ = excelize.CoordinatesToCellName(3, 19+i)
	}

	var doubleArray [25][10]string
	for i := 0; i < 23; i++ {
		for j := 0; j < howManySpItem; j++ {
			doubleArray[i][j], _ = excelize.CoordinatesToCellName(1+i, 37+j)
		}

	}

	for i, _ := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[0][0+i], i+0)
		//在(A,37)紀錄編號
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[1][0+i], n.Grade)
		//在(B,37)紀錄鋼的材質
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[2][0+i], n.Edge)
		//在(c,37)紀錄edge
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[3][0+i], n.Size)
		//在(d,37)紀錄尺寸
	}

	for i, _ := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[4][0+i], i+1)
		//在(E,37)紀錄供應商編號
	}

	for i, _ := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[5][0+i], i+1)
		//在(F,37)紀錄加工廠編號
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[6][0+i], n.UnitPrice)
		//在(G,37)紀錄售價
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[7][0+i], n.Price)
		//在(H,37)紀錄盤價
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[8][0+i], n.ThiPremium)
		//在(i,37)紀錄後寬度加價
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[9][0+i], n.CostOfImport)
		//在(J,37)紀錄進口成本
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[11][0+i], n.FobFee)
		//在(L,37)紀錄FobFee
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[12][0+i], n.Commission)
		//在(M,37)紀錄Commission
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[13][0+i], n.RemainLoss)
		//在(N,37)紀錄RemainLoss
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[17][0+i], n.Quantity)
		//在(R,37)紀錄Quantity
		//HERE'S QUESTION
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[20][0+i], n.Non5Mt)
		//在(U,37)紀錄Non5Mt
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[21][0+i], n.Slinging)
		//在(V,37)紀錄Slinging
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[22][0+i], n.Sticker)
		//在(W,37)紀錄Sticker
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArray[23][0+i], n.Rpcb)
		//在(X,37)紀錄Rpcb
	}

	for i, n := range readPiContent.Body.Sp.ManufacturerOrder {

		if howManyManufacture > 5 {
			for j := 0; j < theArrayMoreThan5; j++ {
				f.InsertRow("SP", 17+i)
			}
		}

		f.SetCellValue("SP", manufacturerOrderArray[i], n.Manufacturer.Name) //在(C,10)放入n.ManuOrderID

		f.SetCellValue("SP", manufacturerOrderArray[i+1], n.SalesTerm) //在(C,11)放入n.SalesTerm
		i++

	}

	for i, n := range readPiContent.Body.Sp.ManufacturerOrder {

		f.SetCellValue("SP", manufacturerContractArray[i], n.ContractID)    //在(H,10)放入n.ManuOrderID
		f.SetCellValue("SP", manufacturerContractArray[i+1], n.PaymentTerm) //在(H,11)放入n.SalesTerm
		i++
	}

	for i, n := range readPiContent.Body.Sp.ManufacturerOrder {
		f.SetCellValue("SP", manufacturerNameAtC19[i], n.Manufacturer.Name) //在(C,19)放入n.name
	}

	//存檔
	if err := f.SaveAs("piForHo222222企業.xlsx"); err != nil {
		fmt.Println(err)
	}
}
