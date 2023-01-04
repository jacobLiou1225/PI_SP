package pkg

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"

	"github.com/xuri/excelize/v2"
)

func BuildSpPdf(outputName string) (filePath string) {

	f, err := excelize.OpenFile("ccc.xlsx")
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
	//最開始的基本資料
	f.SetCellValue("SP", "S5", readPiContent.Body.Sp.DeliveryDate)                    //交貨期
	f.SetCellValue("SP", "P7", readPiContent.Body.Sp.PortOfLoading)                   //裝貨港
	f.SetCellValue("SP", "P10", readPiContent.Body.Pi.Customer.Name)                  //客戶名
	f.SetCellValue("SP", "P14", readPiContent.Body.Sp.ManufacturerOrder[0].SalesTerm) //Sales Term   不過這是跟著第一筆資料 ，怪怪的
	f.SetCellValue("SP", "T7", readPiContent.Body.Sp.PortOfDischarge)                 //卸貨港
	f.SetCellValue("SP", "T10", readPiContent.Body.Sp.ContractID)                     //合約號:
	f.SetCellValue("SP", "T14", readPiContent.Body.Sp.PaymentTerm)                    //Payment Term:
	f.SetCellFormula("SP", "E4", "=T24-I24-S27-S28-S29-S30-H29-H30")                  //預估利潤(USD)設定
	f.SetCellFormula("SP", "E5", "=E4/T24")                                           //毛利率設定

	//manuOrderCount :=len(readPiContent.Body.Sp.ManufacturerOrder)
	var howManyManufacture int = len(readPiContent.Body.Sp.ManufacturerOrder)

	var howManySpItem int = len(readPiContent.Body.Sp.SpItems)

	//三角貿易--廠商
	var manufacturerOrderArray [6]string
	for i := 0; i < 5; i++ {
		manufacturerOrderArray[i], _ = excelize.CoordinatesToCellName(3, 8+2*i)
	}
	//三角貿易--sales Term
	var manufacturerSalesTermArray [6]string
	for i := 0; i < 5; i++ {
		manufacturerSalesTermArray[i], _ = excelize.CoordinatesToCellName(3, 9+2*i)
	}
	//台灣進口--廠商
	var TWImportManufac [6]string
	for i := 0; i < 5; i++ {
		TWImportManufac[i], _ = excelize.CoordinatesToCellName(8, 8+2*i)
	}

	//台灣進口--sales Term
	var TWImportSalesTerm [6]string
	for i := 0; i < 5; i++ {
		TWImportSalesTerm[i], _ = excelize.CoordinatesToCellName(8, 9+2*i)
	}
	//廠商名稱/鋼種
	var manufacturerNameAtC19 [5]string
	for i := 0; i < howManyManufacture; i++ {
		manufacturerNameAtC19[i], _ = excelize.CoordinatesToCellName(3, 19+i)
	}

	//下的關鍵表格
	var doubleArrayPiTerms [25][5]string
	for i := 0; i < 23; i++ {
		for j := 0; j < howManySpItem; j++ {
			doubleArrayPiTerms[i][j], _ = excelize.CoordinatesToCellName(1+i, 37+j)
		}

	}

	//H30 other fee
	f.SetCellValue("SP", "H30", readPiContent.Body.Sp.OtherFee)

	for i, _ := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[0][0+i], i+1)
		//在(A,37)紀錄編號
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[1][0+i], n.Grade)
		//在(B,37)紀錄鋼的材質
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[2][0+i], n.Edge)
		//在(c,37)紀錄edge
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[3][0+i], n.Size)
		//在(d,37)紀錄尺寸
	}

	for i, _ := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[4][0+i], i+1)
		//在(E,37)紀錄供應商編號
	}

	for i, _ := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[5][0+i], i+1)
		//在(F,37)紀錄加工廠編號
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[6][0+i], n.UnitPrice)
		//在(G,37)紀錄售價
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[7][0+i], n.Price)
		//在(H,37)紀錄盤價
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[8][0+i], n.ThiPremium)
		//在(i,37)紀錄後寬度加價
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[9][0+i], n.CostOfImport)
		//在(J,37)紀錄進口成本
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[11][0+i], n.FobFee)
		//在(L,37)紀錄FobFee
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[12][0+i], n.Commission)
		//在(M,37)紀錄Commission
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[13][0+i], n.RemainLoss)
		//在(N,37)紀錄RemainLoss
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[17][0+i], n.Quantity)
		//在(R,37)紀錄Quantity
		//HERE'S QUESTION
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[20][0+i], n.Non5Mt)
		//在(U,37)紀錄Non5Mt
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[21][0+i], n.Slinging)
		//在(V,37)紀錄Slinging
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[22][0+i], n.Sticker)
		//在(W,37)紀錄Sticker
	}

	for i, n := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[23][0+i], n.Rpcb)
		//在(X,37)紀錄Rpcb
	}

	for i, n := range readPiContent.Body.Sp.ManufacturerOrder {

		f.SetCellValue("SP", manufacturerOrderArray[i], n.Manufacturer.Name) //在(C,8)放入n.ManuOrderID
	}

	for i, n := range readPiContent.Body.Sp.ManufacturerOrder {

		f.SetCellValue("SP", manufacturerSalesTermArray[i], n.SalesTerm) //在(C,9)放入n.SalesTerm

	}

	for i, n := range readPiContent.Body.Sp.ManufacturerOrder {

		f.SetCellValue("SP", TWImportManufac[i], n.ContractID) //在(H,8)放入n.ManuOrderID

	}
	for i, n := range readPiContent.Body.Sp.ManufacturerOrder {

		f.SetCellValue("SP", TWImportSalesTerm[i+1], n.PaymentTerm) //在(H,9)放入n.SalesTerm

	}

	for i, n := range readPiContent.Body.Sp.ManufacturerOrder {
		f.SetCellValue("SP", manufacturerNameAtC19[i], n.Manufacturer.Name) //在(C,19)放入n.name
	}

	//存檔
	if err := f.SaveAs(outputName + ".xlsx"); err != nil {
		fmt.Println(err)
	}
	return outputName + ".xlsx"
}
