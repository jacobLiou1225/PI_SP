package pkg

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"
	"strconv"

	"github.com/xuri/excelize/v2"
)

func BuildSp(outputName string) (filePath string) {

	f, err := excelize.OpenFile("spModle.xlsx")
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

	//manuOrderCount :=len(readPiContent.Body.Sp.ManufacturerOrder)
	var howManyManufacture int = len(readPiContent.Body.Sp.ManufacturerOrder)

	var howManySpItem int = len(readPiContent.Body.Sp.SpItems)

	//三角貿易
	var manufacturerOrderArray [10]string
	for i := 0; i < howManyManufacture; i++ {
		manufacturerOrderArray[i], _ = excelize.CoordinatesToCellName(3, 8+i)
	}
	//台灣進口
	var manufacturerContractArray [10]string
	for i := 0; i < howManyManufacture; i++ {
		manufacturerContractArray[i], _ = excelize.CoordinatesToCellName(8, 8+i)
	}
	//廠商名稱/鋼種
	var manufacturerNameAtC19 [5]string
	for i := 0; i < howManyManufacture; i++ {
		manufacturerNameAtC19[i], _ = excelize.CoordinatesToCellName(3, 19+i)
	}

	//數量 (MT)
	var manufacturerAmountAtE19 [5]string
	for i := 0; i < howManyManufacture; i++ {
		manufacturerAmountAtE19[i], _ = excelize.CoordinatesToCellName(5, 19+i)
	}
	//總價 (USD)
	var manufacturerPriceAtH19 [5]string
	for i := 0; i < howManyManufacture; i++ {
		manufacturerPriceAtH19[i], _ = excelize.CoordinatesToCellName(8, 19+i)
	}

	var e19RefrenceTo [5]string
	for i := 0; i < howManyManufacture; i++ {
		e19RefrenceTo[i], _ = excelize.CoordinatesToCellName(32+i, 42)
	}
	var e19 [5]string
	for i := 0; i < howManyManufacture; i++ {
		e19[i], _ = excelize.CoordinatesToCellName(5, 19+i)
	}

	//中下的關鍵表格
	var doubleArrayPiTerms [25][6]string
	for i := 0; i < 23; i++ {
		for j := 0; j < howManySpItem; j++ {
			doubleArrayPiTerms[i][j], _ = excelize.CoordinatesToCellName(1+i, 37+j)
		}

	}

	//doubleArrayManufac
	var doubleArrayManufac [5][4]string
	for i := 0; i < 5; i++ {
		for j := 0; j < 4; j++ {
			doubleArrayManufac[i][j], _ = excelize.CoordinatesToCellName(32+i, 37+j)
		}

	}
	var E37ToE40 [4]string
	for j := 0; j < 4; j++ {
		E37ToE40[j], _ = excelize.CoordinatesToCellName(5, 37+j)
	}
	var R37ToR40 [4]string
	for j := 0; j < 4; j++ {
		R37ToR40[j], _ = excelize.CoordinatesToCellName(18, 37+j)
	}

	// set DoubleArrayManufac formular
	for j := 0; j < 5; j++ {
		for i := 0; i < 4; i++ {
			f.SetCellFormula("SP", doubleArrayManufac[j][i], "=IF("+E37ToE40[i]+"="+strconv.Itoa(j+1)+","+R37ToR40[i]+"/1000,\"0\")")
		}
	}

	//doubleArraySupplier
	var doubleArraySupplier [5][4]string
	for i := 0; i < 5; i++ {
		for j := 0; j < 4; j++ {
			doubleArraySupplier[i][j], _ = excelize.CoordinatesToCellName(37+i, 37+j)
		}

	}

	var AA37ToAA40 [4]string
	for j := 0; j < 4; j++ {
		AA37ToAA40[j], _ = excelize.CoordinatesToCellName(27, 37+j)
	}

	// set doubleArraySupplier formular
	for j := 0; j < 5; j++ {
		for i := 0; i < 4; i++ {
			f.SetCellFormula("SP", doubleArraySupplier[j][i], "=IF("+E37ToE40[i]+"="+strconv.Itoa(j+1)+","+AA37ToAA40[i]+",\"0\")")
		}
	}

	//最開始的基本資料

	f.SetCellValue("SP", "P5", readPiContent.Body.Sp.DeliveryDate)                    //交貨期
	f.SetCellValue("SP", "P7", readPiContent.Body.Sp.PortOfLoading)                   //裝貨港
	f.SetCellValue("SP", "P10", readPiContent.Body.Pi.Customer.Name)                  //客戶名
	f.SetCellValue("SP", "P14", readPiContent.Body.Sp.ManufacturerOrder[0].SalesTerm) //Sales Term   不過這是跟著第一筆資料 ，怪怪的
	f.SetCellValue("SP", "T7", readPiContent.Body.Sp.PortOfDischarge)                 //卸貨港
	f.SetCellValue("SP", "T10", readPiContent.Body.Sp.ContractID)                     //合約號:
	f.SetCellValue("SP", "T14", readPiContent.Body.Sp.PaymentTerm)                    //Payment Term:
	f.SetCellFormula("SP", "E4", "=T24-I24-S27-S28-S29-S30-H29-H30")                  //公視設定
	f.SetCellFormula("SP", "E5", "=E4/T24")                                           //公視設定

	//數量 (MT)計算
	for i := 0; i < 5; i++ {
		f.SetCellFormula("SP", e19[i], "="+e19RefrenceTo[i])

	}
	//銷貨收入
	f.SetCellFormula("SP", "Z37", "=ROUND(R37*G37/1000,2)")
	f.SetCellFormula("SP", "Z38", "=ROUND(R38*G38/1000,2)")
	f.SetCellFormula("SP", "Z39", "=ROUND(R39*G39/1000,2)")
	f.SetCellFormula("SP", "Z40", "=ROUND(R40*G40/1000,2)")
	//出貨成本
	f.SetCellFormula("SP", "AA37", "=ROUND((H37+I37)*R37/1000,0)")
	f.SetCellFormula("SP", "AA38", "=ROUND((H38+I38)*R38/1000,0)")
	f.SetCellFormula("SP", "AA39", "=ROUND((H39+I39)*R39/1000,0)")
	f.SetCellFormula("SP", "AA40", "=ROUND((H40+I40)*R40/1000,0)")

	//加工成本
	f.SetCellFormula("SP", "AB37", "=ROUND(O37*R37/1000,0)")
	f.SetCellFormula("SP", "AB38", "=ROUND(O38*R38/1000,0)")
	f.SetCellFormula("SP", "AB39", "=ROUND(O39*R39/1000,0)")
	f.SetCellFormula("SP", "AB40", "=ROUND(O40*R40/1000,0)")
	//銷售成本
	f.SetCellFormula("SP", "AC37", "=ROUND((H37+I37+O37)*R37/1000,0)")
	f.SetCellFormula("SP", "AC38", "=ROUND((H38+I38+O38)*R38/1000,0)")
	f.SetCellFormula("SP", "AC39", "=ROUND((H39+I39+O39)*R39/1000,0)")
	f.SetCellFormula("SP", "AC40", "=ROUND((H40+I40+O40)*R40/1000,0)")
	//餘料損失
	f.SetCellFormula("SP", "AD37", "=N37*R37/1000")
	f.SetCellFormula("SP", "AD38", "=N38*R38/1000")
	f.SetCellFormula("SP", "AD39", "=N39*R39/1000")
	f.SetCellFormula("SP", "AD40", "=N40*R40/1000")
	//出口報關
	f.SetCellFormula("SP", "AE37", "=(J37+L37)*R37/1000")
	f.SetCellFormula("SP", "AE38", "=(J38+L38)*R38/1000")
	f.SetCellFormula("SP", "AE39", "=(J39+L39)*R39/1000")
	f.SetCellFormula("SP", "AE40", "=(J40+L40)*R40/1000")

	for i, _ := range readPiContent.Body.Sp.SpItems {
		f.SetCellValue("SP", doubleArrayPiTerms[0][0+i], i+0)
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
	if err := f.SaveAs(outputName + ".xlsx"); err != nil {
		fmt.Println(err)
	}
	return outputName + ".xlsx"
}
