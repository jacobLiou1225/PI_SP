package pkg

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"
	"strconv"

	"github.com/xuri/excelize/v2"
)

func BuildSpPdf(outputName string) (filePath string) {

	f, err := excelize.OpenFile("bbb.xlsx")
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

	var e19RefrenceTo [5]string
	for i := 0; i < howManyManufacture; i++ {
		e19RefrenceTo[i], _ = excelize.CoordinatesToCellName(32+i, 42)
	}
	var e19 [5]string
	for i := 0; i < howManyManufacture; i++ {
		e19[i], _ = excelize.CoordinatesToCellName(5, 19+i)
	}

	//下的關鍵表格
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

	//doubleArrayProcessing
	var doubleArrayProcessing [5][4]string
	for i := 0; i < 5; i++ {
		for j := 0; j < 4; j++ {
			doubleArrayProcessing[i][j], _ = excelize.CoordinatesToCellName(42+i, 37+j)
		}

	}

	var F37ToF40 [4]string
	for j := 0; j < 4; j++ {
		F37ToF40[j], _ = excelize.CoordinatesToCellName(6, 37+j)
	}

	var AB37ToAB40 [4]string
	for j := 0; j < 4; j++ {
		AB37ToAB40[j], _ = excelize.CoordinatesToCellName(28, 37+j)
	}

	// set doubleArrayProcessing formular
	for j := 0; j < 5; j++ {
		for i := 0; i < 4; i++ {
			f.SetCellFormula("SP", doubleArrayProcessing[j][i], "=IF("+F37ToF40[i]+"="+strconv.Itoa(j+1)+","+AB37ToAB40[i]+",\"0\")")
		}
	}

	//doubleArrayRealManufac
	var doubleArrayRealManufac [5][4]string
	for i := 0; i < 5; i++ {
		for j := 0; j < 4; j++ {
			doubleArrayRealManufac[i][j], _ = excelize.CoordinatesToCellName(47+i, 37+j)
		}

	}

	var Z37ToZ40 [4]string
	for j := 0; j < 4; j++ {
		Z37ToZ40[j], _ = excelize.CoordinatesToCellName(26, 37+j)
	}

	// set doubleArrayRealManufac formular
	for j := 0; j < 5; j++ {
		for i := 0; i < 4; i++ {
			f.SetCellFormula("SP", doubleArrayRealManufac[j][i], "=IF("+E37ToE40[i]+"="+strconv.Itoa(j+1)+","+Z37ToZ40[i]+",\"0\")")
		}
	}

	// E19 數量
	var E19ToE23 [5]string
	for j := 0; j < 5; j++ {
		E19ToE23[j], _ = excelize.CoordinatesToCellName(5, 19+j)
	}
	var AF42ToAJ42 [5]string
	for j := 0; j < 5; j++ {
		AF42ToAJ42[j], _ = excelize.CoordinatesToCellName(32+j, 42)
	}

	for i := 0; i < 5; i++ {
		f.SetCellFormula("SP", E19ToE23[i], "="+AF42ToAJ42[i]+"")
	}
	//E24總和
	f.SetCellFormula("SP", "E24", "=SUM(E19:G23)")

	//H19總價
	var H19ToH23 [5]string
	for j := 0; j < 5; j++ {
		H19ToH23[j], _ = excelize.CoordinatesToCellName(8, 19+j)
	}

	var AK42TOAO42 [5]string
	for j := 0; j < 5; j++ {
		AK42TOAO42[j], _ = excelize.CoordinatesToCellName(37+j, 42)
	}
	var AP42ToAT42 [5]string
	for j := 0; j < 5; j++ {
		AP42ToAT42[j], _ = excelize.CoordinatesToCellName(42+j, 42)
	}

	for i := 0; i < 5; i++ {
		f.SetCellFormula("SP", H19ToH23[i], "="+AK42TOAO42[i]+"+"+AP42ToAT42[i]+"")
	}
	//I24總和
	f.SetCellFormula("SP", "I24", "=SUM(H19:L23)")

	//P19數量
	var P19ToP23 [5]string
	for j := 0; j < 5; j++ {
		P19ToP23[j], _ = excelize.CoordinatesToCellName(16, 19+j)
	}

	for i := 0; i < 5; i++ {
		f.SetCellFormula("SP", P19ToP23[i], "="+E19ToE23[i]+"")
	}
	//P24總和
	f.SetCellFormula("SP", "P24", "=SUM(P19:R23)")

	//S19數量
	var S19ToS23 [5]string
	for j := 0; j < 5; j++ {
		S19ToS23[j], _ = excelize.CoordinatesToCellName(19, 19+j)
	}

	var AU42AY42 [5]string
	for j := 0; j < 5; j++ {
		AU42AY42[j], _ = excelize.CoordinatesToCellName(47+j, 42)
	}
	for i := 0; i < 5; i++ {
		f.SetCellFormula("SP", S19ToS23[i], "="+AU42AY42[i]+"")
	}
	//T24總和
	f.SetCellFormula("SP", "T24", "=SUM(S19:W23)")

	// 出口費用
	var P37TOP40 [4]string
	for j := 0; j < 4; j++ {
		P37TOP40[j], _ = excelize.CoordinatesToCellName(16, 37+j)
	}
	for i := 0; i < 4; i++ {
		f.SetCellFormula("SP", P37TOP40[i], "=$T$52")
	}

	//相關成本費用:
	//進貨成本
	f.SetCellFormula("SP", "H27", "=V49")
	f.SetCellFormula("SP", "H28", "=AB45")
	f.SetCellFormula("SP", "H29", "=AD45")
	//相關成本費用:
	//銷貨成本
	f.SetCellFormula("SP", "S27", "=SUM(M50:M55)")
	f.SetCellFormula("SP", "S28", "=F58")
	f.SetCellFormula("SP", "S29", "=E46")
	f.SetCellFormula("SP", "S30", "=AE45")
	//*每櫃
	f.SetCellFormula("SP", "Q28", "=F55")

	//匯率：USD/MT
	f.SetCellFormula("SP", "X31", "=T55")

	//毛利
	f.SetCellFormula("SP", "Q37", "=G37-K37-M37-N37-O37-P37-L37")
	f.SetCellFormula("SP", "Q38", "=G38-K38-M38-N38-O38-P38-L38")
	f.SetCellFormula("SP", "Q39", "=G39-K39-M39-N39-O39-P39-L39")
	f.SetCellFormula("SP", "Q40", "=G40-K40-M40-N40-O40-P40-L40")

	//毛利總和
	f.SetCellFormula("SP", "T37", "=Q37*R37/1000*$T$55")
	f.SetCellFormula("SP", "T38", "=Q38*R38/1000*$T$55")
	f.SetCellFormula("SP", "T39", "=Q39*R39/1000*$T$55")
	f.SetCellFormula("SP", "T40", "=Q40*R40/1000*$T$55")
	//H30 other fee
	f.SetCellValue("SP", "H30", readPiContent.Body.Sp.OtherFee)

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

	//數量 (MT)計算
	for i := 0; i < 5; i++ {
		f.SetCellFormula("SP", e19[i], "="+e19RefrenceTo[i])

	}

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
