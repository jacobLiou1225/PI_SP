package build_file

import (
	"fmt"
	"math"
	"strconv"

	orderModel "eirc.app/internal/v1/structure/order"

	"github.com/xuri/excelize/v2"
)

func Round(x, unit float64) float64 {
	return math.Round(x/unit) * unit
}

func decideManufactureToSpItem(income string, readPiSpContent orderModel.PiSp_content) int {

	tempReturn := 0
	manufacureNum := len(readPiSpContent.Sp.ManufacturerOrder)
	for i := 0; i < manufacureNum; i++ {
		if income == readPiSpContent.Body.Sp.ManufacturerOrder[i].Manufacturer.Name {
			tempReturn = i
		}
	}
	return tempReturn
}

func BuildSp(outputName string, readPiSpContent orderModel.PiSp_content, excelOrPdf string) (filePath string) {

	f, err := excelize.OpenFile("./storage/spModle.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	//最開始的基本資料
	f.SetCellValue("SP", "S5", readPiSpContent.Body.Sp.DeliveryDate)    //交貨期
	f.SetCellValue("SP", "P7", readPiSpContent.Body.Sp.PortOfLoading)   //裝貨港
	f.SetCellValue("SP", "P10", readPiSpContent.Body.Pi.Customer.Name)  //客戶
	f.SetCellValue("SP", "P14", readPiSpContent.Body.Sp.DelTerm)        //Sales Term
	f.SetCellValue("SP", "T7", readPiSpContent.Body.Sp.PortOfDischarge) //卸貨港
	f.SetCellValue("SP", "T10", readPiSpContent.Body.Sp.ContractID)     //合約號:
	f.SetCellValue("SP", "T14", readPiSpContent.Body.Sp.PaymentTerm)    //Payment Term:

	if excelOrPdf == "excel" {
		//spitem的加工廠編號
		var SpitemToManufactureNum [50]int
		var howManySpItem int = len(readPiSpContent.Body.Sp.SpItems)
		theSpitemMoreThan4 := 0
		if theSpitemMoreThan4 > 0 {
			theSpitemMoreThan4 -= 4
			for i := 0; i < theSpitemMoreThan4; i++ {
				f.DuplicateRow("SP", 40)
			}
		}
		for i := 0; i < len(readPiSpContent.Body.Sp.SpItems); i++ {
			SpitemToManufactureNum[i] = decideManufactureToSpItem(readPiSpContent.Body.Sp.SpItems[i].SupplierName)

		}
		the36Position := 36 + theSpitemMoreThan4
		the41Position := 41 + theSpitemMoreThan4
		the42Position := 42 + theSpitemMoreThan4
		the45Position := 45 + theSpitemMoreThan4
		the46Position := 46 + theSpitemMoreThan4
		the47Position := 47 + theSpitemMoreThan4
		the48Position := 48 + theSpitemMoreThan4
		the49Position := 49 + theSpitemMoreThan4
		the50Position := 50 + theSpitemMoreThan4
		the51Position := 51 + theSpitemMoreThan4
		the52Position := 52 + theSpitemMoreThan4
		the53Position := 53 + theSpitemMoreThan4
		the54Position := 54 + theSpitemMoreThan4
		the55Position := 55 + theSpitemMoreThan4
		the56Position := 56 + theSpitemMoreThan4
		the57Position := 57 + theSpitemMoreThan4
		the58Position := 58 + theSpitemMoreThan4

		//***************************************明確位置表格

		//台灣進口--sales Term
		var TWImportSalesTerm [6]string
		for i := 0; i < 5; i++ {
			TWImportSalesTerm[i], _ = excelize.CoordinatesToCellName(8, 9+2*i)
		}
		//廠商名稱/鋼種
		var manufacturerNameAtC19 [5]string
		for i := 0; i < 5; i++ {
			manufacturerNameAtC19[i], _ = excelize.CoordinatesToCellName(3, 19+i)
		}

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
		//sp主要的表單
		var doubleArrayPiTerms [25][100]string
		for i := 0; i < 24; i++ {
			for j := 0; j < howManySpItem; j++ {
				doubleArrayPiTerms[i][j], _ = excelize.CoordinatesToCellName(1+i, 37+j)
			}
		}
		//廠商、...表格
		var doubleArrayManufac [5][100]string
		for i := 0; i < 5; i++ {
			for j, _ := range readPiSpContent.Body.Sp.SpItems {
				doubleArrayManufac[i][j], _ = excelize.CoordinatesToCellName(32+i, 37+j)
			}

		}

		//供應商...表格
		var doubleArraySupplier [5][100]string
		for i := 0; i < 5; i++ {
			for j, _ := range readPiSpContent.Body.Sp.SpItems {
				doubleArraySupplier[i][j], _ = excelize.CoordinatesToCellName(37+i, 37+j)
			}

		}
		//加工廠...表格
		var doubleArrayProcessing [5][100]string
		for i := 0; i < 5; i++ {
			for j, _ := range readPiSpContent.Body.Sp.SpItems {
				doubleArrayProcessing[i][j], _ = excelize.CoordinatesToCellName(42+i, 37+j)
			}

		}
		//廠商實際銷貨(總金額）...表格
		var doubleArrayRealManufac [5][100]string
		for i := 0; i < 5; i++ {
			for j, _ := range readPiSpContent.Body.Sp.SpItems {
				doubleArrayRealManufac[i][j], _ = excelize.CoordinatesToCellName(47+i, 37+j)
			}

		}

		//設定紅字(非表格)總和公式
		var Z45ToAE45 [6]string

		for j := 0; j < 6; j++ {
			Z45ToAE45[j], _ = excelize.CoordinatesToCellName(26+j, the45Position)
		}
		//設定紅字(表格)總和公式
		var AF42ToAY42 [20]string

		for j := 0; j < 20; j++ {
			AF42ToAY42[j], _ = excelize.CoordinatesToCellName(32+j, the42Position)
		}

		//H18總價
		var H19ToH23 [5]string
		for j := 0; j < 5; j++ {
			H19ToH23[j], _ = excelize.CoordinatesToCellName(8, 19+j)
		}
		//P19數量
		var P19ToP23 [5]string
		for j := 0; j < 5; j++ {
			P19ToP23[j], _ = excelize.CoordinatesToCellName(16, 19+j)
		}
		//S19數量
		var S19ToS23 [5]string
		for j := 0; j < 5; j++ {
			S19ToS23[j], _ = excelize.CoordinatesToCellName(19, 19+j)
		}

		//銀行費用公式

		var N50ToN55 [6]string
		for j := 0; j < 6; j++ {
			N50ToN55[j], _ = excelize.CoordinatesToCellName(14, the50Position+j)
		}
		//鋼捲成本
		var K37ToK40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			K37ToK40[j], _ = excelize.CoordinatesToCellName(11, 37+j)
		}

		// 加工費總計
		var O37ToO40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			O37ToO40[j], _ = excelize.CoordinatesToCellName(15, 37+j)
		}
		//毛利
		var Q37ToQ40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			Q37ToQ40[j], _ = excelize.CoordinatesToCellName(17, 37+j)
		}

		//毛利總和
		var T37ToT40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			T37ToT40[j], _ = excelize.CoordinatesToCellName(20, 37+j)
		}
		//採購價
		var Y37ToY40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			Y37ToY40[j], _ = excelize.CoordinatesToCellName(25, 37+j)
		}
		//銷售成本
		var AC37ToAC40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			AC37ToAC40[j], _ = excelize.CoordinatesToCellName(29, 37+j)
		}
		//餘料損失
		var AD37ToAD40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			AD37ToAD40[j], _ = excelize.CoordinatesToCellName(30, 37+j)
		}
		//出口報關
		var AE37ToAE40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			AE37ToAE40[j], _ = excelize.CoordinatesToCellName(31, 37+j)
		}

		//***************************************明確位置表格

		//********************************單純計算會用到的

		The42Bottom := 42
		if theSpitemMoreThan4 > 0 {
			The42Bottom += theSpitemMoreThan4
		}
		The50Bottom := 50
		if theSpitemMoreThan4 > 0 {
			The50Bottom += theSpitemMoreThan4
		}

		var F37ToF40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			F37ToF40[j], _ = excelize.CoordinatesToCellName(6, 37+j)
		}

		var AB37ToAB40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			AB37ToAB40[j], _ = excelize.CoordinatesToCellName(28, 37+j)
		}

		var Z37ToZ40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			Z37ToZ40[j], _ = excelize.CoordinatesToCellName(26, 37+j)
		}

		//Z36ToAE36
		var Z36ToAE36 [6]string
		for j := 0; j < 6; j++ {
			Z36ToAE36[j], _ = excelize.CoordinatesToCellName(26+j, 36)
		}
		//Z41ToAE41
		var Z41ToAE41 [6]string

		for j := 0; j < 6; j++ {
			Z41ToAE41[j], _ = excelize.CoordinatesToCellName(26+j, the41Position)
		}
		//AF42ToAJ42
		var AF42ToAJ42 [5]string

		for i := 0; i < 5; i++ {
			AF42ToAJ42[i], _ = excelize.CoordinatesToCellName(32+i, The42Bottom)
		}
		var E19ToE23 [5]string
		for i := 0; i < 5; i++ {
			E19ToE23[i], _ = excelize.CoordinatesToCellName(5, 19+i)
		}
		//E37ToE40
		var E37ToE40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			E37ToE40[j], _ = excelize.CoordinatesToCellName(5, 37+j)
		}
		var R37ToR40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			R37ToR40[j], _ = excelize.CoordinatesToCellName(18, 37+j)
		}
		//AA37ToAA40
		var AA37ToAA40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			AA37ToAA40[j], _ = excelize.CoordinatesToCellName(27, 37+j)
		}
		var AF36ToAY36 [20]string
		for j := 0; j < 20; j++ {
			AF36ToAY36[j], _ = excelize.CoordinatesToCellName(32+j, 36)
		}

		var AF41ToAY41 [20]string

		for j := 0; j < 20; j++ {
			AF41ToAY41[j], _ = excelize.CoordinatesToCellName(32+j, the41Position)
		}

		var AK42TOAO42 [5]string
		for j := 0; j < 5; j++ {
			AK42TOAO42[j], _ = excelize.CoordinatesToCellName(37+j, The42Bottom)
		}
		//
		var AP42ToAT42 [5]string
		for j := 0; j < 5; j++ {
			AP42ToAT42[j], _ = excelize.CoordinatesToCellName(42+j, The42Bottom)
		}
		var AU42AY42 [5]string
		for j := 0; j < 5; j++ {
			AU42AY42[j], _ = excelize.CoordinatesToCellName(47+j, The42Bottom)
		}
		var M50ToM55 [6]string
		for j := 0; j < 6; j++ {
			M50ToM55[j], _ = excelize.CoordinatesToCellName(13, The50Bottom+j)
		}
		////鋼捲成本
		var H37ToH40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			H37ToH40[j], _ = excelize.CoordinatesToCellName(8, 37+j)
		}
		var I37ToI40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			I37ToI40[j], _ = excelize.CoordinatesToCellName(9, 37+j)
		}
		var J37ToJ40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			J37ToJ40[j], _ = excelize.CoordinatesToCellName(10, 37+j)
		}
		////鋼捲成本
		// 加工費總計
		var U37ToU40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			U37ToU40[j], _ = excelize.CoordinatesToCellName(21, 37+j)
		}
		var V37ToV40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			V37ToV40[j], _ = excelize.CoordinatesToCellName(22, 37+j)
		}
		var W37ToW40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			W37ToW40[j], _ = excelize.CoordinatesToCellName(23, 37+j)
		}
		var X37ToX40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			X37ToX40[j], _ = excelize.CoordinatesToCellName(24, 37+j)
		}

		// 加工費總計

		//毛利
		var G37ToG40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			G37ToG40[j], _ = excelize.CoordinatesToCellName(7, 37+j)
		}

		var L37ToL40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			L37ToL40[j], _ = excelize.CoordinatesToCellName(12, 37+j)
		}

		var M37ToM40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			M37ToM40[j], _ = excelize.CoordinatesToCellName(13, 37+j)
		}
		var N37ToN40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			N37ToN40[j], _ = excelize.CoordinatesToCellName(14, 37+j)
		}

		var P37ToP40 [50]string
		for j := range readPiSpContent.Body.Sp.SpItems {
			P37ToP40[j], _ = excelize.CoordinatesToCellName(16, 37+j)
		}

		//毛利

		//毛利總和
		//採購價

		//********************************單純計算會用到的

		//********************************************************SetCellFormula

		//數量 (MT)計算
		for i := 0; i < 5; i++ {
			f.SetCellFormula("SP", E19ToE23[i], "="+AF42ToAJ42[i])

		}

		// set DoubleArrayManufac formular
		for j := 0; j < 5; j++ {
			for i := range readPiSpContent.Body.Sp.SpItems {
				f.SetCellFormula("SP", doubleArrayManufac[j][i], "=IF("+E37ToE40[i]+"="+strconv.Itoa(j+1)+","+R37ToR40[i]+"/1000,\"0\")")
			}
		}

		// set doubleArraySupplier formular
		for j := 0; j < 5; j++ {
			for i := range readPiSpContent.Body.Sp.SpItems {
				f.SetCellFormula("SP", doubleArraySupplier[j][i], "=IF("+E37ToE40[i]+"="+strconv.Itoa(j+1)+","+AA37ToAA40[i]+",\"0\")")
			}
		}

		// set doubleArrayProcessing formular
		for j := 0; j < 5; j++ {
			for i := range readPiSpContent.Body.Sp.SpItems {
				f.SetCellFormula("SP", doubleArrayProcessing[j][i], "=IF("+F37ToF40[i]+"="+strconv.Itoa(j+1)+","+AB37ToAB40[i]+",\"0\")")
			}
		}

		// set doubleArrayRealManufac formular
		for j := 0; j < 5; j++ {
			for i := range readPiSpContent.Body.Sp.SpItems {
				f.SetCellFormula("SP", doubleArrayRealManufac[j][i], "=IF("+E37ToE40[i]+"="+strconv.Itoa(j+1)+","+Z37ToZ40[i]+",\"0\")")
			}
		}
		//銷貨收入、出貨成本、加工成本、銷售成本、餘料損失、出口報關 各自的總和
		for i := 0; i < 6; i++ {
			f.SetCellFormula("SP", Z45ToAE45[i], "=SUM("+Z36ToAE36[i]+":"+Z41ToAE41[i]+")")
		}
		//廠商（數量） 到  廠商實際銷貨(總金額）的各自總和
		for i := 0; i < 20; i++ {
			f.SetCellFormula("SP", AF42ToAY42[i], "=SUM("+AF36ToAY36[i]+":"+AF41ToAY41[i]+")")
		}
		//總價計算
		for i := 0; i < 5; i++ {
			f.SetCellFormula("SP", H19ToH23[i], "="+AK42TOAO42[i]+"+"+AP42ToAT42[i]+"")
		}
		//I24總和
		f.SetCellFormula("SP", "I24", "=SUM(H19:L23)")

		//P19數量
		for i := 0; i < 5; i++ {
			f.SetCellFormula("SP", P19ToP23[i], "="+E19ToE23[i])

		}

		//P24總和
		f.SetCellFormula("SP", "P24", "=SUM(P19:R23)")

		//S19數量
		for i := 0; i < 5; i++ {
			f.SetCellFormula("SP", S19ToS23[i], "="+AU42AY42[i]+"")
		}
		//T24總和
		f.SetCellFormula("SP", "T24", "=SUM(S19:W23)")

		// 出口費用

		for i := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellFormula("SP", P37ToP40[i], "=$T$"+strconv.Itoa(the52Position)+"")
		}

		//銀行費用公式
		for i := 0; i < 6; i++ {
			f.SetCellFormula("SP", N50ToN55[i], "="+M50ToM55[i]+"*$T$"+strconv.Itoa(the55Position)+"")
		}
		//鋼捲成本
		for i := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellFormula("SP", K37ToK40[i], "="+H37ToH40[i]+"+"+I37ToI40[i]+"+"+J37ToJ40[i]+"")
		}
		// 加工費總計
		for i := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellFormula("SP", O37ToO40[i], "="+U37ToU40[i]+"+"+V37ToV40[i]+"+"+W37ToW40[i]+"+"+X37ToX40[i]+"")
		}
		//毛利
		for i := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellFormula("SP", Q37ToQ40[i], "="+G37ToG40[i]+"-"+K37ToK40[i]+"-"+M37ToM40[i]+"-"+N37ToN40[i]+"-"+O37ToO40[i]+"-"+P37ToP40[i]+"-"+L37ToL40[i]+"")
		}

		//毛利總和
		for i := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellFormula("SP", T37ToT40[i], "="+Q37ToQ40[i]+"*"+R37ToR40[i]+"/1000*$T$"+strconv.Itoa(the55Position)+"")
		}

		//採購價
		for i := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellFormula("SP", Y37ToY40[i], "=("+K37ToK40[i]+"+"+O37ToO40[i]+")-"+J37ToJ40[i]+"")
		}

		//銷貨收入
		for i := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellFormula("SP", Z37ToZ40[i], "=ROUND("+R37ToR40[i]+"*"+G37ToG40[i]+"/1000,2)")
		}

		//出貨成本
		for i := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellFormula("SP", AA37ToAA40[i], "=ROUND(("+H37ToH40[i]+"+"+I37ToI40[i]+")*"+R37ToR40[i]+"/1000,0)")
		}

		//加工成本
		for i := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellFormula("SP", AB37ToAB40[i], "=ROUND("+O37ToO40[i]+"*"+R37ToR40[i]+"/1000,0)")
		}

		//銷售成本
		for i := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellFormula("SP", AC37ToAC40[i], "=ROUND(("+H37ToH40[i]+"+"+I37ToI40[i]+"+"+O37ToO40[i]+")*"+R37ToR40[i]+"/1000,0)")
		}

		//餘料損失
		for i := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellFormula("SP", AD37ToAD40[i], "="+N37ToN40[i]+"*"+R37ToR40[i]+"/1000")
		}
		//出口報關
		for i := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellFormula("SP", AE37ToAE40[i], "=("+J37ToJ40[i]+"+"+L37ToL40[i]+")*"+R37ToR40[i]+"/1000")
		}
		////寫死的
		f.SetCellFormula("SP", "E4", "=T24-I24-S27-S28-S29-S30-H29-H30") //預估利潤(USD)設定
		f.SetCellFormula("SP", "E5", "=E4/T24")                          //毛利率設定
		f.SetCellFormula("SP", "E24", "=SUM(E19:G23)")                   //E24總和

		f.SetCellFormula("SP", "H27", "=V"+strconv.Itoa(the49Position)+"")  //相關成本費用:進貨成本
		f.SetCellFormula("SP", "H28", "=AB"+strconv.Itoa(the45Position)+"") //相關成本費用:進貨成本
		f.SetCellFormula("SP", "H29", "=AD"+strconv.Itoa(the45Position)+"") //相關成本費用:進貨成本
		f.SetCellFormula("SP", "Q28", "=F"+strconv.Itoa(the55Position)+"")  //*每櫃
		f.SetCellFormula("SP", "X31", "=T"+strconv.Itoa(the55Position)+"")  //匯率：USD/MT

		f.SetCellFormula("SP", "S27", "=SUM(M"+strconv.Itoa(the50Position)+":M"+strconv.Itoa(the55Position)+")") //相關成本費用:銷貨成本
		f.SetCellFormula("SP", "S28", "=F"+strconv.Itoa(the58Position)+"")                                       //相關成本費用:銷貨成本
		f.SetCellFormula("SP", "S29", "=E"+strconv.Itoa(the46Position)+"")                                       //相關成本費用:銷貨成本
		f.SetCellFormula("SP", "S30", "=AE"+strconv.Itoa(the45Position)+"")                                      //相關成本費用:銷貨成本

		//寫死的

		////////////////////////////////不要印部分

		f.SetCellFormula("SP", "E"+strconv.Itoa(the46Position)+"", "=M37*Q"+strconv.Itoa(the45Position)+"/1000") //E46佣金金額:
		f.SetCellFormula("SP", "E"+strconv.Itoa(the47Position)+"", "=Z"+strconv.Itoa(the45Position)+"")          //E47銷售總額

		f.SetCellFormula("SP", "F"+strconv.Itoa(the57Position)+"", "=SUM(H"+strconv.Itoa(the51Position)+":H"+strconv.Itoa(the56Position)+")+IF(F"+strconv.Itoa(the50Position)+"=1,H"+strconv.Itoa(the50Position)+",0)") //TTL NT$
		f.SetCellFormula("SP", "F"+strconv.Itoa(the58Position)+"", "=F"+strconv.Itoa(the57Position)+"/T"+strconv.Itoa(the55Position)+"")                                                                                //TTL US$

		f.SetCellFormula("SP", "H"+strconv.Itoa(the50Position)+"", "=5620+(380*Q"+strconv.Itoa(the45Position)+"/1000)") //出口費用公式

		f.SetCellFormula("SP", "H"+strconv.Itoa(the51Position)+"", "=F51*(Q"+strconv.Itoa(the45Position)+"/1000)*$T$"+strconv.Itoa(the55Position)+"") //出口費用公式

		f.SetCellFormula("SP", "H"+strconv.Itoa(the52Position)+"", "=F"+strconv.Itoa(the52Position)+"*$O$"+strconv.Itoa(the47Position)+"*$T$"+strconv.Itoa(the55Position)+"") //出口費用公式
		f.SetCellFormula("SP", "H"+strconv.Itoa(the53Position)+"", "=F"+strconv.Itoa(the53Position)+"*$T$"+strconv.Itoa(the55Position)+"")                                    //出口費用公式
		f.SetCellFormula("SP", "H"+strconv.Itoa(the54Position)+"", "=F"+strconv.Itoa(the54Position)+"*$O$"+strconv.Itoa(the48Position)+"*$T$"+strconv.Itoa(the55Position)+"") //出口費用公式
		f.SetCellFormula("SP", "H"+strconv.Itoa(the55Position)+"", "=F"+strconv.Itoa(the55Position)+"*$O$"+strconv.Itoa(the46Position)+"*$T$"+strconv.Itoa(the55Position)+"") //出口費用公式
		f.SetCellFormula("SP", "H"+strconv.Itoa(the56Position)+"", "=F"+strconv.Itoa(the56Position)+"*$O$"+strconv.Itoa(the46Position)+"*$T$"+strconv.Itoa(the55Position)+"") //出口費用公式

		f.SetCellFormula("SP", "S"+strconv.Itoa(the45Position), "=ROUND(SUM(T"+strconv.Itoa(the36Position)+":T"+strconv.Itoa(the41Position)+"),0)") //台幣總毛利

		f.SetCellFormula("SP", "M"+strconv.Itoa(the57Position), "=SUM(N"+strconv.Itoa(the50Position)+":N"+strconv.Itoa(the55Position)+")") //TTL
		f.SetCellFormula("SP", "O"+strconv.Itoa(the45Position), "=SUM(O"+strconv.Itoa(the46Position)+":O"+strconv.Itoa(the48Position)+")") //O45公式
		f.SetCellFormula("SP", "Q"+strconv.Itoa(the45Position), "=SUM(R"+strconv.Itoa(the36Position)+":R"+strconv.Itoa(the41Position)+")") //Q45公斤公式

		f.SetCellFormula("SP", "S"+strconv.Itoa(the46Position), "=$S$"+strconv.Itoa(the45Position)+"/$T$"+strconv.Itoa(the55Position)+"")                                  //美金總毛利
		f.SetCellFormula("SP", "S"+strconv.Itoa(the47Position), "=S"+strconv.Itoa(the45Position)+"/(Z"+strconv.Itoa(the45Position)+"*$T$"+strconv.Itoa(the55Position)+")") //毛利率

		f.SetCellFormula("SP", "T"+strconv.Itoa(the50Position), "=F"+strconv.Itoa(the57Position)+"+M"+strconv.Itoa(the57Position)+"")                                             //T50合計出口&銀行費用公式
		f.SetCellFormula("SP", "T"+strconv.Itoa(the51Position), "=Q"+strconv.Itoa(the45Position)+"")                                                                              //T51預計出口總重量(KG)公式
		f.SetCellFormula("SP", "T"+strconv.Itoa(the52Position), "=$T$"+strconv.Itoa(the50Position)+"/$T$"+strconv.Itoa(the51Position)+"*1000/$T$"+strconv.Itoa(the55Position)+"") //T52平均出口費用(USD/MT)公式

		f.SetCellFormula("SP", "V"+strconv.Itoa(the49Position), "=AA"+strconv.Itoa(the45Position)+"") //隱藏數字V49

		//********************************************************SetCellFormula

		//********************************************************setcellvalue

		f.SetCellValue("SP", "F"+strconv.Itoa(the50Position), readPiSpContent.Body.Sp.FeeDetail.BulkFobCharges)    //出口費用 散貨FOB費用
		f.SetCellValue("SP", "F"+strconv.Itoa(the51Position), readPiSpContent.Body.Sp.FeeDetail.BulkOceanFreight)  //出口費用 出散貨(每噸)
		f.SetCellValue("SP", "F"+strconv.Itoa(the52Position), readPiSpContent.Body.Sp.FeeDetail.TaiOceanFreight)   //出口費用 台灣出口(20'櫃)
		f.SetCellValue("SP", "F"+strconv.Itoa(the53Position), readPiSpContent.Body.Sp.FeeDetail.CsAmericaPremium)  //出口費用 中南美保費
		f.SetCellValue("SP", "F"+strconv.Itoa(the54Position), readPiSpContent.Body.Sp.FeeDetail.TaiOceanFreight40) //出口費用 台灣出口(40'櫃)
		f.SetCellValue("SP", "F"+strconv.Itoa(the55Position), readPiSpContent.Body.Sp.FeeDetail.ChiOceanFreight)   //出口費用 大陸出口(整櫃)
		f.SetCellValue("SP", "F"+strconv.Itoa(the56Position), readPiSpContent.Body.Sp.FeeDetail.Other)             //出口費用 其他出口費用

		f.SetCellValue("SP", "M"+strconv.Itoa(the50Position), readPiSpContent.Body.Sp.FeeDetail.TtRemittanceFee)           //銀行  費用T/T匯款費用
		f.SetCellValue("SP", "M"+strconv.Itoa(the51Position), readPiSpContent.Body.Sp.FeeDetail.PayTheBalanceAfter30Day)   //銀行  費用30天尾款
		f.SetCellValue("SP", "M"+strconv.Itoa(the52Position), readPiSpContent.Body.Sp.FeeDetail.DpLcRemittanceFee)         //銀行  費用DP or L/C匯款費用
		f.SetCellValue("SP", "M"+strconv.Itoa(the53Position), readPiSpContent.Body.Sp.FeeDetail.DpPremium)                 //銀行  費用DP 保費
		f.SetCellValue("SP", "M"+strconv.Itoa(the54Position), readPiSpContent.Body.Sp.FeeDetail.ForwardLcExpenses)         //銀行  費用遠期L/C費用
		f.SetCellValue("SP", "M"+strconv.Itoa(the55Position), readPiSpContent.Body.Sp.FeeDetail.ForwardLcInterestExpenses) //銀行  費用遠期L/C利息費用
		f.SetCellValue("SP", "M"+strconv.Itoa(the56Position), readPiSpContent.Body.Sp.FeeDetail.Use30DayInterestRate)      //銀行  費用30天利息

		f.SetCellValue("SP", "O"+strconv.Itoa(the46Position), readPiSpContent.Body.Sp.TriangleTradeNum) //三角貿易 預計出貨20'櫃數(triangle)
		f.SetCellValue("SP", "O"+strconv.Itoa(the47Position), readPiSpContent.Body.Sp.TaiExportNum)     //三角貿易 預計出貨20'櫃數(tw)
		f.SetCellValue("SP", "O"+strconv.Itoa(the48Position), readPiSpContent.Body.Sp.TaiExport40Num)   //三角貿易 預計出貨40'櫃數

		f.SetCellValue("SP", "T"+strconv.Itoa(the55Position), readPiSpContent.Body.Sp.Rate) //rate

		f.SetCellValue("SP", "H30", readPiSpContent.Body.Sp.FeeDetail.Other) //H30 other fee
		//********************************************************setcellvalue
		////////////////////////////////不要印部分

		for i, _ := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[0][0+i], i+1)
			//在(A,37)紀錄編號
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[1][0+i], n.Grade)
			//在(B,37)紀錄鋼的材質
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[2][0+i], n.Edge)
			//在(c,37)紀錄edge
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[3][0+i], n.Size)
			//在(d,37)紀錄尺寸
		}

		for i, _ := range readPiSpContent.Body.Sp.SpItems {

			f.SetCellValue("SP", doubleArrayPiTerms[4][0+i], SpitemToManufactureNum[i]+1)

		} //在(E,37)紀錄供應商編號

		for i, _ := range readPiSpContent.Body.Sp.SpItems {

			f.SetCellValue("SP", doubleArrayPiTerms[5][0+i], SpitemToManufactureNum[i]+1)

		} //在(F,37)紀錄加工廠編號

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[6][0+i], n.UnitPrice)
			//在(G,37)紀錄售價
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[7][0+i], n.Price)
			//在(H,37)紀錄盤價
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[8][0+i], n.ThiPremium)
			//在(i,37)紀錄後寬度加價
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[9][0+i], n.CostOfImport)
			//在(J,37)紀錄進口成本
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[11][0+i], n.FobFee)
			//在(L,37)紀錄FobFee
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[12][0+i], n.Commission)
			//在(M,37)紀錄Commission
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[13][0+i], n.RemainLoss)
			//在(N,37)紀錄RemainLoss
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[17][0+i], n.Quantity)
			//在(R,37)紀錄Quantity

		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[20][0+i], n.Non5Mt)
			//在(U,37)紀錄Non5Mt
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[21][0+i], n.Slinging)
			//在(V,37)紀錄Slinging
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[22][0+i], n.Sticker)
			//在(W,37)紀錄Sticker
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[23][0+i], n.Rpcb)
			//在(X,37)紀錄Rpcb
		}

		for i, n := range readPiSpContent.Body.Sp.ManufacturerOrder {

			f.SetCellValue("SP", manufacturerOrderArray[i], n.Manufacturer.Name) //在(C,8)放入n.ManuOrderID
		}

		for i, n := range readPiSpContent.Body.Sp.ManufacturerOrder {

			f.SetCellValue("SP", manufacturerSalesTermArray[i], n.SalesTerm) //在(C,9)放入n.SalesTerm

		}

		for i, n := range readPiSpContent.Body.Sp.ManufacturerOrder {

			f.SetCellValue("SP", TWImportManufac[i], n.ContractID) //在(H,8)放入n.ManuOrderID

		}
		for i, n := range readPiSpContent.Body.Sp.ManufacturerOrder {

			f.SetCellValue("SP", TWImportSalesTerm[i], n.PaymentTerm) //在(H,9)放入n.SalesTerm

		}

		for i, n := range readPiSpContent.Body.Sp.ManufacturerOrder {
			f.SetCellValue("SP", manufacturerNameAtC19[i], n.Manufacturer.Name) //在(C,19)放入n.name
		}

		//存檔
		if err := f.SaveAs(outputName + ".xlsx"); err != nil {
			fmt.Println(err)
		}
		return outputName + ".xlsx"

	} else if excelOrPdf == "pdf" {
		var howManySpItem int = len(readPiSpContent.Body.Sp.SpItems)
		theSpitemMoreThan4 := howManySpItem - 4
		if theSpitemMoreThan4 > 0 {
			for i := 0; i < theSpitemMoreThan4; i++ {
				f.DuplicateRow("SP", 40)
			}
		}
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
		for i := 0; i < 5; i++ {
			manufacturerNameAtC19[i], _ = excelize.CoordinatesToCellName(3, 19+i)
		}

		//中下的關鍵表格
		var doubleArrayPiTerms [26][100]string
		for i := 0; i < 25; i++ {
			for j := 0; j < howManySpItem; j++ {
				doubleArrayPiTerms[i][j], _ = excelize.CoordinatesToCellName(1+i, 37+j)
			}

		}
		//spitem的加工廠編號
		var SpitemToManufactureNum [100]int
		for i := 0; i < len(readPiSpContent.Body.Sp.SpItems); i++ {
			SpitemToManufactureNum[i] = decideManufactureToSpItem(readPiSpContent.Body.Sp.SpItems[i].SupplierName)

		}

		for i, _ := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[0][0+i], i+1)
			//在(A,37)紀錄編號
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[1][0+i], n.Grade)
			//在(B,37)紀錄鋼的材質
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[2][0+i], n.Edge)
			//在(c,37)紀錄edge
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[3][0+i], n.Size)
			//在(d,37)紀錄尺寸
		}

		for i, _ := range readPiSpContent.Body.Sp.SpItems {

			f.SetCellValue("SP", doubleArrayPiTerms[4][0+i], SpitemToManufactureNum[i]+1)

		} //在(E,37)紀錄供應商編號

		for i, _ := range readPiSpContent.Body.Sp.SpItems {

			f.SetCellValue("SP", doubleArrayPiTerms[5][0+i], SpitemToManufactureNum[i]+1)

		} //在(F,37)紀錄加工廠編號

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[6][0+i], n.UnitPrice)
			//在(G,37)紀錄售價
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[7][0+i], n.Price)
			//在(H,37)紀錄盤價
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[8][0+i], n.ThiPremium)
			//在(i,37)紀錄後寬度加價
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[9][0+i], n.CostOfImport)
			//在(J,37)紀錄進口成本
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[11][0+i], n.FobFee)
			//在(L,37)紀錄FobFee
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[12][0+i], n.Commission)
			//在(M,37)紀錄Commission
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[13][0+i], n.RemainLoss)
			//在(N,37)紀錄RemainLoss
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[17][0+i], n.Quantity)
			//在(R,37)紀錄Quantity
			//HERE'S QUESTION
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[20][0+i], n.Non5Mt)
			//在(U,37)紀錄Non5Mt
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[21][0+i], n.Slinging)
			//在(V,37)紀錄Slinging
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[22][0+i], n.Sticker)
			//在(W,37)紀錄Sticker
		}

		for i, n := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[23][0+i], n.Rpcb)
			//在(X,37)紀錄Rpcb
		}

		for i, n := range readPiSpContent.Body.Sp.ManufacturerOrder {

			f.SetCellValue("SP", manufacturerOrderArray[i], n.Manufacturer.Name) //在(C,8)放入n.ManuOrderID
		}

		for i, n := range readPiSpContent.Body.Sp.ManufacturerOrder {

			f.SetCellValue("SP", manufacturerSalesTermArray[i], n.SalesTerm) //在(C,9)放入n.SalesTerm

		}

		for i, n := range readPiSpContent.Body.Sp.ManufacturerOrder {

			f.SetCellValue("SP", TWImportManufac[i], n.ContractID) //在(H,8)放入n.ManuOrderID

		}
		for i, n := range readPiSpContent.Body.Sp.ManufacturerOrder {

			f.SetCellValue("SP", TWImportSalesTerm[i], n.PaymentTerm) //在(H,9)放入n.SalesTerm

		}

		for i, n := range readPiSpContent.Body.Sp.ManufacturerOrder {
			f.SetCellValue("SP", manufacturerNameAtC19[i], n.Manufacturer.Name) //在(C,19)放入n.name
		}

		//鋼捲成本
		var ironArray [100]float64
		for i, n := range readPiSpContent.Body.Sp.SpItems {
			ironArray[i] = n.Price + n.ThiPremium + n.CostOfImport
			f.SetCellValue("SP", doubleArrayPiTerms[10][0+i], ironArray[i])
		}
		// 加工費總計
		var fabricatorArray [100]float64
		for i, n := range readPiSpContent.Body.Sp.SpItems {

			fabricatorArray[i] = n.Non5Mt + n.Slinging + n.Sticker + n.Rpcb
			f.SetCellValue("SP", doubleArrayPiTerms[14][0+i], fabricatorArray[i])

		}

		//出口費用公式計算
		totalExportWeight := 0.0 //預計出口總重量( KG)
		for _, n := range readPiSpContent.Body.Sp.SpItems {
			totalExportWeight += n.Quantity

		} //預計出口總重量( KG)

		exprotTtl := 0.0 //出口費用ttl

		if readPiSpContent.Body.Sp.FeeDetail.BulkFobCharges == 1 {
			exprotTtl = 5620 + (380 * totalExportWeight / 1000) + readPiSpContent.Body.Sp.FeeDetail.BulkOceanFreight*readPiSpContent.Body.Sp.Rate*(totalExportWeight/1000) + readPiSpContent.Body.Sp.FeeDetail.TaiOceanFreight*readPiSpContent.Body.Sp.Rate*readPiSpContent.Body.Sp.TaiExportNum + readPiSpContent.Body.Sp.FeeDetail.CsAmericaPremium*readPiSpContent.Body.Sp.Rate + readPiSpContent.Body.Sp.FeeDetail.TaiOceanFreight40*readPiSpContent.Body.Sp.TaiExport40Num*readPiSpContent.Body.Sp.Rate + readPiSpContent.Body.Sp.FeeDetail.ChiOceanFreight*readPiSpContent.Body.Sp.TriangleTradeNum*readPiSpContent.Body.Sp.Rate + readPiSpContent.Body.Sp.FeeDetail.Other*readPiSpContent.Body.Sp.TriangleTradeNum*readPiSpContent.Body.Sp.Rate //出口費用ttl
		} else {
			exprotTtl = readPiSpContent.Body.Sp.FeeDetail.BulkOceanFreight*readPiSpContent.Body.Sp.Rate*(totalExportWeight/1000) + readPiSpContent.Body.Sp.FeeDetail.TaiOceanFreight*readPiSpContent.Body.Sp.Rate*readPiSpContent.Body.Sp.TaiExportNum + readPiSpContent.Body.Sp.FeeDetail.CsAmericaPremium*readPiSpContent.Body.Sp.Rate + readPiSpContent.Body.Sp.FeeDetail.TaiOceanFreight40*readPiSpContent.Body.Sp.TaiExport40Num*readPiSpContent.Body.Sp.Rate + readPiSpContent.Body.Sp.FeeDetail.ChiOceanFreight*readPiSpContent.Body.Sp.TriangleTradeNum*readPiSpContent.Body.Sp.Rate + readPiSpContent.Body.Sp.FeeDetail.Other*readPiSpContent.Body.Sp.TriangleTradeNum*readPiSpContent.Body.Sp.Rate //出口費用ttl
		}

		feeTtl := readPiSpContent.Body.Sp.FeeDetail.TtRemittanceFee*readPiSpContent.Body.Sp.Rate + readPiSpContent.Body.Sp.FeeDetail.PayTheBalanceAfter30Day*readPiSpContent.Body.Sp.Rate + readPiSpContent.Body.Sp.FeeDetail.DpLcRemittanceFee*readPiSpContent.Body.Sp.Rate + readPiSpContent.Body.Sp.FeeDetail.DpPremium*readPiSpContent.Body.Sp.Rate + readPiSpContent.Body.Sp.FeeDetail.ForwardLcExpenses*readPiSpContent.Body.Sp.Rate + readPiSpContent.Body.Sp.FeeDetail.ForwardLcInterestExpenses*readPiSpContent.Body.Sp.Rate //銀行費用ttl

		totalExportAndBankFee := exprotTtl + feeTtl //合計出口&銀行費用

		averageExportFee := totalExportAndBankFee / totalExportWeight * 1000 / readPiSpContent.Body.Sp.Rate //平均出口費用(USD/MT)，也是 出口費用

		//出口費用
		for i, _ := range readPiSpContent.Body.Sp.SpItems {
			f.SetCellValue("SP", doubleArrayPiTerms[15][0+i], averageExportFee)
		}
		//毛利
		var grossProfit [100]float64
		for i, n := range readPiSpContent.Body.Sp.SpItems {
			grossProfit[i] = n.UnitPrice - ironArray[i] - n.FobFee - n.Commission - n.RemainLoss - fabricatorArray[i] - totalExportAndBankFee

			f.SetCellValue("SP", doubleArrayPiTerms[16][0+i], grossProfit[i])
		}
		//毛利總計
		var totalGrossProfit [100]float64
		for i, n := range readPiSpContent.Body.Sp.SpItems {
			totalGrossProfit[i] = grossProfit[i] * n.Quantity / 1000 * readPiSpContent.Body.Sp.Rate
			f.SetCellValue("SP", doubleArrayPiTerms[19][0+i], totalGrossProfit[i])
		}
		//採購價
		var buyingPrice [100]float64
		for i, n := range readPiSpContent.Body.Sp.SpItems {
			buyingPrice[i] = (ironArray[i] + fabricatorArray[i]) - n.CostOfImport
			f.SetCellValue("SP", doubleArrayPiTerms[24][0+i], buyingPrice[i])
		}

		//////////////////////////////////////////////////////////////////////////////////////回頭做E19~E23的數量

		var E19ToE23 [5]string
		for i := 0; i < 5; i++ {
			E19ToE23[i], _ = excelize.CoordinatesToCellName(5, 19+i)
		}

		var E19ToE23Amount [100]float64
		for i := 0; i < 5; i++ {
			for j, n := range readPiSpContent.Body.Sp.SpItems {
				if SpitemToManufactureNum[j] == i {
					E19ToE23Amount[i] += n.Quantity
				}
			}
			f.SetCellValue("SP", E19ToE23[i], E19ToE23Amount[i]/1000)
		}

		//總價 (USD)
		var H19ToH23 [5]string
		for i := 0; i < 5; i++ {
			H19ToH23[i], _ = excelize.CoordinatesToCellName(8, 19+i)
		}

		var H19ToH23supply [100]float64  //出貨成本計算
		var H19ToH23Procing [100]float64 //加工成本計算
		for i := 0; i < 5; i++ {
			for j, n := range readPiSpContent.Body.Sp.SpItems {
				if SpitemToManufactureNum[j] == i {
					H19ToH23supply[i] += Round((n.Price+n.ThiPremium)*n.Quantity/1000, 1)
					H19ToH23Procing[i] += Round(fabricatorArray[j]*n.Quantity/1000, 1)
				}
				fmt.Println("H19ToH23Procing", H19ToH23Procing[i])
			}

			f.SetCellValue("SP", H19ToH23[i], H19ToH23supply[i]+H19ToH23Procing[i])
		}
		//數量 (MT)
		var P19ToP23 [5]string
		for i := 0; i < 5; i++ {
			P19ToP23[i], _ = excelize.CoordinatesToCellName(16, 19+i)
		}
		for i := 0; i < 5; i++ {
			f.SetCellValue("SP", P19ToP23[i], E19ToE23Amount[i]/1000)
		}
		//總價 (USD)
		var S19ToS23 [5]string
		for i := 0; i < 5; i++ {
			S19ToS23[i], _ = excelize.CoordinatesToCellName(19, 19+i)
		}
		var S19ToS23TotalPrice [100]float64
		for i := 0; i < 5; i++ {
			for j, n := range readPiSpContent.Body.Sp.SpItems {
				if SpitemToManufactureNum[j] == i {
					S19ToS23TotalPrice[i] += n.UnitPrice * n.Quantity / 1000
				}
			}
			f.SetCellValue("SP", S19ToS23[i], S19ToS23TotalPrice[i])
		}

		//E24數量 (MT)
		totalAmount := 0.0
		for i := 0; i < 5; i++ {
			totalAmount += E19ToE23Amount[i] / 1000
		}
		f.SetCellValue("SP", "E24", totalAmount)

		//I24總價 (USD)
		totalPrice := 0.0
		for i := 0; i < 5; i++ {
			totalPrice += H19ToH23supply[i] + H19ToH23Procing[i]
		}
		//T24總價 (USD) 總和
		totalPriceT24 := 0.0
		for i := 0; i < 5; i++ {
			totalPriceT24 += S19ToS23TotalPrice[i]
		}

		f.SetCellValue("SP", "T24", totalPriceT24) //T24總價 (USD) 總和

		f.SetCellValue("SP", "I24", totalPrice) //P24數量 (MT)

		f.SetCellValue("SP", "P24", totalAmount) //S24總價 (USD)

		//鋼捲成本(盤價＋厚度＋進口)
		H27IronCost := 0.0

		for _, n := range readPiSpContent.Body.Sp.SpItems {

			H27IronCost += Round((n.Price+n.ThiPremium)*n.Quantity/1000, 1)

		}
		f.SetCellValue("SP", "H27", H27IronCost)

		//加工費用總計(包裝＋拋砂＋貼膜＋修邊＋切版)

		H28TotalFabricatorFee := 0.0

		for i, n := range readPiSpContent.Body.Sp.SpItems {

			H28TotalFabricatorFee += Round(fabricatorArray[i]*n.Quantity/1000, 1)

		}
		f.SetCellValue("SP", "H28", H28TotalFabricatorFee)

		//餘料損失

		H29LosingOther := 0.0

		for _, n := range readPiSpContent.Body.Sp.SpItems {

			H29LosingOther += n.RemainLoss * n.Quantity / 1000

		}
		f.SetCellValue("SP", "H29", H29LosingOther)

		//其他費用

		f.SetCellValue("SP", "H30", readPiSpContent.Body.Sp.FeeDetail.Other)

		//銀行費用

		S27BankFee := 0.0
		S27BankFee = readPiSpContent.Body.Sp.FeeDetail.TtRemittanceFee + readPiSpContent.Body.Sp.FeeDetail.PayTheBalanceAfter30Day + readPiSpContent.Body.Sp.FeeDetail.DpLcRemittanceFee + readPiSpContent.Body.Sp.FeeDetail.DpPremium + readPiSpContent.Body.Sp.FeeDetail.ForwardLcExpenses + readPiSpContent.Body.Sp.FeeDetail.ForwardLcInterestExpenses
		f.SetCellValue("SP", "S27", S27BankFee)
		//*每櫃
		f.SetCellValue("SP", "Q28", readPiSpContent.Body.Sp.FeeDetail.ChiOceanFreight)
		// 運保費
		f.SetCellValue("SP", "S28", exprotTtl/readPiSpContent.Body.Sp.Rate)

		//佣金
		f.SetCellValue("SP", "S29", readPiSpContent.Body.Sp.SpItems[0].Commission*totalExportWeight/1000)
		//進出口報關
		S30ExInportFee := 0.0

		for _, n := range readPiSpContent.Body.Sp.SpItems {

			S30ExInportFee += (n.CostOfImport + n.FobFee) * n.Quantity / 1000

		}
		f.SetCellValue("SP", "S30", S30ExInportFee)
		//匯率：USD/MT
		f.SetCellValue("SP", "X31", readPiSpContent.Body.Sp.Rate)

		//預估利潤(USD)在一開始
		predictProfit := 0.0
		predictProfit = totalPriceT24 - totalPrice - H29LosingOther - S27BankFee - exprotTtl/readPiSpContent.Body.Sp.Rate - readPiSpContent.Body.Sp.SpItems[0].Commission*totalExportWeight/1000 - S30ExInportFee - readPiSpContent.Body.Sp.FeeDetail.Other
		f.SetCellValue("SP", "E4", predictProfit)

		//毛利率
		f.SetCellValue("SP", "E5", predictProfit/totalPriceT24)
		//刪除欄
		for i := 0; i < 26; i++ {
			f.RemoveCol("SP", "Z")
		}

		for i := 0; i < 16; i++ {
			f.RemoveRow("SP", 43+theSpitemMoreThan4)
		}

		//存檔

		if err := f.SaveAs(outputName + ".xlsx"); err != nil {
			fmt.Println(err)
		}
		return outputName + ".xlsx"

	} else if excelOrPdf != "pdf" && excelOrPdf != "excel" {
		fmt.Print("請在檔名後輸入版本,\"pdf\"或,\"excel\"")
		return ""
	} else {
		return ""
	}

}
