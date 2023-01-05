package build_file

import (
	"fmt"
	"strings"

	orderModel "eirc.app/internal/v1/structure/order"
	"github.com/xuri/excelize/v2"
)

func countStringLine(i string) int {
	return (len(i) / 155)
}
func BuildPi(outputName string, readPiContent orderModel.Pi_content) (filePath string) {
	f, err := excelize.OpenFile("./storage/piModle.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	//判斷C7是不是PROMETAL
	ifPsIsTrue := strings.Contains(readPiContent.ContractId, "PS")
	if ifPsIsTrue {
		f.SetCellValue("PI", "C7", "PROMETAL INTERNATIONAL CO., LTD")
	} else {
		f.SetCellValue("PI", "C7", "MEGLOBE CO., LTD")
	}
	cell, err := f.GetCellValue("PI", "C7")
	if err != nil {
		fmt.Println(err)
		return
	}

	//Pi開頭的基本資料:
	f.SetCellValue("PI", "C9", readPiContent.Customer.Name)
	f.SetCellValue("PI", "C10", readPiContent.Attention)
	f.SetCellValue("PI", "C11", readPiContent.Tel)
	f.SetCellValue("PI", "C12", readPiContent.Address)
	f.SetCellValue("PI", "I7", readPiContent.ContractId)
	f.SetCellValue("PI", "I9", readPiContent.OrdDate)

	//discribe部分:

	discribeStyle, _ := f.NewStyle(`{
		"alignment":{
			"horizontal":"left",
			"wrap_text":true,
			"vertical":"top"
		},
		"font": {
			"family": "Times New Roman"	,
			"size" : 11	
		}
	}`)

	discribeStr := readPiContent.Description
	discribeSplitString := strings.Split(discribeStr, "\n")
	discribeLen := len(discribeSplitString)
	discribeIsMoreThan2 := discribeLen > 2
	theInsertDiscribeNumber := discribeLen - 2
	if discribeIsMoreThan2 {
		for i := 0; i < theInsertDiscribeNumber; i++ {
			f.InsertRow("PI", 15)
		}
	}

	var discribemArray [20]string
	for i := 0; i < 20; i++ {
		discribemArray[i], _ = excelize.CoordinatesToCellName(3, 14+i)
	}

	for i := 0; i < discribeLen; i++ {
		PaymentTermPosition, _ := excelize.CoordinatesToCellName(3, 14+i)
		mergeLinePosition, _ := excelize.CoordinatesToCellName(7, 14+i)
		f.SetCellValue("PI", discribemArray[i], discribeSplitString[i])
		f.MergeCell("PI", PaymentTermPosition, mergeLinePosition)
		f.SetCellStyle("PI", PaymentTermPosition, mergeLinePosition, discribeStyle)
	}

	//Pi第一個表格
	var countPi int = len(readPiContent.PiItems)
	var theInsertPiNumber int = countPi - 4
	countPiIsMoreThan4 := countPi > 4
	piFirstArrayBottom := 23
	if discribeIsMoreThan2 {
		piFirstArrayBottom = 23 + theInsertDiscribeNumber
	}
	piFirstArrayCoppy := 22
	if discribeIsMoreThan2 {
		piFirstArrayCoppy = 22 + theInsertDiscribeNumber
	}
	if countPiIsMoreThan4 {
		for i := 0; i < theInsertPiNumber; i++ {
			f.DuplicateRowTo("PI", piFirstArrayCoppy, piFirstArrayBottom+i)
		}
	}

	//用雙陣列做第一個表格
	piFirstArrayHead := 20
	if discribeIsMoreThan2 {
		piFirstArrayHead = 20 + theInsertDiscribeNumber
	}
	var doublePiArray [11][100]string

	for i := 0; i < 11; i++ {
		for j := 0; j < countPi; j++ {
			doublePiArray[i][j], _ = excelize.CoordinatesToCellName(1+i, piFirstArrayHead+j)
		}

	}

	for i, n := range readPiContent.PiItems {
		f.SetCellValue("PI", doublePiArray[0][0+i], n.ItemNum)
		//在(A,19)紀錄編號
	}
	for i, n := range readPiContent.PiItems {
		f.SetCellValue("PI", doublePiArray[1][0+i], n.Grade)
		//在(B,19)紀錄Grade
	}
	for i, n := range readPiContent.PiItems {
		f.SetCellValue("PI", doublePiArray[2][0+i], n.Edge)
		//在(C,19)紀錄Edge
	}
	for i, n := range readPiContent.PiItems {
		f.SetCellValue("PI", doublePiArray[3][0+i], n.Size)
		//在(D,19)紀錄Size
	}
	for i, n := range readPiContent.PiItems {
		f.SetCellValue("PI", doublePiArray[5][0+i], n.Quantity)
		//在(F,19)紀錄Quantity
	}
	for i, _ := range readPiContent.PiItems {
		f.SetCellValue("PI", doublePiArray[6][0+i], "USD")
		//在(G,19)寫下USD
	}
	for i, n := range readPiContent.PiItems {
		f.SetCellValue("PI", doublePiArray[7][0+i], n.UnitPrice)
		//在(H,19)紀錄UnitPrice
	}
	for i, _ := range readPiContent.PiItems {
		f.SetCellValue("PI", doublePiArray[9][0+i], "USD")
		//在(J,19)寫下USD
	}
	for i, n := range readPiContent.PiItems {
		f.SetCellValue("PI", doublePiArray[10][0+i], n.Amount)
		//在(K,19)紀錄Amount
	}

	//紀錄total
	totalPosition := 24
	if countPiIsMoreThan4 && discribeIsMoreThan2 {
		totalPosition = 24 + theInsertDiscribeNumber + theInsertPiNumber
	} else if countPiIsMoreThan4 {
		totalPosition = 24 + theInsertPiNumber
	} else if discribeIsMoreThan2 {
		totalPosition = 24 + theInsertDiscribeNumber
	}
	quantityPosition, _ := excelize.CoordinatesToCellName(6, totalPosition)
	amountPosition, _ := excelize.CoordinatesToCellName(11, totalPosition)
	f.SetCellValue("PI", quantityPosition, readPiContent.Quantity)
	f.SetCellValue("PI", amountPosition, readPiContent.Amount)

	//Delivery Allowance之後的表格:

	var DeliveryAllowance [11]string
	DeliveryAllowancePosition := 27
	if countPiIsMoreThan4 && discribeIsMoreThan2 {
		DeliveryAllowancePosition = 27 + theInsertDiscribeNumber + theInsertPiNumber
	} else if countPiIsMoreThan4 {
		DeliveryAllowancePosition = 27 + theInsertPiNumber
	} else if discribeIsMoreThan2 {
		DeliveryAllowancePosition = 27 + theInsertDiscribeNumber
	}
	for i := 0; i < 11; i++ {
		DeliveryAllowance[i], _ = excelize.CoordinatesToCellName(3, DeliveryAllowancePosition+i)
	}

	f.SetCellValue("PI", DeliveryAllowance[0], readPiContent.DelAllowance)
	f.SetCellValue("PI", DeliveryAllowance[1], readPiContent.ThiAllowance)
	f.SetCellValue("PI", DeliveryAllowance[2], readPiContent.Packing)
	f.SetCellValue("PI", DeliveryAllowance[3], readPiContent.PackageWei)
	f.SetCellValue("PI", DeliveryAllowance[4], readPiContent.ShippingMark)
	f.SetCellValue("PI", DeliveryAllowance[5], readPiContent.InvoiceAmount)
	f.SetCellValue("PI", DeliveryAllowance[6], readPiContent.Shipment)
	f.SetCellValue("PI", DeliveryAllowance[7], readPiContent.DelTerm)
	f.SetCellValue("PI", DeliveryAllowance[8], readPiContent.ParShipment)
	f.SetCellValue("PI", DeliveryAllowance[9], readPiContent.PortOfLoading)

	//PaymentTerm 這裡要用到/n
	howManyPaymentTermNewLine := strings.Count(readPiContent.PaymentTerm, "\n") + 1
	theInsertPaymentTermRowNumber := howManyPaymentTermNewLine - 2
	PaymentTermIsMoreThan2 := theInsertPaymentTermRowNumber > 0
	InsertPaymentTermHead := 38
	if countPiIsMoreThan4 && discribeIsMoreThan2 {
		InsertPaymentTermHead = 38 + theInsertDiscribeNumber + theInsertPiNumber
	} else if countPiIsMoreThan4 {
		InsertPaymentTermHead = 38 + theInsertPiNumber
	} else if discribeIsMoreThan2 {
		InsertPaymentTermHead = 38 + theInsertDiscribeNumber
	}

	//插入新的row，因為原版位置不夠
	if PaymentTermIsMoreThan2 {
		for i := 0; i < theInsertPaymentTermRowNumber; i++ {
			f.InsertRow("PI", InsertPaymentTermHead)
		}
	}

	//編寫格式規則
	paymentTermStyle, _ := f.NewStyle(`{
		"alignment":{
			"horizontal":"left",
			"wrap_text":true,
			"vertical":"top"
		},
		"font": {
			"family": "Times New Roman"	,
			"size" : 11	
		}
	}`)

	paymentTerStart := 37

	if countPiIsMoreThan4 && discribeIsMoreThan2 {
		paymentTerStart = 37 + theInsertDiscribeNumber + theInsertPiNumber
	} else if countPiIsMoreThan4 {
		paymentTerStart = 37 + theInsertPiNumber
	} else if discribeIsMoreThan2 {
		paymentTerStart = 37 + theInsertDiscribeNumber
	}

	paymentTermStr := readPiContent.PaymentTerm
	paymentTermSplitString := strings.Split(paymentTermStr, "\n")
	paymentTermLen := len(paymentTermSplitString)
	var paymentTermArrat [20]string
	for i := 0; i < 20; i++ {
		paymentTermArrat[i], _ = excelize.CoordinatesToCellName(3, paymentTerStart+i)
	}

	for i := 0; i < paymentTermLen; i++ {
		PaymentTermPosition, _ := excelize.CoordinatesToCellName(3, paymentTerStart+i)
		mergeLinePosition, _ := excelize.CoordinatesToCellName(8, paymentTerStart+i)
		f.SetCellValue("PI", paymentTermArrat[i], paymentTermSplitString[i])
		f.MergeCell("PI", PaymentTermPosition, mergeLinePosition)
		f.SetCellStyle("PI", PaymentTermPosition, mergeLinePosition, paymentTermStyle)
	}

	//Beneficiary Name表格:
	var BeneficiaryName [6]string
	BeneficiaryNameStart := 40
	if countPiIsMoreThan4 && PaymentTermIsMoreThan2 && discribeIsMoreThan2 {
		BeneficiaryNameStart = 40 + theInsertPiNumber + theInsertPaymentTermRowNumber + theInsertDiscribeNumber
	} else if countPiIsMoreThan4 && PaymentTermIsMoreThan2 {
		BeneficiaryNameStart = 40 + theInsertPiNumber + theInsertPaymentTermRowNumber
	} else if countPiIsMoreThan4 && discribeIsMoreThan2 {
		BeneficiaryNameStart = 40 + theInsertPiNumber + theInsertDiscribeNumber
	} else if PaymentTermIsMoreThan2 && discribeIsMoreThan2 {
		BeneficiaryNameStart = 40 + theInsertPaymentTermRowNumber + theInsertDiscribeNumber
	} else if PaymentTermIsMoreThan2 {
		BeneficiaryNameStart = 40 + theInsertPaymentTermRowNumber
	} else if countPiIsMoreThan4 {
		BeneficiaryNameStart = 40 + theInsertPiNumber
	} else if discribeIsMoreThan2 {
		BeneficiaryNameStart = 40 + theInsertDiscribeNumber
	}

	for i := 0; i < 5; i++ {
		BeneficiaryName[i], _ = excelize.CoordinatesToCellName(3, BeneficiaryNameStart+i)

	}

	f.SetCellValue("PI", BeneficiaryName[0], readPiContent.RemittanceBeneficiaryInfo.Name)
	f.SetCellValue("PI", BeneficiaryName[1], readPiContent.RemittanceBeneficiaryInfo.AcNo)
	f.SetCellValue("PI", BeneficiaryName[2], readPiContent.RemittanceBeneficiaryInfo.Bank)
	f.SetCellValue("PI", BeneficiaryName[3], readPiContent.RemittanceBeneficiaryInfo.SwiftCode)
	f.SetCellValue("PI", BeneficiaryName[4], readPiContent.RemittanceBeneficiaryInfo.Address)

	//terms部分 要用到/n
	//howManyTermsNewLine := strings.Count(readPiContent.Terms, "\n") + 1
	//theInertTermsRowNumber := howManyTermsNewLine - 2
	//因為行數不夠 要插入的行數

	//terms的格式設定

	TermStyle, _ := f.NewStyle(`{
			"alignment":{
				"horizontal":"left",
				"wrap_text":true,
				"vertical":"top"
			},
			"font": {
				"family": "Times New Roman",
				"size" : 10
			}
		}`)

	str := readPiContent.Terms
	termsSplitString := strings.Split(str, "\n")
	termsSplitStringLength := len(termsSplitString)
	var topOfTheTerms int = 47
	if countPiIsMoreThan4 && PaymentTermIsMoreThan2 && discribeIsMoreThan2 {
		topOfTheTerms = 47 + theInsertPiNumber + theInsertPaymentTermRowNumber + theInsertDiscribeNumber
	} else if countPiIsMoreThan4 && PaymentTermIsMoreThan2 {
		topOfTheTerms = 47 + theInsertPiNumber + theInsertPaymentTermRowNumber
	} else if countPiIsMoreThan4 && discribeIsMoreThan2 {
		topOfTheTerms = 47 + theInsertPiNumber + theInsertDiscribeNumber
	} else if PaymentTermIsMoreThan2 && discribeIsMoreThan2 {
		topOfTheTerms = 47 + theInsertPaymentTermRowNumber + theInsertDiscribeNumber
	} else if PaymentTermIsMoreThan2 {
		topOfTheTerms = 47 + theInsertPaymentTermRowNumber
	} else if countPiIsMoreThan4 {
		topOfTheTerms = 47 + theInsertPiNumber
	} else if discribeIsMoreThan2 {
		topOfTheTerms = 47 + theInsertDiscribeNumber
	}
	for i := 0; i < termsSplitStringLength; i++ {
		f.InsertRow("PI", topOfTheTerms)
	}
	var termsNewLineArray [20]int
	for i := 0; i < termsSplitStringLength; i++ {
		termsNewLineArray[i] = countStringLine(termsSplitString[i])
		fmt.Println("第", i, "筆資料共有這麼多行=", termsNewLineArray[i])
	}

	var countInsertString int = 0
	l, _ := excelize.CoordinatesToCellName(11, 1)
	for i := 0; i < termsSplitStringLength; i++ {
		for j := 0; j < termsNewLineArray[i]; j++ {
			f.InsertRow("PI", topOfTheTerms)
			countInsertString++
			topOfTheTerms += 1
			fmt.Println("現在到", topOfTheTerms)
		}

		k, _ := excelize.CoordinatesToCellName(1, topOfTheTerms)

		l, _ = excelize.CoordinatesToCellName(11, topOfTheTerms-int(termsNewLineArray[i]))

		f.MergeCell("PI", k, l)
		f.SetCellStyle("PI", k, l, TermStyle)
		f.SetCellValue("PI", l, termsSplitString[i])

		topOfTheTerms += 1
	}

	countInsertString += termsSplitStringLength
	fmt.Println("我們插入了這些行", countInsertString)

	termsIsMoreThan1 := countInsertString > 0

	//公司及購買者部分:
	buttomPosition := 50
	if countPiIsMoreThan4 && termsIsMoreThan1 && PaymentTermIsMoreThan2 && discribeIsMoreThan2 {
		buttomPosition = 50 + theInsertPiNumber + countInsertString + theInsertPaymentTermRowNumber + theInsertDiscribeNumber
		//1、2、3、4
	} else if countPiIsMoreThan4 && termsIsMoreThan1 && PaymentTermIsMoreThan2 {
		buttomPosition = 50 + theInsertPiNumber + countInsertString + theInsertPaymentTermRowNumber
	} else if countPiIsMoreThan4 && discribeIsMoreThan2 && PaymentTermIsMoreThan2 {
		buttomPosition = 50 + theInsertPiNumber + theInsertDiscribeNumber + theInsertPaymentTermRowNumber
	} else if countPiIsMoreThan4 && termsIsMoreThan1 && discribeIsMoreThan2 {
		buttomPosition = 50 + theInsertPiNumber + theInsertDiscribeNumber + countInsertString
	} else if termsIsMoreThan1 && discribeIsMoreThan2 && PaymentTermIsMoreThan2 {
		buttomPosition = 50 + theInsertDiscribeNumber + countInsertString + theInsertPaymentTermRowNumber
	} else if countPiIsMoreThan4 && termsIsMoreThan1 {
		buttomPosition = 50 + theInsertPiNumber + countInsertString
	} else if countPiIsMoreThan4 && PaymentTermIsMoreThan2 {
		buttomPosition = 50 + theInsertPiNumber + theInsertPaymentTermRowNumber
	} else if countPiIsMoreThan4 && discribeIsMoreThan2 {
		buttomPosition = 50 + theInsertPiNumber + theInsertDiscribeNumber
	} else if termsIsMoreThan1 && PaymentTermIsMoreThan2 {
		buttomPosition = 50 + countInsertString + theInsertPaymentTermRowNumber
	} else if termsIsMoreThan1 && discribeIsMoreThan2 {
		buttomPosition = 50 + countInsertString + theInsertDiscribeNumber
	} else if PaymentTermIsMoreThan2 && discribeIsMoreThan2 {
		buttomPosition = 50 + theInsertPaymentTermRowNumber + theInsertDiscribeNumber
	} else if countPiIsMoreThan4 {
		buttomPosition = 50 + theInsertPiNumber
	} else if termsIsMoreThan1 {
		buttomPosition = 50 + countInsertString
	} else if PaymentTermIsMoreThan2 {
		buttomPosition = 50 + theInsertPaymentTermRowNumber
	} else if discribeIsMoreThan2 {
		buttomPosition = 50 + theInsertDiscribeNumber
	}
	buttomSellerPosition, _ := excelize.CoordinatesToCellName(2, buttomPosition)
	buttomBuyerPosition, _ := excelize.CoordinatesToCellName(7, buttomPosition)
	f.SetCellValue("PI", buttomSellerPosition, cell)
	f.SetCellValue("PI", buttomBuyerPosition, readPiContent.Customer.Name)

	//存檔
	if err := f.SaveAs("./storage/order/" + outputName + ".xlsx"); err != nil {
		fmt.Println(err)
	}

	//回傳連結改 public (路由)
	return "/public/order/" + outputName + ".xlsx"
}
