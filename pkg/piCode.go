package pkg

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// formal顯示
type Read_Pi_content struct {
	Code      int       `json:"code"`
	Timestamp time.Time `json:"timestamp"`
	Body      struct {
		Pi struct {
			OID             string  `json:"o_id"`
			ContractID      string  `json:"contract_id"`
			Attention       string  `json:"attention"`
			Tel             string  `json:"tel"`
			Address         string  `json:"address"`
			OrdDate         string  `json:"ord_date"`
			Description     string  `json:"description"`
			DelAllowance    string  `json:"del_allowance"`
			ThiAllowance    string  `json:"thi_allowance"`
			Packing         string  `json:"packing"`
			PackageWei      string  `json:"package_wei"`
			ShippingMark    string  `json:"shipping_mark"`
			InvoiceAmount   string  `json:"invoice_amount"`
			Shipment        string  `json:"shipment"`
			DelTerm         string  `json:"del_term"`
			ParShipment     string  `json:"par_shipment"`
			PortOfLoading   string  `json:"port_of_loading"`
			PaymentTerm     string  `json:"payment_term"`
			PortOfDischarge string  `json:"port_of_discharge"`
			Transhipment    string  `json:"transhipment"`
			Remark          string  `json:"remark"`
			Terms           string  `json:"terms"`
			Quantity        float64 `json:"quantity"`
			Amount          float64 `json:"amount"`
			Customer        struct {
				CusID           string    `json:"cus_id"`
				ShortName       string    `json:"short_name"`
				EngName         string    `json:"eng_name"`
				Name            string    `json:"name"`
				ZipCode         string    `json:"zip_code"`
				Address         string    `json:"address"`
				Tel             string    `json:"tel"`
				Fax             string    `json:"fax"`
				Email           string    `json:"email"`
				Attention       string    `json:"attention"`
				AttentionPhone  string    `json:"attention_phone"`
				Remark          string    `json:"remark"`
				IsDeleted       bool      `json:"is_deleted"`
				CreatedAt       time.Time `json:"created_at"`
				BeneficiaryInfo struct {
					BeneID    string `json:"bene_id"`
					Name      string `json:"name"`
					AcNo      string `json:"ac_no"`
					Bank      string `json:"bank"`
					SwiftCode string `json:"swift_code"`
					Address   string `json:"address"`
				} `json:"BeneficiaryInfo"`
			} `json:"Customer"`
			PiItems []struct {
				PiItemID  string  `json:"pi_item_id"`
				ItemNum   int     `json:"item_num"`
				Grade     string  `json:"grade"`
				Edge      string  `json:"edge"`
				Size      string  `json:"size"`
				Quantity  float64 `json:"quantity"`
				UnitPrice float64 `json:"unit_price"`
				Amount    float64 `json:"amount"`
			} `json:"PiItems"`
		} `json:"PI"`
		Sp struct {
			OID                     string  `json:"o_id"`
			RemittanceBank          string  `json:"remittance_bank"`
			DeliveryDate            string  `json:"delivery_date"`
			PortOfLoading           string  `json:"port_of_loading"`
			DelTerm                 string  `json:"del_term"`
			PortOfDischarge         string  `json:"port_of_discharge"`
			ContractID              string  `json:"contract_id"`
			PaymentTerm             string  `json:"payment_term"`
			ImportExportDeclaration int     `json:"import_export_declaration"`
			OtherFee                int     `json:"other_fee"`
			TriangleTradeNum        int     `json:"triangle_trade_num"`
			TaiExportNum            int     `json:"tai_export_num"`
			TaiExport40Num          int     `json:"tai_export_40_num"`
			Remark                  string  `json:"remark"`
			Rate                    int     `json:"rate"`
			CoilCost                int     `json:"coil_cost"`
			TotalProcessingCost     int     `json:"total_processing_cost"`
			RemainLoss              int     `json:"remain_loss"`
			BankFee                 int     `json:"bank_fee"`
			ShippingInsuranceFee    int     `json:"shipping_insurance_fee"`
			Commission              int     `json:"commission"`
			Quantity                float64 `json:"quantity"`
			GrossProfitNt           int     `json:"gross_profit_nt"`
			GrossProfitUs           int     `json:"gross_profit_us"`
			GrossMargin             string  `json:"gross_margin"`
			SalesRevenue            int     `json:"sales_revenue"`
			ShippingCost            int     `json:"shipping_cost"`
			Customer                struct {
				CusID           string    `json:"cus_id"`
				ShortName       string    `json:"short_name"`
				EngName         string    `json:"eng_name"`
				Name            string    `json:"name"`
				ZipCode         string    `json:"zip_code"`
				Address         string    `json:"address"`
				Tel             string    `json:"tel"`
				Fax             string    `json:"fax"`
				Email           string    `json:"email"`
				Attention       string    `json:"attention"`
				AttentionPhone  string    `json:"attention_phone"`
				Remark          string    `json:"remark"`
				IsDeleted       bool      `json:"is_deleted"`
				CreatedAt       time.Time `json:"created_at"`
				BeneficiaryInfo struct {
					BeneID    string `json:"bene_id"`
					Name      string `json:"name"`
					AcNo      string `json:"ac_no"`
					Bank      string `json:"bank"`
					SwiftCode string `json:"swift_code"`
					Address   string `json:"address"`
				} `json:"BeneficiaryInfo"`
			} `json:"Customer"`
			SpItems []struct {
				SpItemID            string  `json:"sp_item_id"`
				ItemNum             int     `json:"item_num"`
				Grade               string  `json:"grade"`
				Edge                string  `json:"edge"`
				Size                string  `json:"size"`
				SupplierName        string  `json:"supplier_name"`
				FabricatorName      string  `json:"fabricator_name"`
				UnitPrice           float64 `json:"unit_price"`
				Price               float64 `json:"price"`
				ThiPremium          int     `json:"thi_premium"`
				CostOfImport        int     `json:"cost_of_import"`
				CoilCost            int     `json:"coil_cost"`
				FobFee              int     `json:"fob_fee"`
				Commission          int     `json:"commission"`
				RemainLoss          int     `json:"remain_loss"`
				TotalProcessingCost int     `json:"total_processing_cost"`
				ExportFee           int     `json:"export_fee"`
				GrossProfit         int     `json:"gross_profit"`
				Quantity            float64 `json:"quantity"`
				TotalGrossProfit    int     `json:"total_gross_profit"`
				Non5Mt              int     `json:"non_5mt"`
				Slinging            int     `json:"slinging"`
				Sticker             int     `json:"sticker"`
				Rpcb                int     `json:"rpcb"`
				PurchasePrice       int     `json:"purchase_price"`
				SalesRevenue        int     `json:"sales_revenue"`
				ShippingCost        int     `json:"shipping_cost"`
				ProcessingCost      int     `json:"processing_cost"`
				SalesCost           int     `json:"sales_cost"`
				RemainLossCost      int     `json:"remain_loss_cost"`
				ExportDeclaration   int     `json:"export_declaration"`
				BankFee             int     `json:"bank_fee"`
			} `json:"SpItems"`
			FeeDetail struct {
				FeeID                     string `json:"fee_id"`
				BulkFobCharges            int    `json:"bulk_fob_charges"`
				BulkOceanFreight          int    `json:"bulk_ocean_freight"`
				TaiOceanFreight           int    `json:"tai_ocean_freight"`
				CsAmericaPremium          int    `json:"cs_america_premium"`
				TaiOceanFreight40         int    `json:"tai_ocean_freight_40"`
				ChiOceanFreight           int    `json:"chi_ocean_freight"`
				Other                     int    `json:"other"`
				TtRemittanceFee           int    `json:"tt_remittance_fee"`
				PayTheBalanceAfter30Day   int    `json:"pay_the_balance_after_30day"`
				DpLcRemittanceFee         int    `json:"dp_lc_remittance_fee"`
				DpPremium                 int    `json:"dp_premium"`
				ForwardLcExpenses         int    `json:"forward_lc_expenses"`
				ForwardLcInterestExpenses int    `json:"forward_lc_interest_expenses"`
				Use30DayInterestRate      int    `json:"use_30day_interest_rate"`
			} `json:"FeeDetail"`
			ManufacturerFee struct {
				SupplierFabricatorCost []struct {
					Name     string  `json:"name"`
					Quantity float64 `json:"quantity"`
					Total    int     `json:"total"`
				} `json:"SupplierFabricatorCost"`
				SupplierRevenue []struct {
					Name     string  `json:"name"`
					Quantity float64 `json:"quantity"`
					Total    int     `json:"total"`
				} `json:"SupplierRevenue"`
			} `json:"ManufacturerFee"`
			ManufacturerOrder []struct {
				ManuOrderID  string `json:"manu_order_id"`
				SalesTerm    string `json:"sales_term"`
				PaymentTerm  string `json:"payment_term"`
				ContractID   string `json:"contract_id"`
				Manufacturer struct {
					ManufacturerID  string `json:"manufacturer_id"`
					EngName         string `json:"eng_name"`
					Name            string `json:"name"`
					BeneficiaryInfo struct {
						Name      string `json:"name"`
						AcNo      string `json:"ac_no"`
						Bank      string `json:"bank"`
						SwiftCode string `json:"swift_code"`
						Address   string `json:"address"`
					} `json:"BeneficiaryInfo"`
				} `json:"Manufacturer"`
			} `json:"ManufacturerOrder"`
		} `json:"SP"`
	} `json:"body"`
}

func countStringLine(i string) int {
	return (len(i) / 155)
}
func BuildPi(outputName string) (filePath string) {
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

	//判斷C7是不是PROMETAL
	ifPsIsTrue := strings.Contains(readPiContent.Body.Pi.ContractID, "PS")
	if ifPsIsTrue == true {
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
	f.SetCellValue("PI", "C9", readPiContent.Body.Pi.Customer.Name)
	f.SetCellValue("PI", "C10", readPiContent.Body.Pi.Attention)
	f.SetCellValue("PI", "C11", readPiContent.Body.Pi.Tel)
	f.SetCellValue("PI", "C12", readPiContent.Body.Pi.Address)
	f.SetCellValue("PI", "I7", readPiContent.Body.Pi.ContractID)
	f.SetCellValue("PI", "I9", readPiContent.Body.Pi.OrdDate)

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

	discribeStr := readPiContent.Body.Pi.Description
	discribeSplitString := strings.Split(discribeStr, "\n")
	discribeLen := len(discribeSplitString)
	discribeIsMoreThan2 := discribeLen > 2
	theInsertDiscribeNumber := discribeLen - 2
	if discribeIsMoreThan2 == true {
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
	var countPi int = len(readPiContent.Body.Pi.PiItems)
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

	////////////////////////////////////////////////////////////////////////////////////////12.26
	//array1的調整字形
	/*
		fistArrayBottomRow := 23 + theInsertPiNumber
		fistArrayHeadRow := 20

		var fistArrayBottom [11]string
		for i := 0; i < 11; i++ {
			fistArrayBottom[i], _ = excelize.CoordinatesToCellName(i+1, fistArrayBottomRow)
		}
		var fistArrayHead [11]string
		for i := 0; i < 11; i++ {
			fistArrayHead[i], _ = excelize.CoordinatesToCellName(i+1, fistArrayHeadRow)
		}

		centerFontStyle, _ := f.NewStyle(`{
			"alignment":{
				"horizontal":"center",
				"vertical":"center"
			},
			"font": {
				"family": "Times New Roman"	,
				"size" : 12
			}
		}`)

		rightFontStyle, _ := f.NewStyle(`{
			"alignment":{
				"horizontal":"right",
				"vertical":"center"
			},
			"font": {
				"family": "Times New Roman"	,
				"size" : 12
			}
		}`)

		leftFontStyle, _ := f.NewStyle(`{
			"alignment":{
				"horizontal":"left",
				"vertical":"center"
			},
			"font": {
				"family": "Times New Roman"	,
				"size" : 12
			}
		}`)

		specialFontStyle, _ := f.NewStyle(`{
			"alignment":{
				"horizontal":"left",
				"vertical":"center"
			},
			"font": {
				"family": "Times New Roman"	,
				"size" : 12
			},
			"number_format":2



		}`)
		amountFontStyle, _ := f.NewStyle(`{
			"alignment":{
				"horizontal":"right",
				"vertical":"center"
			},
			"font": {
				"family": "Times New Roman"	,
				"size" : 12
			}


		}`)
		for i := 0; i < 3; i = i + 2 {
			f.SetCellStyle("PI", fistArrayBottom[i], fistArrayHead[i], centerFontStyle)
		}

		for i := 6; i < 10; i = i + 3 {
			f.SetCellStyle("PI", fistArrayBottom[i], fistArrayHead[i], rightFontStyle)
		}

		for i := 1; i < 4; i = i + 2 {
			f.SetCellStyle("PI", fistArrayBottom[i], fistArrayHead[i], leftFontStyle)
		}
		f.SetCellStyle("PI", fistArrayBottom[7], fistArrayHead[7], centerFontStyle)

		//f.SetCellStyle("PI", fistArrayBottom[10], fistArrayHead[10], amountFontStyle)

		//f.SetCellStyle("PI", fistArrayBottom[5], fistArrayHead[5], specialFontStyle)
	*/
	////////////////////////////////////////////////////////////////////////////////////////12.26
	//用雙陣列做第一個表格
	piFirstArrayHead := 20
	if discribeIsMoreThan2 == true {
		piFirstArrayHead = 20 + theInsertDiscribeNumber
	}
	var doublePiArray [11][100]string

	for i := 0; i < 11; i++ {
		for j := 0; j < countPi; j++ {
			doublePiArray[i][j], _ = excelize.CoordinatesToCellName(1+i, piFirstArrayHead+j)
		}

	}

	for i, n := range readPiContent.Body.Pi.PiItems {
		f.SetCellValue("PI", doublePiArray[0][0+i], n.ItemNum)
		//在(A,19)紀錄編號
	}
	for i, n := range readPiContent.Body.Pi.PiItems {
		f.SetCellValue("PI", doublePiArray[1][0+i], n.Grade)
		//在(B,19)紀錄Grade
	}
	for i, n := range readPiContent.Body.Pi.PiItems {
		f.SetCellValue("PI", doublePiArray[2][0+i], n.Edge)
		//在(C,19)紀錄Edge
	}
	for i, n := range readPiContent.Body.Pi.PiItems {
		f.SetCellValue("PI", doublePiArray[3][0+i], n.Size)
		//在(D,19)紀錄Size
	}
	for i, n := range readPiContent.Body.Pi.PiItems {
		f.SetCellValue("PI", doublePiArray[5][0+i], n.Quantity)
		//在(F,19)紀錄Quantity
	}
	for i, _ := range readPiContent.Body.Pi.PiItems {
		f.SetCellValue("PI", doublePiArray[6][0+i], "USD")
		//在(G,19)寫下USD
	}
	for i, n := range readPiContent.Body.Pi.PiItems {
		f.SetCellValue("PI", doublePiArray[7][0+i], n.UnitPrice)
		//在(H,19)紀錄UnitPrice
	}
	for i, _ := range readPiContent.Body.Pi.PiItems {
		f.SetCellValue("PI", doublePiArray[9][0+i], "USD")
		//在(J,19)寫下USD
	}
	for i, n := range readPiContent.Body.Pi.PiItems {
		f.SetCellValue("PI", doublePiArray[10][0+i], n.Amount)
		//在(K,19)紀錄Amount
	}

	//紀錄total
	totalPosition := 24
	if countPiIsMoreThan4 == true && discribeIsMoreThan2 {
		totalPosition = 24 + theInsertDiscribeNumber + theInsertPiNumber
	} else if countPiIsMoreThan4 == true {
		totalPosition = 24 + theInsertPiNumber
	} else if discribeIsMoreThan2 == true {
		totalPosition = 24 + theInsertDiscribeNumber
	}
	quantityPosition, _ := excelize.CoordinatesToCellName(6, totalPosition)
	amountPosition, _ := excelize.CoordinatesToCellName(11, totalPosition)
	f.SetCellValue("PI", quantityPosition, readPiContent.Body.Pi.Quantity)
	f.SetCellValue("PI", amountPosition, readPiContent.Body.Pi.Amount)

	//Delivery Allowance之後的表格:

	var DeliveryAllowance [11]string
	DeliveryAllowancePosition := 27
	if countPiIsMoreThan4 == true && discribeIsMoreThan2 {
		DeliveryAllowancePosition = 27 + theInsertDiscribeNumber + theInsertPiNumber
	} else if countPiIsMoreThan4 == true {
		DeliveryAllowancePosition = 27 + theInsertPiNumber
	} else if discribeIsMoreThan2 == true {
		DeliveryAllowancePosition = 27 + theInsertDiscribeNumber
	}
	for i := 0; i < 11; i++ {
		DeliveryAllowance[i], _ = excelize.CoordinatesToCellName(3, DeliveryAllowancePosition+i)
	}

	f.SetCellValue("PI", DeliveryAllowance[0], readPiContent.Body.Pi.DelAllowance)
	f.SetCellValue("PI", DeliveryAllowance[1], readPiContent.Body.Pi.ThiAllowance)
	f.SetCellValue("PI", DeliveryAllowance[2], readPiContent.Body.Pi.Packing)
	f.SetCellValue("PI", DeliveryAllowance[3], readPiContent.Body.Pi.PackageWei)
	f.SetCellValue("PI", DeliveryAllowance[4], readPiContent.Body.Pi.ShippingMark)
	f.SetCellValue("PI", DeliveryAllowance[5], readPiContent.Body.Pi.InvoiceAmount)
	f.SetCellValue("PI", DeliveryAllowance[6], readPiContent.Body.Pi.Shipment)
	f.SetCellValue("PI", DeliveryAllowance[7], readPiContent.Body.Pi.DelTerm)
	f.SetCellValue("PI", DeliveryAllowance[8], readPiContent.Body.Pi.ParShipment)
	f.SetCellValue("PI", DeliveryAllowance[9], readPiContent.Body.Pi.PortOfLoading)

	//PaymentTerm 這裡要用到/n
	howManyPaymentTermNewLine := strings.Count(readPiContent.Body.Pi.PaymentTerm, "\n") + 1
	theInsertPaymentTermRowNumber := howManyPaymentTermNewLine - 2
	PaymentTermIsMoreThan2 := theInsertPaymentTermRowNumber > 0
	InsertPaymentTermHead := 38
	if countPiIsMoreThan4 == true && discribeIsMoreThan2 {
		InsertPaymentTermHead = 38 + theInsertDiscribeNumber + theInsertPiNumber
	} else if countPiIsMoreThan4 == true {
		InsertPaymentTermHead = 38 + theInsertPiNumber
	} else if discribeIsMoreThan2 == true {
		InsertPaymentTermHead = 38 + theInsertDiscribeNumber
	}

	//插入新的row，因為原版位置不夠
	if PaymentTermIsMoreThan2 == true {
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

	if countPiIsMoreThan4 == true && discribeIsMoreThan2 {
		paymentTerStart = 37 + theInsertDiscribeNumber + theInsertPiNumber
	} else if countPiIsMoreThan4 == true {
		paymentTerStart = 37 + theInsertPiNumber
	} else if discribeIsMoreThan2 == true {
		paymentTerStart = 37 + theInsertDiscribeNumber
	}

	paymentTermStr := readPiContent.Body.Pi.PaymentTerm
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
	if countPiIsMoreThan4 && PaymentTermIsMoreThan2 && discribeIsMoreThan2 == true {
		BeneficiaryNameStart = 40 + theInsertPiNumber + theInsertPaymentTermRowNumber + theInsertDiscribeNumber
	} else if countPiIsMoreThan4 && PaymentTermIsMoreThan2 == true {
		BeneficiaryNameStart = 40 + theInsertPiNumber + theInsertPaymentTermRowNumber
	} else if countPiIsMoreThan4 && discribeIsMoreThan2 == true {
		BeneficiaryNameStart = 40 + theInsertPiNumber + theInsertDiscribeNumber
	} else if PaymentTermIsMoreThan2 && discribeIsMoreThan2 == true {
		BeneficiaryNameStart = 40 + theInsertPaymentTermRowNumber + theInsertDiscribeNumber
	} else if PaymentTermIsMoreThan2 == true {
		BeneficiaryNameStart = 40 + theInsertPaymentTermRowNumber
	} else if countPiIsMoreThan4 == true {
		BeneficiaryNameStart = 40 + theInsertPiNumber
	} else if discribeIsMoreThan2 == true {
		BeneficiaryNameStart = 40 + theInsertDiscribeNumber
	}

	for i := 0; i < 5; i++ {
		BeneficiaryName[i], _ = excelize.CoordinatesToCellName(3, BeneficiaryNameStart+i)

	}

	f.SetCellValue("PI", BeneficiaryName[0], readPiContent.Body.Pi.Customer.BeneficiaryInfo.Name)
	f.SetCellValue("PI", BeneficiaryName[1], readPiContent.Body.Pi.Customer.BeneficiaryInfo.AcNo)
	f.SetCellValue("PI", BeneficiaryName[2], readPiContent.Body.Pi.Customer.BeneficiaryInfo.Bank)
	f.SetCellValue("PI", BeneficiaryName[3], readPiContent.Body.Pi.Customer.BeneficiaryInfo.SwiftCode)
	f.SetCellValue("PI", BeneficiaryName[4], readPiContent.Body.Pi.Customer.BeneficiaryInfo.Address)

	//terms部分 要用到/n
	//howManyTermsNewLine := strings.Count(readPiContent.Body.Pi.Terms, "\n") + 1
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

	str := readPiContent.Body.Pi.Terms
	termsSplitString := strings.Split(str, "\n")
	termsSplitStringLength := len(termsSplitString)
	var topOfTheTerms int = 47
	if countPiIsMoreThan4 && PaymentTermIsMoreThan2 && discribeIsMoreThan2 == true {
		topOfTheTerms = 47 + theInsertPiNumber + theInsertPaymentTermRowNumber + theInsertDiscribeNumber
	} else if countPiIsMoreThan4 && PaymentTermIsMoreThan2 == true {
		topOfTheTerms = 47 + theInsertPiNumber + theInsertPaymentTermRowNumber
	} else if countPiIsMoreThan4 && discribeIsMoreThan2 == true {
		topOfTheTerms = 47 + theInsertPiNumber + theInsertDiscribeNumber
	} else if PaymentTermIsMoreThan2 && discribeIsMoreThan2 == true {
		topOfTheTerms = 47 + theInsertPaymentTermRowNumber + theInsertDiscribeNumber
	} else if PaymentTermIsMoreThan2 == true {
		topOfTheTerms = 47 + theInsertPaymentTermRowNumber
	} else if countPiIsMoreThan4 == true {
		topOfTheTerms = 47 + theInsertPiNumber
	} else if discribeIsMoreThan2 == true {
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
	if countPiIsMoreThan4 && termsIsMoreThan1 && PaymentTermIsMoreThan2 && discribeIsMoreThan2 == true {
		buttomPosition = 50 + theInsertPiNumber + countInsertString + theInsertPaymentTermRowNumber + theInsertDiscribeNumber
		//1、2、3、4
	} else if countPiIsMoreThan4 && termsIsMoreThan1 && PaymentTermIsMoreThan2 == true {
		buttomPosition = 50 + theInsertPiNumber + countInsertString + theInsertPaymentTermRowNumber
	} else if countPiIsMoreThan4 && discribeIsMoreThan2 && PaymentTermIsMoreThan2 == true {
		buttomPosition = 50 + theInsertPiNumber + theInsertDiscribeNumber + theInsertPaymentTermRowNumber
	} else if countPiIsMoreThan4 && termsIsMoreThan1 && discribeIsMoreThan2 == true {
		buttomPosition = 50 + theInsertPiNumber + theInsertDiscribeNumber + countInsertString
	} else if termsIsMoreThan1 && discribeIsMoreThan2 && PaymentTermIsMoreThan2 == true {
		buttomPosition = 50 + theInsertDiscribeNumber + countInsertString + theInsertPaymentTermRowNumber
	} else if countPiIsMoreThan4 && termsIsMoreThan1 == true {
		buttomPosition = 50 + theInsertPiNumber + countInsertString
	} else if countPiIsMoreThan4 && PaymentTermIsMoreThan2 == true {
		buttomPosition = 50 + theInsertPiNumber + theInsertPaymentTermRowNumber
	} else if countPiIsMoreThan4 && discribeIsMoreThan2 == true {
		buttomPosition = 50 + theInsertPiNumber + theInsertDiscribeNumber
	} else if termsIsMoreThan1 && PaymentTermIsMoreThan2 == true {
		buttomPosition = 50 + countInsertString + theInsertPaymentTermRowNumber
	} else if termsIsMoreThan1 && discribeIsMoreThan2 == true {
		buttomPosition = 50 + countInsertString + theInsertDiscribeNumber
	} else if PaymentTermIsMoreThan2 && discribeIsMoreThan2 == true {
		buttomPosition = 50 + theInsertPaymentTermRowNumber + theInsertDiscribeNumber
	} else if countPiIsMoreThan4 == true {
		buttomPosition = 50 + theInsertPiNumber
	} else if termsIsMoreThan1 == true {
		buttomPosition = 50 + countInsertString
	} else if PaymentTermIsMoreThan2 == true {
		buttomPosition = 50 + theInsertPaymentTermRowNumber
	} else if discribeIsMoreThan2 == true {
		buttomPosition = 50 + theInsertDiscribeNumber
	}
	buttomSellerPosition, _ := excelize.CoordinatesToCellName(2, buttomPosition)
	buttomBuyerPosition, _ := excelize.CoordinatesToCellName(7, buttomPosition)
	f.SetCellValue("PI", buttomSellerPosition, cell)
	f.SetCellValue("PI", buttomBuyerPosition, readPiContent.Body.Pi.Customer.Name)

	//存檔
	if err := f.SaveAs(outputName + ".xlsx"); err != nil {
		fmt.Println(err)
	}
	return outputName + ".xlsx"
}
