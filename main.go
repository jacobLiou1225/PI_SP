package main

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
			OID             string `json:"o_id"`
			ContractID      string `json:"contract_id"`
			Attention       string `json:"attention"`
			Tel             string `json:"tel"`
			Address         string `json:"address"`
			OrdDate         string `json:"ord_date"`
			Description     string `json:"description"`
			DelAllowance    string `json:"del_allowance"`
			ThiAllowance    string `json:"thi_allowance"`
			Packing         string `json:"packing"`
			PackageWei      string `json:"package_wei"`
			ShippingMark    string `json:"shipping_mark"`
			InvoiceAmount   string `json:"invoice_amount"`
			Shipment        string `json:"shipment"`
			DelTerm         string `json:"del_term"`
			ParShipment     string `json:"par_shipment"`
			PortOfLoading   string `json:"port_of_loading"`
			PaymentTerm     string `json:"payment_term"`
			PortOfDischarge string `json:"port_of_discharge"`
			Transhipment    string `json:"transhipment"`
			Remark          string `json:"remark"`
			Terms           string `json:"terms"`
			Quantity        int    `json:"quantity"`
			Amount          int    `json:"amount"`
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
				PiItemID  string `json:"pi_item_id"`
				ItemNum   int    `json:"item_num"`
				Grade     string `json:"grade"`
				Edge      string `json:"edge"`
				Size      string `json:"size"`
				Quantity  int    `json:"quantity"`
				UnitPrice int    `json:"unit_price"`
				Amount    int    `json:"amount"`
			} `json:"PiItems"`
		} `json:"PI"`
		Sp struct {
			OID                     string `json:"o_id"`
			RemittanceBank          string `json:"remittance_bank"`
			DeliveryDate            string `json:"delivery_date"`
			PortOfLoading           string `json:"port_of_loading"`
			DelTerm                 string `json:"del_term"`
			PortOfDischarge         string `json:"port_of_discharge"`
			ContractID              string `json:"contract_id"`
			PaymentTerm             string `json:"payment_term"`
			ImportExportDeclaration int    `json:"import_export_declaration"`
			OtherFee                int    `json:"other_fee"`
			TriangleTradeNum        int    `json:"triangle_trade_num"`
			TaiExportNum            int    `json:"tai_export_num"`
			TaiExport40Num          int    `json:"tai_export_40_num"`
			Remark                  string `json:"remark"`
			Rate                    int    `json:"rate"`
			CoilCost                int    `json:"coil_cost"`
			TotalProcessingCost     int    `json:"total_processing_cost"`
			RemainLoss              int    `json:"remain_loss"`
			BankFee                 int    `json:"bank_fee"`
			ShippingInsuranceFee    int    `json:"shipping_insurance_fee"`
			Commission              int    `json:"commission"`
			Quantity                int    `json:"quantity"`
			GrossProfitNt           int    `json:"gross_profit_nt"`
			GrossProfitUs           int    `json:"gross_profit_us"`
			GrossMargin             string `json:"gross_margin"`
			SalesRevenue            int    `json:"sales_revenue"`
			ShippingCost            int    `json:"shipping_cost"`
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
				SpItemID            string `json:"sp_item_id"`
				ItemNum             int    `json:"item_num"`
				Grade               string `json:"grade"`
				Edge                string `json:"edge"`
				Size                string `json:"size"`
				SupplierName        string `json:"supplier_name"`
				FabricatorName      string `json:"fabricator_name"`
				UnitPrice           int    `json:"unit_price"`
				Price               int    `json:"price"`
				ThiPremium          int    `json:"thi_premium"`
				CostOfImport        int    `json:"cost_of_import"`
				CoilCost            int    `json:"coil_cost"`
				FobFee              int    `json:"fob_fee"`
				Commission          int    `json:"commission"`
				RemainLoss          int    `json:"remain_loss"`
				TotalProcessingCost int    `json:"total_processing_cost"`
				ExportFee           int    `json:"export_fee"`
				GrossProfit         int    `json:"gross_profit"`
				Quantity            int    `json:"quantity"`
				TotalGrossProfit    int    `json:"total_gross_profit"`
				Non5Mt              int    `json:"non_5mt"`
				Slinging            int    `json:"slinging"`
				Sticker             int    `json:"sticker"`
				Rpcb                int    `json:"rpcb"`
				PurchasePrice       int    `json:"purchase_price"`
				SalesRevenue        int    `json:"sales_revenue"`
				ShippingCost        int    `json:"shipping_cost"`
				ProcessingCost      int    `json:"processing_cost"`
				SalesCost           int    `json:"sales_cost"`
				RemainLossCost      int    `json:"remain_loss_cost"`
				ExportDeclaration   int    `json:"export_declaration"`
				BankFee             int    `json:"bank_fee"`
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
					Name     string `json:"name"`
					Quantity int    `json:"quantity"`
					Total    int    `json:"total"`
				} `json:"SupplierFabricatorCost"`
				SupplierRevenue []struct {
					Name     string `json:"name"`
					Quantity int    `json:"quantity"`
					Total    int    `json:"total"`
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

func main() {
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
	f.SetCellValue("PI", "C14", readPiContent.Body.Pi.Description)
	f.SetCellValue("PI", "I9", readPiContent.Body.Pi.OrdDate)

	//Pi第一個表格
	var countPi int = len(readPiContent.Body.Pi.PiItems)
	var theInsertPiNumber int = countPi - 5
	countPiIsMoreThan5 := countPi > 5

	if countPiIsMoreThan5 == true {
		for i := 0; i < theInsertPiNumber; i++ {
			f.InsertRow("PI", 23+i)
		}
	}
	//用雙陣列做第一個表格
	var doublePiArray [11][100]string

	for i := 0; i < 11; i++ {
		for j := 0; j < countPi; j++ {
			doublePiArray[i][j], _ = excelize.CoordinatesToCellName(1+i, 19+j)
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
	if countPiIsMoreThan5 == true {
		quantityPosition, _ := excelize.CoordinatesToCellName(6, 24+theInsertPiNumber)
		amountPosition, _ := excelize.CoordinatesToCellName(11, 24+theInsertPiNumber)
		f.SetCellValue("PI", quantityPosition, readPiContent.Body.Pi.Quantity)
		f.SetCellValue("PI", amountPosition, readPiContent.Body.Pi.Amount)
	} else {
		f.SetCellValue("PI", "F24", readPiContent.Body.Pi.Quantity)
		f.SetCellValue("PI", "K24", readPiContent.Body.Pi.Amount)
	}

	//Delivery Allowance之後的表格:

	var DeliveryAllowance [11]string

	if countPiIsMoreThan5 == true {
		for i := 0; i < 11; i++ {
			DeliveryAllowance[i], _ = excelize.CoordinatesToCellName(3, theInsertPiNumber+27+i)
		}
	} else {
		for i := 0; i < 11; i++ {
			DeliveryAllowance[i], _ = excelize.CoordinatesToCellName(3, 27+i)
		}
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
	theInertPaymentTermRowNumber := howManyPaymentTermNewLine - 3
	PaymentTermIsMoreThan2 := theInertPaymentTermRowNumber > 0
	//插入新的row，因為原版位置不夠
	if PaymentTermIsMoreThan2 {
		if countPiIsMoreThan5 == true {
			for i := 0; i < theInertPaymentTermRowNumber; i++ {
				f.InsertRow("PI", 39+theInsertPiNumber)
			}
		} else {
			for i := 0; i < theInertPaymentTermRowNumber; i++ {
				f.InsertRow("PI", 39)
			}
		}

	}
	//編寫格式規則
	//{"alignment":{"horizontal":"center"},"font":{"bold":true,"italic":true}}
	paymentTermStyle, _ := f.NewStyle(`{
		"alignment":{
			"horizontal":"left",
			"wrap_text":true,
			"vertical":"top"
		},
		"font": {
			"family": "Times New Roman"	,
			"size" : 9	
		}
	}`)

	//寫進PaymentTerm表格中
	if countPiIsMoreThan5 == true {
		PaymentTermPosition, _ := excelize.CoordinatesToCellName(3, 37+theInsertPiNumber)
		mergeLinePosition, _ := excelize.CoordinatesToCellName(3, 37+theInsertPiNumber+howManyPaymentTermNewLine)
		f.SetCellValue("PI", DeliveryAllowance[10], readPiContent.Body.Pi.PaymentTerm)
		f.MergeCell("PI", PaymentTermPosition, mergeLinePosition)
		f.SetCellStyle("PI", PaymentTermPosition, mergeLinePosition, paymentTermStyle)
	} else {
		PaymentTermPosition, _ := excelize.CoordinatesToCellName(3, 37)
		mergeLinePosition, _ := excelize.CoordinatesToCellName(3, 37+howManyPaymentTermNewLine)
		f.SetCellValue("PI", DeliveryAllowance[10], readPiContent.Body.Pi.PaymentTerm)
		f.MergeCell("PI", PaymentTermPosition, mergeLinePosition)
		f.SetCellStyle("PI", PaymentTermPosition, mergeLinePosition, paymentTermStyle)
	}

	//f.SetCellFormula("PI", "G10", "=LEN(C37)-LEN(SUBSTITUTE(C37,\")\",\"\"))")
	//theNumberOfNewLine, _ := f.GetCellValue("PI", "G10")
	//convertNewLine, _ := strconv.Atoi(theNumberOfNewLine)
	//fmt.Println(convertNewLine)

	/*formulaType, ref := excelize.STCellFormulaTypeArray, "L1:L1"
	f.SetCellFormula("PI", "L1", "={1,2,\"a\",\"b\"}",
		excelize.FormulaOpts{Ref: &ref, Type: &formulaType})
	*/

	//appale, _ := f.CalcCellValue("PI", "G10")
	//fmt.Println("G10有" + appale)

	//Beneficiary Name表格:
	var BeneficiaryName [6]string

	if countPiIsMoreThan5 == true && PaymentTermIsMoreThan2 == true {
		for i := 0; i < 6; i++ {
			BeneficiaryName[i], _ = excelize.CoordinatesToCellName(3, theInsertPiNumber+theInertPaymentTermRowNumber+41+i)
		}
	} else if countPiIsMoreThan5 == true {
		for i := 0; i < 6; i++ {
			DeliveryAllowance[i], _ = excelize.CoordinatesToCellName(3, 41+i+theInsertPiNumber)
		}
	} else if PaymentTermIsMoreThan2 == true {
		for i := 0; i < 6; i++ {
			DeliveryAllowance[i], _ = excelize.CoordinatesToCellName(3, 41+i+theInertPaymentTermRowNumber)
		}
	} else {
		for i := 0; i < 6; i++ {
			DeliveryAllowance[i], _ = excelize.CoordinatesToCellName(3, 41+i)
		}
	}

	f.SetCellValue("PI", BeneficiaryName[0], readPiContent.Body.Pi.Customer.BeneficiaryInfo.Name)
	f.SetCellValue("PI", BeneficiaryName[1], readPiContent.Body.Pi.Customer.BeneficiaryInfo.AcNo)
	f.SetCellValue("PI", BeneficiaryName[2], readPiContent.Body.Pi.Customer.BeneficiaryInfo.Bank)
	f.SetCellValue("PI", BeneficiaryName[3], readPiContent.Body.Pi.Customer.BeneficiaryInfo.SwiftCode)
	f.SetCellValue("PI", BeneficiaryName[4], readPiContent.Body.Pi.Customer.BeneficiaryInfo.Address)

	//terms部分 要用到/n
	howManyTermsNewLine := strings.Count(readPiContent.Body.Pi.Terms, "\n") + 1
	theInertTermsRowNumber := howManyTermsNewLine - 2
	termsIsMoreThan1 := theInertTermsRowNumber > 0

	termsString := readPiContent.Body.Pi.Terms
	termsStringSplit := strings.Split(termsString, "\n")
	howManySplitLine := len(termsStringSplit)
	fmt.Println(termsStringSplit[2])
	fmt.Println("分開了這麼多行:+", howManySplitLine)

	if termsIsMoreThan1 == true {
		if countPiIsMoreThan5 == true && PaymentTermIsMoreThan2 == true {
			for i := 0; i < theInertTermsRowNumber; i++ {
				f.InsertRow("PI", 48+theInsertPiNumber+theInertPaymentTermRowNumber)
			}
		} else if countPiIsMoreThan5 {
			for i := 0; i < theInertTermsRowNumber; i++ {
				f.InsertRow("PI", 48+theInsertPiNumber)
			}
		} else if PaymentTermIsMoreThan2 {
			for i := 0; i < theInertTermsRowNumber; i++ {
				f.InsertRow("PI", 48+theInertPaymentTermRowNumber)
			}
		} else {
			for i := 0; i < theInertTermsRowNumber; i++ {
				f.InsertRow("PI", 48)
			}
		}

	}
	//terms的格式設定
	TermStyle, _ := f.NewStyle(`{
		"alignment":{
			"horizontal":"left",
			"wrap_text":true,
			"vertical":"top"
		},
		"font": {
			"family": "Times New Roman",
			"size" : 8.5
		}
	}`)

	//建立表格: terms
	if countPiIsMoreThan5 == true && PaymentTermIsMoreThan2 == true {
		termPosition, _ := excelize.CoordinatesToCellName(1, 48+theInsertPiNumber+theInertPaymentTermRowNumber)
		termsMergeLinePosition, _ := excelize.CoordinatesToCellName(1, 48+theInsertPiNumber+theInertPaymentTermRowNumber+howManyTermsNewLine)
		f.SetCellValue("PI", termPosition, readPiContent.Body.Pi.Terms)
		f.MergeCell("PI", termPosition, termsMergeLinePosition)
		f.SetCellStyle("PI", termPosition, termsMergeLinePosition, TermStyle)
	} else if countPiIsMoreThan5 == true {
		termPosition, _ := excelize.CoordinatesToCellName(1, 48+theInsertPiNumber)
		termsMergeLinePosition, _ := excelize.CoordinatesToCellName(1, 48+theInsertPiNumber+howManyTermsNewLine)
		f.SetCellValue("PI", termPosition, readPiContent.Body.Pi.Terms)
		f.MergeCell("PI", termPosition, termsMergeLinePosition)
		f.SetCellStyle("PI", termPosition, termsMergeLinePosition, TermStyle)
	} else if PaymentTermIsMoreThan2 == true {
		termPosition, _ := excelize.CoordinatesToCellName(1, 48+theInertPaymentTermRowNumber)
		termsMergeLinePosition, _ := excelize.CoordinatesToCellName(1, 48+theInertPaymentTermRowNumber+howManyTermsNewLine)
		f.SetCellValue("PI", termPosition, readPiContent.Body.Pi.Terms)
		f.MergeCell("PI", termPosition, termsMergeLinePosition)
		f.SetCellStyle("PI", termPosition, termsMergeLinePosition, TermStyle)
	} else {

		termPosition, _ := excelize.CoordinatesToCellName(1, 48)
		termsMergeLinePosition, _ := excelize.CoordinatesToCellName(1, 48+howManyTermsNewLine)
		f.SetCellValue("PI", termPosition, readPiContent.Body.Pi.Terms)
		f.MergeCell("PI", termPosition, termsMergeLinePosition)
		f.SetCellStyle("PI", termPosition, termsMergeLinePosition, TermStyle)

	}

	//seller部分:
	if countPiIsMoreThan5 == true && termsIsMoreThan1 == true && PaymentTermIsMoreThan2 {
		sellerPosition, _ := excelize.CoordinatesToCellName(2, 52+theInsertPiNumber+theInertPaymentTermRowNumber+theInertTermsRowNumber)
		f.SetCellValue("PI", sellerPosition, cell)
	} else if countPiIsMoreThan5 == true && termsIsMoreThan1 == true {
		sellerPosition, _ := excelize.CoordinatesToCellName(2, 52+theInsertPiNumber+theInertTermsRowNumber)
		f.SetCellValue("PI", sellerPosition, cell)
	} else if countPiIsMoreThan5 == true && PaymentTermIsMoreThan2 == true {
		sellerPosition, _ := excelize.CoordinatesToCellName(2, 52+theInsertPiNumber+theInertPaymentTermRowNumber)
		f.SetCellValue("PI", sellerPosition, cell)
	} else if termsIsMoreThan1 == true && PaymentTermIsMoreThan2 == true {
		sellerPosition, _ := excelize.CoordinatesToCellName(2, 52+theInertTermsRowNumber+theInertPaymentTermRowNumber)
		f.SetCellValue("PI", sellerPosition, cell)
	} else if countPiIsMoreThan5 == true {
		sellerPosition, _ := excelize.CoordinatesToCellName(2, 52+theInsertPiNumber)
		f.SetCellValue("PI", sellerPosition, cell)
	} else if termsIsMoreThan1 == true {
		sellerPosition, _ := excelize.CoordinatesToCellName(2, 52+theInertTermsRowNumber)
		f.SetCellValue("PI", sellerPosition, cell)
	} else if PaymentTermIsMoreThan2 == true {
		sellerPosition, _ := excelize.CoordinatesToCellName(2, 52+theInertPaymentTermRowNumber)
		f.SetCellValue("PI", sellerPosition, cell)
	} else {
		sellerPosition, _ := excelize.CoordinatesToCellName(2, 52)
		f.SetCellValue("PI", sellerPosition, cell)
	}

	//公司及購買者部分:
	if countPiIsMoreThan5 == true && termsIsMoreThan1 == true && PaymentTermIsMoreThan2 {

		buttomSellerPosition, _ := excelize.CoordinatesToCellName(2, 54+theInsertPiNumber+theInertPaymentTermRowNumber+theInertTermsRowNumber)
		buttomBuyerPosition, _ := excelize.CoordinatesToCellName(7, 54+theInsertPiNumber+theInertPaymentTermRowNumber+theInertTermsRowNumber)
		f.SetCellValue("PI", buttomSellerPosition, cell)
		f.SetCellValue("PI", buttomBuyerPosition, readPiContent.Body.Pi.Customer.Name)
	} else if countPiIsMoreThan5 == true && termsIsMoreThan1 == true {
		buttomSellerPosition, _ := excelize.CoordinatesToCellName(2, 54+theInsertPiNumber+theInertTermsRowNumber)
		buttomBuyerPosition, _ := excelize.CoordinatesToCellName(7, 54+theInsertPiNumber+theInertTermsRowNumber)
		f.SetCellValue("PI", buttomSellerPosition, cell)
		f.SetCellValue("PI", buttomBuyerPosition, readPiContent.Body.Pi.Customer.Name)
	} else if countPiIsMoreThan5 == true && PaymentTermIsMoreThan2 == true {
		buttomSellerPosition, _ := excelize.CoordinatesToCellName(2, 54+theInsertPiNumber+theInertPaymentTermRowNumber)
		buttomBuyerPosition, _ := excelize.CoordinatesToCellName(7, 54+theInsertPiNumber+theInertPaymentTermRowNumber)
		f.SetCellValue("PI", buttomSellerPosition, cell)
		f.SetCellValue("PI", buttomBuyerPosition, readPiContent.Body.Pi.Customer.Name)
	} else if termsIsMoreThan1 == true && PaymentTermIsMoreThan2 == true {
		buttomSellerPosition, _ := excelize.CoordinatesToCellName(2, 54+theInertPaymentTermRowNumber+theInertTermsRowNumber)
		buttomBuyerPosition, _ := excelize.CoordinatesToCellName(7, 54+theInertPaymentTermRowNumber+theInertTermsRowNumber)
		f.SetCellValue("PI", buttomSellerPosition, cell)
		f.SetCellValue("PI", buttomBuyerPosition, readPiContent.Body.Pi.Customer.Name)
	} else if countPiIsMoreThan5 == true {
		buttomSellerPosition, _ := excelize.CoordinatesToCellName(2, 54+theInsertPiNumber)
		buttomBuyerPosition, _ := excelize.CoordinatesToCellName(7, 54+theInsertPiNumber)
		f.SetCellValue("PI", buttomSellerPosition, cell)
		f.SetCellValue("PI", buttomBuyerPosition, readPiContent.Body.Pi.Customer.Name)
	} else if termsIsMoreThan1 == true {
		buttomSellerPosition, _ := excelize.CoordinatesToCellName(2, 54+theInertTermsRowNumber)
		buttomBuyerPosition, _ := excelize.CoordinatesToCellName(7, 54+theInertTermsRowNumber)
		f.SetCellValue("PI", buttomSellerPosition, cell)
		f.SetCellValue("PI", buttomBuyerPosition, readPiContent.Body.Pi.Customer.Name)
	} else if PaymentTermIsMoreThan2 == true {
		buttomSellerPosition, _ := excelize.CoordinatesToCellName(2, 54+theInertPaymentTermRowNumber)
		buttomBuyerPosition, _ := excelize.CoordinatesToCellName(7, 54+theInertPaymentTermRowNumber)
		f.SetCellValue("PI", buttomSellerPosition, cell)
		f.SetCellValue("PI", buttomBuyerPosition, readPiContent.Body.Pi.Customer.Name)
	} else {
		f.SetCellValue("PI", "B54", cell)
		f.SetCellValue("PI", "G54", readPiContent.Body.Pi.Customer.Name)
	}

	//存檔
	if err := f.SaveAs("piForHo222222企業.xlsx"); err != nil {
		fmt.Println(err)
	}

	fmt.Println("sp長度=", len(readPiContent.Body.Sp.SpItems))

	/*

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
	*/
}
