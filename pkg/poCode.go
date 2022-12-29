package pkg

import (
	"encoding/json"
	"fmt"
	"io/ioutil"
	"net/http"
	"time"

	"github.com/xuri/excelize/v2"
)

type poJson struct {
	Code      int       `json:"code"`
	Timestamp time.Time `json:"timestamp"`
	Body      struct {
		PurID              string    `json:"pur_id"`
		OID                string    `json:"o_id"`
		OrdNum             string    `json:"ord_num"`
		Attention          string    `json:"attention"`
		OrdFrom            string    `json:"ord_from"`
		PoDate             string    `json:"po_date"`
		ExpectDate         string    `json:"expect_date"`
		WeightTolerance    string    `json:"weight_tolerance"`
		ThicknessTolerance string    `json:"thickness_tolerance"`
		TermsOfTrade       string    `json:"terms_of_trade"`
		DischargePort      string    `json:"discharge_port"`
		MarkFormat         string    `json:"mark_format"`
		Requirements       string    `json:"requirements"`
		Remark             string    `json:"remark"`
		BonitaCaseID       string    `json:"bonita_case_id"`
		CreatedAt          time.Time `json:"created_at"`
		CreatedBy          string    `json:"created_by"`
		UpdatedAt          time.Time `json:"updated_at"`
		UpdatedBy          string    `json:"updated_by"`
		PlateWeight        int       `json:"plate_weight"`
		CoilWeight         int       `json:"coil_weight"`
		TotalQuantity      float64   `json:"total_quantity"`
		TotalAmount        float64   `json:"total_amount"`
		Order              struct {
			OID         string `json:"o_id"`
			ContractID  string `json:"contract_id"`
			Attention   string `json:"attention"`
			Tel         string `json:"tel"`
			Address     string `json:"address"`
			OrdDate     string `json:"ord_date"`
			Description string `json:"description"`
		} `json:"Order"`
		PoItems []struct {
			PoItemID      string `json:"po_item_id"`
			ItemNum       int    `json:"item_num"`
			Grade         string `json:"grade"`
			Edge          string `json:"edge"`
			Size          string `json:"size"`
			Quantity      int    `json:"quantity"`
			FobFoshan     int    `json:"fob_foshan"`
			FinishedPrice int    `json:"finished_price"`
			Amount        int    `json:"amount"`
			Remark        string `json:"remark"`
			IsAssigned    bool   `json:"is_assigned"`
		} `json:"PoItems"`
		OriginalCoil []struct {
			OriCoilID      string `json:"ori_coil_id"`
			Grade          string `json:"grade"`
			SteelPlantName string `json:"steel_plant_name"`
			Size           string `json:"size"`
			Quantity       int    `json:"quantity"`
		} `json:"OriginalCoil"`
		BackingPaper        []interface{} `json:"backing_paper"`
		Sticker             []interface{} `json:"sticker"`
		Packing             []interface{} `json:"packing"`
		WeightPerPiece      []interface{} `json:"weight_per_piece"`
		SprayWord           []interface{} `json:"spray_word"`
		Diameter            []interface{} `json:"diameter"`
		Pallet              []interface{} `json:"pallet"`
		ShippingMark        []interface{} `json:"shipping_mark"`
		DirectionOfEntrance []interface{} `json:"direction_of_entrance"`
	} `json:"body"`
}

func BuildPo(outputName string) (filePath string) {

	f, err := excelize.OpenFile("poModle.xlsx")
	if err != nil {
		fmt.Println(err)
		return
	}
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	res, err := http.Get("https://api.testing.eirc.app/meglobe/v1.0/purchase-order/e7dbc0b1-b7a9-4f71-8f5e-b556294a9518")
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
	var readPiContent poJson
	json.Unmarshal(body, &readPiContent)
	//開始的表格 寫死內容
	f.SetCellValue("PO", "B3", readPiContent.Body.Order.Tel)
	f.SetCellValue("PO", "C7", readPiContent.Body.Attention)
	f.SetCellValue("PO", "C8", readPiContent.Body.OrdFrom)
	f.SetCellValue("PO", "H7", readPiContent.Body.ExpectDate)
	f.SetCellValue("PO", "H8", readPiContent.Body.Order.ContractID)
	f.SetCellValue("PO", "H9", readPiContent.Body.OrdNum)
	f.SetCellFormula("PO", "C9", "=H7+45")

	//插入row數
	poItemLength := len(readPiContent.Body.PoItems)
	fmt.Println(poItemLength)
	if poItemLength > 4 {
		for i := 0; i < poItemLength-4; i++ {
			f.DuplicateRowTo("PO", 13, 14+i)
		}

	}
	//編號設定
	var NumerOfArray [10]string
	for i := 0; i < poItemLength; i++ {
		NumerOfArray[i], _ = excelize.CoordinatesToCellName(1, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", NumerOfArray[i], i+1)
	}

	//鋼種設定
	var IronOfArray [10]string
	for i := 0; i < poItemLength; i++ {
		IronOfArray[i], _ = excelize.CoordinatesToCellName(2, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", IronOfArray[i], readPiContent.Body.PoItems[i].Grade)
	}

	//S/M設定
	var SMOfArray [10]string
	for i := 0; i < poItemLength; i++ {
		SMOfArray[i], _ = excelize.CoordinatesToCellName(3, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", SMOfArray[i], readPiContent.Body.PoItems[i].Edge)
	}

	//Size設定
	var sizeArray [10]string
	for i := 0; i < poItemLength; i++ {
		sizeArray[i], _ = excelize.CoordinatesToCellName(4, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", sizeArray[i], readPiContent.Body.PoItems[i].Size)
	}

	//訂單重量設定
	var weightOfArray [10]string
	for i := 0; i < poItemLength; i++ {
		weightOfArray[i], _ = excelize.CoordinatesToCellName(6, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", weightOfArray[i], readPiContent.Body.PoItems[i].Quantity)
	}

	//FOB設定
	var FOBOfArray [10]string
	for i := 0; i < poItemLength; i++ {
		FOBOfArray[i], _ = excelize.CoordinatesToCellName(7, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", FOBOfArray[i], readPiContent.Body.PoItems[i].FobFoshan)
	}
	//總額(USD)
	var H12TOHXX [10]string
	for i := 0; i < poItemLength; i++ {
		H12TOHXX[i], _ = excelize.CoordinatesToCellName(8, 12+i)
	}
	var F12TOFXX [10]string
	for i := 0; i < poItemLength; i++ {
		F12TOFXX[i], _ = excelize.CoordinatesToCellName(6, 12+i)
	}
	var G12TOGXX [10]string
	for i := 0; i < poItemLength; i++ {
		G12TOGXX[i], _ = excelize.CoordinatesToCellName(7, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellFormula("PO", H12TOHXX[i], "="+F12TOFXX[i]+"*"+G12TOGXX[i]+"")
	}

	//成品價/基價
	var productPriceOfArray [10]string
	for i := 0; i < poItemLength; i++ {
		productPriceOfArray[i], _ = excelize.CoordinatesToCellName(9, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", productPriceOfArray[i], readPiContent.Body.PoItems[i].FinishedPrice)
	}

	//備註
	var RemarkOfArray [10]string
	for i := 0; i < poItemLength; i++ {
		RemarkOfArray[i], _ = excelize.CoordinatesToCellName(10, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", RemarkOfArray[i], readPiContent.Body.PoItems[i].Remark)
	}

	//訂單重量(MT) total
	f.SetCellFormula("PO", "F16", "=SUM(F12:F15)")
	//總額(USD) total
	f.SetCellFormula("PO", "H16", "=SUM(H12:H15)")

	//存檔
	if err := f.SaveAs(outputName + ".xlsx"); err != nil {
		fmt.Println(err)
	}
	return outputName + ".xlsx"
}
