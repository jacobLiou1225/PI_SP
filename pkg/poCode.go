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

func CheckBuy(n []interface{}) string {
	fmt.Println("CheckContainer")
	temp_ans := ""
	for _, v := range n {
		valStr := fmt.Sprint(v)
		if strings.Contains(valStr, "不需要") {
			temp_ans = "□需要■不需要"
		} else {
			temp_ans = "■需要□不需要"
		}
	}
	return temp_ans
}
func CheckHeavy(n []interface{}) string {
	fmt.Println("CheckContainer")
	temp_ans := ""
	for _, v := range n {
		valStr := fmt.Sprint(v)
		if strings.Contains(valStr, "不扣重") {
			temp_ans = "□扣重■不扣重"
		} else {
			temp_ans = "■扣重□不扣重"
		}
	}
	return temp_ans

}

func CheckPoliFilm(n []interface{}) string {

	fmt.Println("CheckContainer")
	temp_ans := ""
	for _, v := range n {
		valStr := fmt.Sprint(v)
		if strings.Contains(valStr, "POLI-FILM") {
			temp_ans = "■POLI-FILM"
		} else {
			temp_ans = "□POLI-FILM"
		}
	}
	return temp_ans

}

func CheckNovacel(n []interface{}) string {
	fmt.Println("NOVACEL")
	temp_ans := ""
	for _, v := range n {
		valStr := fmt.Sprint(v)
		if strings.Contains(valStr, "POLI-FILM") {
			temp_ans = "■NOVACEL"
		} else {
			temp_ans = "□NOVACEL"
		}
	}
	return temp_ans

}
func CheckBlueBlack(n []interface{}) string {
	temp_ans := ""
	for _, v := range n {
		valStr := fmt.Sprint(v)
		if strings.Contains(valStr, "黑白") {
			temp_ans = "□藍色■黑白"
		} else {
			temp_ans = "■藍色□黑白"
		}
	}
	return temp_ans
}

/*func CheckMicroLaser(n interface{}) string {
	interfaceString := n.(string)
	if strings.Contains(interfaceString, "100") {
		return "■100 Micro Laser PE, □80 Micro PE, □70 Micro PE□50 Micro PE"
	} else if strings.Contains(interfaceString, "80") {
		return "□100 Micro Laser PE, ■80 Micro PE, □70 Micro PE□50 Micro PE"
	} else if strings.Contains(interfaceString, "70") {
		return "□100 Micro Laser PE, □80 Micro PE, ■70 Micro PE□50 Micro PE"
	} else {
		return "□100 Micro Laser PE, □80 Micro PE, □70 Micro PE■50 Micro PE"
	}

}*/

func CheckContainer(n []interface{}) string {

	fmt.Println("CheckContainer")
	temp_ans := ""
	for _, v := range n {
		valStr := fmt.Sprint(v)
		if strings.Contains(valStr, "中性") {
			temp_ans = "■中性包裝 □標準外銷包裝"
		} else {
			temp_ans = "□中性包裝 ■標準外銷包裝 "
		}
	}
	return temp_ans
}

func CheckMicroLaser(n []interface{}) string {

	fmt.Println("CheckContainer")
	temp_ans := ""
	for _, v := range n {
		valStr := fmt.Sprint(v)
		if strings.Contains(valStr, "100") {
			temp_ans = "■100 Micro Laser PE, □80 Micro PE, □70 Micro PE□50 Micro PE"
		} else if strings.Contains(valStr, "80") {
			temp_ans = "□100 Micro Laser PE, ■80 Micro PE, □70 Micro PE□50 Micro PE"
		} else if strings.Contains(valStr, "70") {
			temp_ans = "□100 Micro Laser PE, □80 Micro PE, ■70 Micro PE□50 Micro PE"
		} else {
			temp_ans = "□100 Micro Laser PE, □80 Micro PE, □70 Micro PE■50 Micro PE"
		}
	}
	return temp_ans
}

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

	res, err := http.Get("https://api.testing.eirc.app/meglobe/v1.0/purchase-order/843fb5ca-fe4b-4d61-bb28-711ef64212b0")
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
	var readPoContent poJson
	json.Unmarshal(body, &readPoContent)
	//開始的表格 寫死內容
	f.SetCellValue("PO", "B3", readPoContent.Body.Order.Tel)
	f.SetCellValue("PO", "C7", readPoContent.Body.Attention)
	f.SetCellValue("PO", "C8", readPoContent.Body.OrdFrom)
	f.SetCellValue("PO", "H7", readPoContent.Body.ExpectDate)
	f.SetCellValue("PO", "H8", readPoContent.Body.Order.ContractID)
	f.SetCellValue("PO", "H9", readPoContent.Body.OrdNum)
	f.SetCellFormula("PO", "C9", "=H7+45")

	//插入row數
	poItemLength := len(readPoContent.Body.PoItems)
	theInsertRow := 0
	if poItemLength-4 > 0 {
		theInsertRow = poItemLength - 4
		for i := 0; i < theInsertRow; i++ {
			f.DuplicateRow("PO", 13)
		}

	}
	// 1.襯紙~15. 特殊要求位置設定
	var C18ToC33 [15]string
	for i := 0; i < 15; i++ {
		C18ToC33[i], _ = excelize.CoordinatesToCellName(3, 18+theInsertRow)
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
		f.SetCellValue("PO", IronOfArray[i], readPoContent.Body.PoItems[i].Grade)
	}

	//S/M設定
	var SMOfArray [10]string
	for i := 0; i < poItemLength; i++ {
		SMOfArray[i], _ = excelize.CoordinatesToCellName(3, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", SMOfArray[i], readPoContent.Body.PoItems[i].Edge)
	}

	//Size設定
	var sizeArray [10]string
	for i := 0; i < poItemLength; i++ {
		sizeArray[i], _ = excelize.CoordinatesToCellName(4, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", sizeArray[i], readPoContent.Body.PoItems[i].Size)
	}

	//訂單重量設定
	var weightOfArray [10]string
	for i := 0; i < poItemLength; i++ {
		weightOfArray[i], _ = excelize.CoordinatesToCellName(6, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", weightOfArray[i], readPoContent.Body.PoItems[i].Quantity)
	}

	//FOB設定
	var FOBOfArray [10]string
	for i := 0; i < poItemLength; i++ {
		FOBOfArray[i], _ = excelize.CoordinatesToCellName(7, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", FOBOfArray[i], readPoContent.Body.PoItems[i].FobFoshan)
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
		f.SetCellValue("PO", productPriceOfArray[i], readPoContent.Body.PoItems[i].FinishedPrice)
	}

	//備註
	var RemarkOfArray [10]string
	for i := 0; i < poItemLength; i++ {
		RemarkOfArray[i], _ = excelize.CoordinatesToCellName(10, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", RemarkOfArray[i], readPoContent.Body.PoItems[i].Remark)
	}

	//訂單重量(MT) total
	f.SetCellFormula("PO", "F16", "=SUM(F12:F15)")
	//總額(USD) total
	f.SetCellFormula("PO", "H16", "=SUM(H12:H15)")
	//1. 襯紙:
	backing_paper := CheckHeavy(readPoContent.Body.BackingPaper) + "; " + CheckBuy(readPoContent.Body.BackingPaper)
	f.SetCellValue("PO", C18ToC33[0], backing_paper)

	//存檔
	if err := f.SaveAs(outputName + ".xlsx"); err != nil {
		fmt.Println(err)
	}
	return outputName + ".xlsx"
}
