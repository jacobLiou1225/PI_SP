package pkg

import (
	"encoding/json"
	"fmt"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io/ioutil"
	"net/http"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func countPoLine(i string) int {
	return (len(i) / 55)
}

func CheckBuy(n []interface{}) string {
	temp_ans := ""
	for _, v := range n {
		valStr := fmt.Sprint(v)
		if strings.Contains(valStr, "不") {
			temp_ans = "□需要■不需要"
		} else if strings.Contains(valStr, "需要") {
			temp_ans = "■需要□不需要"
		} else {
			temp_ans = "□需要□不需要"
		}
	}
	return temp_ans
}
func CheckHeavy(n []interface{}) string {
	temp_ans := ""
	for _, v := range n {
		valStr := fmt.Sprint(v)
		if strings.Contains(valStr, "不") {
			temp_ans = "□扣重■不扣重"
		} else if strings.Contains(valStr, "扣重") {
			temp_ans = "■扣重□不扣重"
		} else {
			temp_ans = "□扣重□不扣重"
		}
	}
	return temp_ans

}

func CheckDirection(n []interface{}) string {
	temp_ans := ""
	for _, v := range n {
		valStr := fmt.Sprint(v)
		if strings.Contains(valStr, "左右") {
			temp_ans = "□上下■左右"
		} else {
			temp_ans = "■上下□左右"
		}
	}
	return temp_ans

}

func CheckPoliFilm(n []interface{}) string {
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
func CheckDiameter(n []interface{}) string {
	temp_ans := ""
	for _, v := range n {
		valStr := fmt.Sprint(v)
		if strings.Contains(valStr, "508") {
			temp_ans = "□300mm ■508mm □610mm"
		} else if strings.Contains(valStr, "300") {
			temp_ans = "■300mm □508mm □610mm"
		} else if strings.Contains(valStr, "610") {
			temp_ans = "□300mm □508mm ■610mm"
		}
	}
	return temp_ans
}

func CheckNational(n []interface{}) string {
	temp_ans := ""
	for _, v := range n {
		valStr := fmt.Sprint(v)
		if strings.Contains(valStr, "國產") {
			temp_ans = "■國產                          □進口"
		} else {
			temp_ans = "□國產                          ■進口"
		}
	}
	return temp_ans
}
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

func CheckWeight(n []interface{}) string {

	temp_ans := ""
	for _, v := range n {
		valStr := fmt.Sprint(v)
		if strings.Contains(valStr, "淨重") {
			temp_ans = " ■淨重□毛重"
		} else {
			temp_ans = " □淨重■毛重 "
		}
	}
	return temp_ans
}

type poJson struct {
	Code      int       `json:"code"`
	Timestamp time.Time `json:"timestamp"`
	Body      struct {
		PurID                     string      `json:"pur_id"`
		OID                       string      `json:"o_id"`
		OrdNum                    string      `json:"ord_num"`
		Attention                 string      `json:"attention"`
		OrdFrom                   string      `json:"ord_from"`
		PoDate                    string      `json:"po_date"`
		ExpectDate                string      `json:"expect_date"`
		WeightTolerance           string      `json:"weight_tolerance"`
		ThicknessTolerance        string      `json:"thickness_tolerance"`
		TermsOfTrade              string      `json:"terms_of_trade"`
		DischargePort             string      `json:"discharge_port"`
		MarkFormat                string      `json:"mark_format"`
		Requirements              string      `json:"requirements"`
		Remark                    string      `json:"remark"`
		DelTerm                   string      `json:"del_term"`
		BonitaCaseID              string      `json:"bonita_case_id"`
		CreatedAt                 time.Time   `json:"created_at"`
		CreatedBy                 string      `json:"created_by"`
		UpdatedAt                 time.Time   `json:"updated_at"`
		UpdatedBy                 interface{} `json:"updated_by"`
		PlateWeight               float64     `json:"plate_weight"` ///////////////////////////////
		CoilWeight                float64     `json:"coil_weight"`  ///////////////////////////////
		AllAssigned               bool        `json:"all_assigned"`
		TotalQuantity             float64     `json:"total_quantity"`
		TotalAmount               float64     `json:"total_amount"`
		TotalOriginalCoilQuantity float64     `json:"total_original_coil_quantity"`
		Order                     struct {
			OID         string `json:"o_id"`
			ContractID  string `json:"contract_id"`
			Attention   string `json:"attention"`
			Tel         string `json:"tel"`
			Address     string `json:"address"`
			OrdDate     string `json:"ord_date"`
			Description string `json:"description"`
		} `json:"Order"`
		PoItems []struct {
			PoItemID string `json:"po_item_id"`
			ItemNum  int    `json:"item_num"`
			Grade    string `json:"grade"`
			Edge     string `json:"edge"`

			Size          string  `json:"size"`
			Quantity      float64 `json:"quantity"`
			UnitPrice     float64 `json:"unit_price"`
			FinishedPrice float64 `json:"finished_price"`
			Amount        float64 `json:"amount"`
			Remark        string  `json:"remark"`
			IsAssigned    bool    `json:"is_assigned"`
			IsPurchased   bool    `json:"is_purchased"`
			IsAssignedCI  bool    `json:"is_assignedCI"`
			PiItemID      string  `json:"pi_item_id"`
			PiItem        struct {
				PiItemID  string  `json:"pi_item_id"`
				OID       string  `json:"o_id"`
				ItemNum   int     `json:"item_num"`
				Grade     string  `json:"grade"`
				Edge      string  `json:"edge"`
				Size      string  `json:"size"`
				Quantity  float64 `json:"quantity"`
				UnitPrice float64 `json:"unit_price"`
			} `json:"PiItem"`
		} `json:"PoItems"`
		OriginalCoil []struct {
			OriCoilID      string  `json:"ori_coil_id"`
			ItemNum        int     `json:"item_num"`
			Grade          string  `json:"grade"`
			SteelPlantName string  `json:"steel_plant_name"`
			Size           string  `json:"size"`
			Quantity       float64 `json:"quantity"`
		} `json:"OriginalCoil"`
		BackingPaper []interface {
		} `json:"backing_paper"`
		Sticker []interface {
		} `json:"sticker"`
		Packing []interface {
		} `json:"packing"`
		WeightPerPiece []interface {
		} `json:"weight_per_piece"`
		SprayWord []interface {
		} `json:"spray_word"`
		Diameter []interface {
		} `json:"diameter"`
		Pallet []interface {
		} `json:"pallet"`
		ShippingMark        []interface{} `json:"shipping_mark"`
		DirectionOfEntrance []interface {
		} `json:"direction_of_entrance"`
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
	pictureMoveRow := 0
	poItemLength := len(readPoContent.Body.PoItems)
	theInsertRow := 0
	if poItemLength-4 > 0 {
		theInsertRow = poItemLength - 4
		for i := 0; i < theInsertRow; i++ {
			f.DuplicateRow("PO", 13)
		}

	}
	pictureMoveRow += theInsertRow
	// 1.襯紙~15. 特殊要求位置設定
	var C18ToC33 [16]string
	for i := 0; i < 16; i++ {
		C18ToC33[i], _ = excelize.CoordinatesToCellName(3, 18+theInsertRow+i)
	}

	// OriginalCoil---grade
	originalCoilGrade := make([]string, len(readPoContent.Body.OriginalCoil))
	for i := range originalCoilGrade {
		originalCoilGrade[i], _ = excelize.CoordinatesToCellName(8, 27+theInsertRow+i*3)
	}

	// OriginalCoil---size
	originalCoilSize := make([]string, len(readPoContent.Body.OriginalCoil))
	for i := range originalCoilSize {
		originalCoilSize[i], _ = excelize.CoordinatesToCellName(8, 28+theInsertRow+i*3)
	}
	// OriginalCoil---quantity
	originalCoilQuantity := make([]string, len(readPoContent.Body.OriginalCoil))
	for i := range originalCoilQuantity {
		originalCoilQuantity[i], _ = excelize.CoordinatesToCellName(9, 28+theInsertRow+i*3)
	}
	//I37公式設定
	i28Position, _ := excelize.CoordinatesToCellName(9, 28+theInsertRow)
	i36Position, _ := excelize.CoordinatesToCellName(9, 36+theInsertRow)
	i37Position, _ := excelize.CoordinatesToCellName(9, 37+theInsertRow)

	//編號設定
	NumerOfArray := make([]string, poItemLength)
	for i := range NumerOfArray {
		NumerOfArray[i], _ = excelize.CoordinatesToCellName(1, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", NumerOfArray[i], readPoContent.Body.PoItems[i].ItemNum)
	}

	//鋼種設定
	IronOfArray := make([]string, poItemLength)
	for i := range IronOfArray {
		IronOfArray[i], _ = excelize.CoordinatesToCellName(2, 12+i)
	}

	for i := range IronOfArray {
		f.SetCellValue("PO", IronOfArray[i], readPoContent.Body.PoItems[i].Grade)
	}
	//S/M設定
	SMOfArray := make([]string, poItemLength)
	for i := range SMOfArray {
		SMOfArray[i], _ = excelize.CoordinatesToCellName(3, 12+i)
	}

	for i := range SMOfArray {
		f.SetCellValue("PO", SMOfArray[i], readPoContent.Body.PoItems[i].Edge)
	}

	//Size設定

	sizeArray := make([]string, poItemLength)
	for i := range sizeArray {
		sizeArray[i], _ = excelize.CoordinatesToCellName(4, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", sizeArray[i], readPoContent.Body.PoItems[i].Size)
	}

	//訂單重量設定

	weightOfArray := make([]string, poItemLength)
	for i := range weightOfArray {
		weightOfArray[i], _ = excelize.CoordinatesToCellName(6, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", weightOfArray[i], readPoContent.Body.PoItems[i].Quantity)
	}

	//FOB設定

	FOBOfArray := make([]string, poItemLength)
	for i := range FOBOfArray {
		FOBOfArray[i], _ = excelize.CoordinatesToCellName(7, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", FOBOfArray[i], readPoContent.Body.PoItems[i].UnitPrice)
	}
	//總額(USD)
	H12TOHXX := make([]string, poItemLength)
	for i := 0; i < poItemLength; i++ {
		H12TOHXX[i], _ = excelize.CoordinatesToCellName(8, 12+i)
	}
	F12TOFXX := make([]string, poItemLength)
	for i := 0; i < poItemLength; i++ {
		F12TOFXX[i], _ = excelize.CoordinatesToCellName(6, 12+i)
	}

	G12TOGXX := make([]string, poItemLength)
	for i := 0; i < poItemLength; i++ {
		G12TOGXX[i], _ = excelize.CoordinatesToCellName(7, 12+i)
	}

	//成品價/基價
	productPriceOfArray := make([]string, poItemLength)
	for i := 0; i < poItemLength; i++ {
		productPriceOfArray[i], _ = excelize.CoordinatesToCellName(9, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", productPriceOfArray[i], readPoContent.Body.PoItems[i].FinishedPrice)
	}

	//備註

	RemarkOfArray := make([]string, poItemLength)
	for i := 0; i < poItemLength; i++ {
		RemarkOfArray[i], _ = excelize.CoordinatesToCellName(10, 12+i)
	}

	for i := 0; i < poItemLength; i++ {
		f.SetCellValue("PO", RemarkOfArray[i], readPoContent.Body.PoItems[i].Remark)
	}
	//總額(USD)
	for i := 0; i < poItemLength; i++ {
		f.SetCellFormula("PO", H12TOHXX[i], "="+F12TOFXX[i]+"*"+G12TOGXX[i]+"")
	}
	//訂單重量(MT) total
	f.SetCellFormula("PO", "F"+strconv.Itoa(16+theInsertRow)+"", "=SUM(F12:F"+strconv.Itoa(15+theInsertRow)+")")
	//總額(USD) total
	f.SetCellFormula("PO", "H"+strconv.Itoa(16+theInsertRow)+"", "=SUM(H12:H"+strconv.Itoa(15+theInsertRow)+")")
	//1. 襯紙:
	backing_paper := CheckHeavy(readPoContent.Body.BackingPaper) + "; " + CheckBuy(readPoContent.Body.BackingPaper)
	f.SetCellValue("PO", C18ToC33[0], backing_paper)

	//2. 貼膜
	sticker := CheckNational(readPoContent.Body.Sticker) + "   " + CheckPoliFilm(readPoContent.Body.Sticker) + CheckNovacel(readPoContent.Body.Sticker)
	f.SetCellValue("PO", C18ToC33[1], sticker)
	sticker_blue_black := CheckBlueBlack(readPoContent.Body.Sticker)
	f.SetCellValue("PO", C18ToC33[2], sticker_blue_black)
	sticker_Micro_Laser_PE := CheckMicroLaser(readPoContent.Body.Sticker) + ";" + CheckHeavy(readPoContent.Body.Sticker)
	f.SetCellValue("PO", C18ToC33[3], sticker_Micro_Laser_PE)

	//3. 包裝
	packing := CheckContainer(readPoContent.Body.Packing)
	f.SetCellValue("PO", C18ToC33[4], packing)
	//4. 單件重量:
	weight_per_piece := "卷: Max. " + strconv.FormatFloat(readPoContent.Body.CoilWeight, 'f', 2, 64) + "MT 版: Max. " + strconv.FormatFloat(readPoContent.Body.PlateWeight, 'f', 2, 64) + "MT" + " " + CheckWeight(readPoContent.Body.WeightPerPiece)
	f.SetCellValue("PO", C18ToC33[5], weight_per_piece)
	//5. 噴字:
	spray_word := CheckBuy(readPoContent.Body.SprayWord) + " (GRADE/ FINISH/ THEORETICAL THICKNESS/ WIDTH/ LENGTH/ COIL NO.)"
	f.SetCellValue("PO", C18ToC33[6], spray_word)

	//6. 內徑:
	diameter := CheckDiameter(readPoContent.Body.Diameter)
	f.SetCellValue("PO", C18ToC33[7], diameter)
	//7. 棧板:
	pallet := CheckBuy(readPoContent.Body.Pallet)
	f.SetCellValue("PO", C18ToC33[8], pallet)
	//8. 嘜頭:
	shipping_mark := readPoContent.Body.ShippingMark
	f.SetCellValue("PO", C18ToC33[9], shipping_mark)
	//9. 洞口方向:
	direction_of_entrance := CheckDirection(readPoContent.Body.DirectionOfEntrance)
	f.SetCellValue("PO", C18ToC33[10], direction_of_entrance)
	//10. 重量公差:
	weight_tolerance := readPoContent.Body.WeightTolerance
	f.SetCellValue("PO", C18ToC33[11], weight_tolerance)
	//11. 厚度公差:
	thickness_tolerance := readPoContent.Body.ThicknessTolerance
	f.SetCellValue("PO", C18ToC33[12], thickness_tolerance)
	//12. 貿易條件:
	terms_of_trade := readPoContent.Body.TermsOfTrade
	f.SetCellValue("PO", C18ToC33[13], terms_of_trade)
	//13. 卸貨港:
	discharge_port := readPoContent.Body.DischargePort
	f.SetCellValue("PO", C18ToC33[14], discharge_port)
	//14. 麥頭格式:
	mark_format := readPoContent.Body.MarkFormat
	f.SetCellValue("PO", C18ToC33[15], mark_format)
	// OriginalCoil---grade 設值
	for i, n := range readPoContent.Body.OriginalCoil {
		f.SetCellValue("PO", originalCoilGrade[i], n.Grade)
	}
	// OriginalCoil---size 設值
	for i, n := range readPoContent.Body.OriginalCoil {
		f.SetCellValue("PO", originalCoilSize[i], n.Size)
	}
	// OriginalCoil---quantity 設值
	for i, n := range readPoContent.Body.OriginalCoil {
		f.SetCellValue("PO", originalCoilQuantity[i], n.Quantity)
	}

	f.SetCellFormula("PO", i37Position, "=SUM("+i28Position+":"+i36Position+")")

	str := readPoContent.Body.Remark
	RemarkSplitString := strings.Split(str, "\n")     // 1*客戶付款條件為10%訂金, 90% BL,、2客戶...、3客戶...、4客戶...
	RemarkSplitStringLength := len(RemarkSplitString) //4

	discribeStyle, _ := f.NewStyle(`{
		"alignment":{
			"horizontal":"left",
			"wrap_text":true,
			"vertical":"top"
		},
		"font": {
			"family": "Times New Roman"	,
			"size" : 16
		}
	}`)

	RemarkNewLineArray := make([]int, RemarkSplitStringLength) //每個註解需要用到幾行

	RemarkFirstPosition, _ := excelize.CoordinatesToCellName(0, 0)
	RemarkSecPosition, _ := excelize.CoordinatesToCellName(0, 0)

	coculate := 0

	for i := range RemarkNewLineArray {
		RemarkNewLineArray[i] = countPoLine(RemarkSplitString[i]) //每個註解需要用到幾行
		fmt.Println("第", i, "筆資料共有這麼多行=", RemarkNewLineArray[i])
		if RemarkNewLineArray[i] > 0 {

			coculate += RemarkNewLineArray[i]
			if i == 0 {
				f.InsertRow("PO", theInsertRow+36)
				RemarkFirstPosition, _ = excelize.CoordinatesToCellName(3, theInsertRow+35)
				RemarkSecPosition, _ = excelize.CoordinatesToCellName(3, theInsertRow+35+1)
			} else {
				f.InsertRow("PO", theInsertRow+36+i+coculate)
				RemarkFirstPosition, _ = excelize.CoordinatesToCellName(3, theInsertRow+36+coculate+i-1)
				RemarkSecPosition, _ = excelize.CoordinatesToCellName(3, theInsertRow+36+i+coculate)
			}

			f.MergeCell("PO", RemarkFirstPosition, RemarkSecPosition)
			f.SetCellValue("PO", RemarkFirstPosition, RemarkSplitString[i])
			f.SetCellStyle("PO", RemarkFirstPosition, RemarkSecPosition, discribeStyle)

		}
	}
	pictureMoveRow += coculate
	f.AddPicture("PO", "A"+strconv.Itoa(41+pictureMoveRow)+"", "Signature.jpg", `{
        "x_scale": 0.7        
    }`)

	//存檔
	if err := f.SaveAs(outputName + ".xlsx"); err != nil {
		fmt.Println(err)
	}
	return outputName + ".xlsx"
}
