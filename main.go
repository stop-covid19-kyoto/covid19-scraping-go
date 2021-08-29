package main

import (
	"bytes"
	"encoding/base64"
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/gmail/v1"
)

func main() {
	// 引数をパースする
	flag.Parse()
	// パースした引数をargs変数に格納する
	args := flag.Args()
	// メールボックスを取得
	fmt.Println("メールボックスを取得しています")
	spreadsheet, lastUpdateTime, err := getSpreadsheetFromGmail(args)
	errorHandling(err)
	// 最終更新日時を変換
	fmt.Println("最終更新日時を変換しています")
	lastUpdate := getLastUpdate(lastUpdateTime)
	// 感染者の属性を変換
	fmt.Println("感染者の属性を変換しています")
	patients := getPatients(*spreadsheet, lastUpdateTime)
	// 感染者のサマリーを変換
	fmt.Println("感染者のサマリーを変換しています")
	patientsSummary := getPatientsSummary(patients, lastUpdateTime)
	// ニュースを変換
	fmt.Println("ニュースを変換しています")
	news := getNews(*spreadsheet)
	// PCR検査件数のサマリーを変換
	fmt.Println("PCR検査件数のサマリーを変換しています")
	inspectionsSummary := getInspectionsSummary(*spreadsheet, lastUpdateTime)
	// 検査陽性者の情報を変換
	fmt.Println("検査陽性者の情報を変換しています")
	mainSummary := getMainSummary(*spreadsheet, lastUpdateTime)

	exportDatas(lastUpdate, patients, patientsSummary, news, inspectionsSummary, mainSummary)

	// fmt.Println(spreadsheet.GetCellValue("陽性者の属性", "A1"))

}

func getLastUpdate(lastUpdateTime time.Time) LastUpdate {
	var lastUpdate LastUpdate
	lastUpdate.LastUpdate = lastUpdateTime.Format("2006/01/02 15:04")
	return lastUpdate
}

func getSpreadsheetFromGmail(args []string) (*excelize.File, time.Time, error) {
	var attachmentId string

	config := oauth2.Config{
		ClientID:     args[0],
		ClientSecret: args[1],
		Endpoint:     google.Endpoint,
		RedirectURL:  "urn:ietf:wg:oauth:2.0:oob",
		Scopes:       []string{"https://mail.google.com/"},
	}

	token := oauth2.Token{
		TokenType:    "Bearer",
		RefreshToken: args[2],
	}

	client := config.Client(oauth2.NoContext, &token)

	fmt.Println("メールボックスを開きます")
	server, err := gmail.New(client)
	errorHandling(err)

	messages, err := server.Users.Messages.List("me").MaxResults(100).Q("NOT in:draft has:attachment from:" + args[3] + " OR from:" + args[4]).Do()
	errorHandling(err)

	if len(messages.Messages) <= 0 {
		return nil, time.Now(), errors.New("MessagesNotFound")
	}

	message := messages.Messages[0]

	messageData, err := server.Users.Messages.Get("me", message.Id).Format("full").Do()
	errorHandling(err)

	headers := messageData.Payload.Headers

	for _, header := range headers {
		if header.Name == "Subject" {
			break
		}
	}

	for _, part := range messageData.Payload.Parts {
		// if strings.LastIndex()
		fileName := string([]rune(part.Filename))
		if strings.HasSuffix(fileName, "data.xlsx") {
			attachmentId = part.Body.AttachmentId
			break
		}
	}

	attachment, err := server.Users.Messages.Attachments.Get("me", message.Id, attachmentId).Do()
	errorHandling(err)

	var receivedTime = time.Date(0, 0, 0, 0, 0, 0, 0, time.Local)

	for _, header := range headers {
		if header.Name == "Date" {
			receivedTime, err = time.Parse("Mon, 02 Jan 2006 15:04:05 -0700", header.Value)
			errorHandling(err)
			break
		}
	}

	// fmt.Println(message.Payload)

	spreadsheetData, err := base64.URLEncoding.DecodeString(attachment.Data)
	errorHandling(err)

	spreadsheet, err := excelize.OpenReader(bytes.NewBuffer(spreadsheetData))

	return spreadsheet, receivedTime, err
}

func getPatients(spreadsheet excelize.File, lastUpdateTime time.Time) Patients {

	rows, err := spreadsheet.GetRows("陽性者の属性")
	errorHandling(err)

	rowLength := len(rows)

	patients := Patients{}
	patientsSlice := make([]Patient, 0)

	for i := 1; i <= rowLength; i++ {
		patient := Patient{}

		for j := 1; j <= 6; j++ {
			cellPos, _ := excelize.CoordinatesToCellName(j, i)
			// fmt.Println(cellPos)
			cellValue, _ := spreadsheet.GetCellValue("陽性者の属性", cellPos)
			switch j {
			case 1:
				noStr := strings.Replace(cellValue, "例目", "", 1)
				no, _ := strconv.Atoi(noStr)
				patient.No = no
				break
			case 2:
				date := convertJpEraToDate(cellValue)
				patient.ReleasedDate = convertDateToString(date)
				patient.Date = date.Format("2006-01-02")
				break
			case 3:
				genderPos, _ := excelize.CoordinatesToCellName(j+1, i)
				gender, _ := spreadsheet.GetCellValue("陽性者の属性", genderPos)
				if strings.HasSuffix(cellValue, "未満") {
					patient.GenAndGender = strings.Replace(cellValue, "未満", "歳未満", 1) + gender
				} else if strings.HasSuffix(cellValue, "以上") {
					patient.GenAndGender = strings.Replace(cellValue, "以上", "歳以上", 1) + gender
				} else if cellValue == "-" {
					if gender == "-" {
						patient.GenAndGender = ""
					} else {
						patient.GenAndGender = gender
					}
				} else {
					patient.GenAndGender = cellValue + "代" + gender
				}
				break
			case 5:
				patient.Living = cellValue
				break
			case 6:
				patient.Leaved = cellValue
				break
			}
		}

		// fmt.Println(patient.Date)

		patientsSlice = append(patientsSlice, patient)
	}

	patients.LastUpdate = lastUpdateTime.Format("2006/01/02 15:04")

	patientsSliceReversed := []Patient{}

	for i := len(patientsSlice) - 1; i >= 0; i-- {
		patientsSliceReversed = append(patientsSliceReversed, patientsSlice[i])
	}

	patients.Data = patientsSliceReversed

	return patients
}

func getPatientsSummary(patients Patients, lastUpdateTime time.Time) PatientsSummary {
	patientsCount := 0
	patientsSummary := PatientsSummary{}

	patientsSummaryData := []Summary{}

	recentPatientDate := time.Date(0, 0, 0, 0, 0, 0, 0, time.Local)

	for i := 0; i < len(patients.Data); i++ {
		// for i := 0; i < 10; i++ {

		date, _ := time.Parse("2006-01-02", patients.Data[i].Date)

		patientsCount++

		if i != 0 {
			zeroDays := int(date.Sub(recentPatientDate).Hours()) / 24
			if zeroDays >= 1 {
				for j := 1; j < zeroDays; j++ {
					zeroDate := recentPatientDate.AddDate(0, 0, j)
					patientsSummaryData = append(patientsSummaryData, Summary{zeroDate.Format("2006-01-02T15:04:00") + ".000Z", 0})
					// fmt.Println(zeroDate)
				}
			}
		}

		if recentPatientDate != date || i == 0 {
			recentPatientDate = date
			summary := Summary{patients.Data[i].ReleasedDate, patientsCount}
			patientsSummaryData = append(patientsSummaryData, summary)
			patientsCount = 0
			continue
		}

		if i+1 == len(patients.Data) {
			patientsCount++
			recentPatientDate = date
			summary := Summary{patients.Data[i].ReleasedDate, patientsCount}
			patientsSummaryData = append(patientsSummaryData, summary)
			patientsCount = 0
		}

		recentPatientDate = date
	}

	patientsSummary.Data = patientsSummaryData
	patientsSummary.LastUpdate = lastUpdateTime.Format("2006/01/02 15:04")

	return patientsSummary
}

func getMainSummary(spreadsheet excelize.File, lastUpdateTime time.Time) MainSummary {
	rows, err := spreadsheet.GetRows("PCR検査件数")
	errorHandling(err)

	rowLength := len(rows)

	mainSummary := MainSummary{}

	mainSummaryPatients := MainSummaryPatients{}

	for i := 0; i < 8; i++ {
		mainSummaryPatients.Children = append(mainSummaryPatients.Children, SummarySection{})
	}

	// 陽性患者数
	for j := 1; j <= rowLength; j++ {
		cellPos, _ := excelize.CoordinatesToCellName(3, j)
		cellValue, _ := spreadsheet.GetCellValue("PCR検査件数", cellPos)
		if cellValue != "" {
			mainSummaryPatients.Attr = "陽性患者数"
			value, _ := strconv.Atoi(cellValue)
			mainSummaryPatients.Value = value
			break
		}
	}

	// 入院患者数
	for j := 1; j <= rowLength; j++ {
		cellPos, _ := excelize.CoordinatesToCellName(5, j)
		cellValue, _ := spreadsheet.GetCellValue("PCR検査件数", cellPos)
		if cellValue != "" {
			mainSummaryPatients.Children[0].Attr = "入院中・入院調整中"
			value, _ := strconv.Atoi(cellValue)
			mainSummaryPatients.Children[0].Value = value
			break
		}
	}

	// 高度重症病床
	for j := 1; j <= rowLength; j++ {
		cellPos, _ := excelize.CoordinatesToCellName(6, j)
		cellValue, _ := spreadsheet.GetCellValue("PCR検査件数", cellPos)
		if cellValue != "" {
			mainSummaryPatients.Children[1].Attr = "高度重症病床"
			value, _ := strconv.Atoi(cellValue)
			mainSummaryPatients.Children[1].Value = value
			break
		}
	}

	// その他
	for j := 1; j <= rowLength; j++ {
		cellPos, _ := excelize.CoordinatesToCellName(7, j)
		cellValue, _ := spreadsheet.GetCellValue("PCR検査件数", cellPos)
		if cellValue != "" {
			mainSummaryPatients.Children[2].Attr = "その他"
			value, _ := strconv.Atoi(cellValue)
			mainSummaryPatients.Children[2].Value = value
			break
		}
	}

	// 宿泊施設
	for j := 1; j <= rowLength; j++ {
		cellPos, _ := excelize.CoordinatesToCellName(8, j)
		cellValue, _ := spreadsheet.GetCellValue("PCR検査件数", cellPos)
		if cellValue != "" {
			mainSummaryPatients.Children[3].Attr = "宿泊施設"
			value, _ := strconv.Atoi(cellValue)
			mainSummaryPatients.Children[3].Value = value
			break
		}
	}

	// 自宅療養
	for j := 1; j <= rowLength; j++ {
		cellPos, _ := excelize.CoordinatesToCellName(9, j)
		cellValue, _ := spreadsheet.GetCellValue("PCR検査件数", cellPos)
		if cellValue != "" {
			mainSummaryPatients.Children[4].Attr = "自宅療養"
			value, _ := strconv.Atoi(cellValue)
			mainSummaryPatients.Children[4].Value = value
			break
		}
	}

	// 死亡
	for j := 1; j <= rowLength; j++ {
		cellPos, _ := excelize.CoordinatesToCellName(10, j)
		cellValue, _ := spreadsheet.GetCellValue("PCR検査件数", cellPos)
		if cellValue != "" {
			mainSummaryPatients.Children[5].Attr = "死亡"
			value, _ := strconv.Atoi(cellValue)
			mainSummaryPatients.Children[5].Value = value
			break
		}
	}

	// 退院
	for j := 1; j <= rowLength; j++ {
		cellPos, _ := excelize.CoordinatesToCellName(4, j)
		cellValue, _ := spreadsheet.GetCellValue("PCR検査件数", cellPos)
		if cellValue != "" {
			mainSummaryPatients.Children[6].Attr = "退院・解除"
			value, _ := strconv.Atoi(cellValue)
			mainSummaryPatients.Children[6].Value = value
			break
		}
	}

	// 調整中
	for j := 1; j <= rowLength; j++ {
		cellPos, _ := excelize.CoordinatesToCellName(11, j)
		cellValue, _ := spreadsheet.GetCellValue("PCR検査件数", cellPos)
		if cellValue != "" {
			mainSummaryPatients.Children[7].Attr = "調整中"
			value, _ := strconv.Atoi(cellValue)
			mainSummaryPatients.Children[7].Value = value
			break
		}
	}

	inspectionsStr, _ := spreadsheet.GetCellValue("PCR検査件数", "B2")
	mainSummary.Attr = "検査実施人数"
	mainSummary.Value, _ = strconv.Atoi(inspectionsStr)

	mainSummary.Children = append(mainSummary.Children, mainSummaryPatients)

	mainSummary.LastUpdate = lastUpdateTime.Format("2006/01/02 15:04")

	return mainSummary
}

func getInspectionsSummary(spreadsheet excelize.File, lastUpdateTime time.Time) InspectionsSummary {
	rows, err := spreadsheet.GetRows("PCR検査件数")
	errorHandling(err)

	rowLength := len(rows)

	inspectionsSummary := InspectionsSummary{}

	inspectionsSummaryData := []Summary{}

	for i := 1; i <= rowLength; i++ {
		inspections := Summary{}
		for j := 1; j <= rowLength; j++ {
			cellPos, _ := excelize.CoordinatesToCellName(j, i)
			// fmt.Println(cellPos)
			cellValue, _ := spreadsheet.GetCellValue("PCR検査件数", cellPos)

			switch j {
			case 1:
				inspections.Date = cellValue
				break
			case 2:
				if i == rowLength {
					subtotal, _ := strconv.Atoi(cellValue)
					inspections.SubTotal = subtotal
					break
				} else {
					recentDateCellPos, _ := excelize.CoordinatesToCellName(j, i+1)
					recentDateCellValue, _ := spreadsheet.GetCellValue("PCR検査件数", recentDateCellPos)
					subtotal, _ := strconv.Atoi(cellValue)
					recentSubtotal, _ := strconv.Atoi(recentDateCellValue)
					inspections.SubTotal = subtotal - recentSubtotal
					break
				}
			}
		}
		inspectionsSummaryData = append(inspectionsSummaryData, inspections)
	}

	inspectionsSummaryDataReversed := []Summary{}

	for i := len(inspectionsSummaryData) - 1; i >= 0; i-- {
		inspectionsSummaryDataReversed = append(inspectionsSummaryDataReversed, inspectionsSummaryData[i])
	}

	inspectionsSummary.Data = inspectionsSummaryDataReversed
	inspectionsSummary.LastUpdate = lastUpdateTime.Format("2006/01/02 15:04")

	return inspectionsSummary
}

func getNews(spreadsheet excelize.File) News {

	rows, err := spreadsheet.GetRows("最新の情報")
	errorHandling(err)

	rowLength := len(rows)

	news := News{}
	newsItemSlice := make([]NewsItem, 0)

	for i := 1; i <= rowLength; i++ {
		newsItem := NewsItem{}
		for j := 1; j <= rowLength; j++ {
			cellPos, _ := excelize.CoordinatesToCellName(j, i)
			// fmt.Println(cellPos)
			cellValue, _ := spreadsheet.GetCellValue("最新の情報", cellPos)

			switch j {
			case 1:
				newsItem.Date = cellValue
				break
			case 2:
				newsItem.Text = cellValue
				break
			case 3:
				newsItem.Url = cellValue
				break
			}
		}
		newsItemSlice = append(newsItemSlice, newsItem)
	}

	news.NewsItems = newsItemSlice

	return news
}

func exportDatas(lastUpdate LastUpdate, patients Patients, patientsSummary PatientsSummary, news News, inspectionsSummary InspectionsSummary, mainSummary MainSummary) {
	lastUpdateBytes, err := json.Marshal(lastUpdate)
	errorHandling(err)
	lastUpdateIndented := new(bytes.Buffer)
	json.Indent(lastUpdateIndented, lastUpdateBytes, "", "    ")

	patientsBytes, err := json.Marshal(patients)
	errorHandling(err)
	patientsIndented := new(bytes.Buffer)
	json.Indent(patientsIndented, patientsBytes, "", "    ")

	patientsSummaryBytes, err := json.Marshal(patientsSummary)
	errorHandling(err)
	patientsSummaryIndented := new(bytes.Buffer)
	json.Indent(patientsSummaryIndented, patientsSummaryBytes, "", "    ")

	newsBytes, err := json.Marshal(news)
	errorHandling(err)
	newsIndented := new(bytes.Buffer)
	json.Indent(newsIndented, newsBytes, "", "    ")

	inspectionsSummaryBytes, err := json.Marshal(inspectionsSummary)
	errorHandling(err)
	inspectionsSummaryIndented := new(bytes.Buffer)
	json.Indent(inspectionsSummaryIndented, inspectionsSummaryBytes, "", "    ")

	mainSummaryBytes, err := json.Marshal(mainSummary)
	errorHandling(err)
	mainSummaryIndented := new(bytes.Buffer)
	json.Indent(mainSummaryIndented, mainSummaryBytes, "", "    ")

	lastUpdateFile, err := os.Create("./data/last_update.json")
	defer lastUpdateFile.Close()
	lastUpdateFile.Write(lastUpdateIndented.Bytes())

	patientsFile, err := os.Create("./data/patients.json")
	defer patientsFile.Close()
	patientsFile.Write(patientsIndented.Bytes())

	patientsSummaryFile, err := os.Create("./data/patients_summary.json")
	defer patientsSummaryFile.Close()
	patientsSummaryFile.Write(patientsSummaryIndented.Bytes())

	newsFile, err := os.Create("./data/news.json")
	defer newsFile.Close()
	newsFile.Write(newsIndented.Bytes())

	inspectionsSummaryFile, err := os.Create("./data/inspections_summary.json")
	defer inspectionsSummaryFile.Close()
	inspectionsSummaryFile.Write(inspectionsSummaryIndented.Bytes())

	mainSummaryFile, err := os.Create("./data/main_summary.json")
	defer mainSummaryFile.Close()
	mainSummaryFile.Write(mainSummaryIndented.Bytes())

	/*
		patients, err := json.Marshal(Patients{})
		errorHandling(err)

		patientsSummaryBytes, err := json.Marshal(PatientsSummary{})
		errorHandling(err)

		newsBytes, err := json.Marshal(News{})
		errorHandling(err)

		mainSummary, err := json.Marshal(MainSummary{})
		errorHandling(err)
	*/

}

func convertDateToString(date time.Time) string {
	return date.Format("2006-01-02T15:04:00") + ".000Z"
}

func convertJpEraToDate(date string) time.Time {
	var dateArray [3]int = [3]int{0, 0, 0}

	j := 0

	for i := 0; i < 3; i++ {
		numStr := ""
		k := 0
		for {
			// fmt.Println(j)
			if k > 5 {
				// fmt.Println("break")
				j++
				break
			} else if (string([]rune(date)[j]) == "年") || (string([]rune(date)[j]) == "月") || (string([]rune(date)[j]) == "日") {
				// fmt.Println("break" + string([]rune(date)[j]))
				j++
				break
			} else if (string([]rune(date)[j]) == "元") && (string([]rune(date)[j+1]) == "年") {
				// fmt.Println("break" + string([]rune(date)[j]))
				numStr = "1"
				j += 2
				break
			}
			if isNumber(string([]rune(date)[j])) {
				// fmt.Println(string([]rune(date)[j]))
				numStr += string([]rune(date)[j])
			}
			k++
			j++
		}
		dateArray[i], _ = strconv.Atoi(numStr)
	}

	switch string([]rune(date)[:2]) {
	case "昭和":
		dateArray[0] += 1925
		break
	case "平成":
		dateArray[0] += 1988
		break
	case "令和":
		dateArray[0] += 2018
		break
	}

	return time.Date(dateArray[0], time.Month(dateArray[1]), dateArray[2], 0, 0, 0, 0, time.Local)

}

func isNumber(str string) bool {
	for i := 0; i <= 9; i++ {
		if strconv.Itoa(i) == str {
			//fmt.Println("b")
			return true
		}
	}
	return false
}

func errorHandling(err error) {
	if err != nil {
		// fmt.Println(err)
		os.Exit(-1)
	}
}

type LastUpdate struct {
	LastUpdate string `json:"last_update"`
}

type Patient struct {
	No           int    `json:"No"`
	ReleasedDate string `json:"リリース日"`
	Living       string `json:"居住地"`
	GenAndGender string `json:"年代と性別"`
	Leaved       string `json:"退院"`
	Date         string `json:"date"`
}

type Patients struct {
	Data       []Patient `json:"data"`
	LastUpdate string    `json:"last_update"`
}

type Summary struct {
	Date     string `json:"日付"`
	SubTotal int    `json:"小計"`
}

type PatientsSummary struct {
	Data       []Summary `json:"data"`
	LastUpdate string    `json:"last_update"`
}

type NewsItem struct {
	Date string `json:"date"`
	Text string `json:"text"`
	Url  string `json:"url"`
}

type News struct {
	NewsItems []NewsItem `json:"newsItems"`
}

type SummarySection struct {
	Attr  string `json:"attr"`
	Value int    `json:"value"`
}

type MainSummaryPatients struct {
	Attr     string           `json:"attr"`
	Value    int              `json:"value"`
	Children []SummarySection `json:"children"`
}

type MainSummary struct {
	Attr       string                `json:"attr"`
	Value      int                   `json:"value"`
	Children   []MainSummaryPatients `json:"children"`
	LastUpdate string                `json:"last_update"`
}

type InspectionsSummary struct {
	Data       []Summary `json:"data"`
	LastUpdate string    `json:"last_update"`
}
