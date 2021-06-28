package main

import (
	"encoding/csv"
	"flag"
	"io"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"time"

	"github.com/tealeg/xlsx"
	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/transform"
)

type fPos struct {
	hokenjya int
	kigo     int
	bango    int
	seinen   int
	jday     int
	cose     int
}

func failOnError(err error) {
	if err != nil {
		log.Fatal("Error:", err)
	}
}

func main() {
	flag.Parse()

	// ログファイル準備
	logfile, err := os.OpenFile("./log.txt", os.O_APPEND|os.O_CREATE|os.O_WRONLY, os.ModePerm)
	failOnError(err)
	defer logfile.Close()

	log.SetOutput(logfile)
	log.Print("Start\r\n")

	//ドロップされたファイルの数で処理を分ける
	filesu := flag.NArg()
	if (filesu == 0) || (filesu > 2) {
		log.Print("ドロップファイルエラー。処理を終了します。")
		os.Exit(1)
	}

	if filesu == 1 {
		// ファイルを読み込んで二次元配列に入れる
		records := readfile(flag.Arg(0))

		// 協会けんぽ資格用のファイル処理
		precs := processRecord(records)

		// ファイルを協会けんぽ資格確認用のCSVに出力
		saveCsv(precs)

	} else {
		// ファイルの読み込み
		recordsK, recordsN := readfile2(flag.Arg(0), flag.Arg(1))

		// ファイルの出力
		outDir, _ := filepath.Split(flag.Arg(0))
		saveExcel(outDir, recordsK, recordsN)
	}

	log.Print("Finesh !\r\n")

}

func readfile(filename string) [][]string {
	// 入力ファイル準備
	infile, err := os.Open(filename)
	failOnError(err)
	defer infile.Close()

	reader := csv.NewReader(transform.NewReader(infile, japanese.ShiftJIS.NewDecoder()))
	reader.Comma = '\t'

	//CSVファイルを２次元配列に展開
	readrecords := make([][]string, 0)
	record, err := reader.Read() // 1行読み出す
	if err == io.EOF {
		return readrecords
	} else {
		failOnError(err)
	}

	colMax := len(record) - 1
	//readrecords = append(readrecords, record[:colMax])
	readrecords = append(readrecords, record[:colMax])

	for {
		record, err := reader.Read() // 1行読み出す
		if err == io.EOF {
			break
		} else {
			// log.Print(record)
			// log.Print(len(record))
			failOnError(err)
		}

		readrecords = append(readrecords, record[:colMax])

	}

	return readrecords
}

func processRecord(precs [][]string) [][]string {

	// 協会けんぽ資格確認用の項目を抽出する
	filePos := fPos{hokenjya: -1, kigo: -1, bango: -1, seinen: -1, jday: -1, cose: -1}

	hedRow := precs[0]
	for pos, colName := range hedRow {
		switch colName {
		case "保険者番号":
			filePos.hokenjya = pos
		case "健康保険記号":
			filePos.kigo = pos
		case "健康保険番号":
			filePos.bango = pos
		case "生年月日":
			filePos.seinen = pos
		case "受診日":
			filePos.jday = pos
		case "ｺｰｽ区分ｺｰﾄﾞ":
			filePos.cose = pos
		}
	}

	// 協会けんぽ資格確認用の項目があったかチェック
	kcheck := true

	if filePos.hokenjya == -1 {
		kcheck = false
		log.Print("保険者番号の項目がありませんでした")
	} else if filePos.kigo == -1 {
		kcheck = false
		log.Print("健康保険記号の項目がありませんでした")
	} else if filePos.bango == -1 {
		kcheck = false
		log.Print("健康保険番号の項目がありませんでした")
	} else if filePos.seinen == -1 {
		kcheck = false
		log.Print("生年月日の項目がありませんでした")
	} else if filePos.jday == -1 {
		kcheck = false
		log.Print("受診日の項目がありませんでした")
	} else if filePos.cose == -1 {
		kcheck = false
		log.Print("ｺｰｽ区分ｺｰﾄﾞの項目がありませんでした")
	}

	if !kcheck {
		//failOnError("必要な項目が足りない為、処理を中止します。")
		log.Print("必要な項目が足りない為、処理を中止します。")
		os.Exit(1)
	}

	// 協会けんぽ資格確認用の対象者および項目のレコードを作成する
	krecs := make([][]string, 0)
	for i, v := range precs {
		if (i != 0) && coseCheck(v[filePos.cose]) {
			k1 := right("00000000"+v[filePos.hokenjya], 8)
			k2 := right("00000000"+v[filePos.kigo], 8)
			k3 := right("0000000"+v[filePos.bango], 7)
			k4 := "00"
			k5 := setSeireki(v[filePos.seinen])
			k6 := yymmdd(v[filePos.jday])
			k7 := setCose(v[filePos.cose])
			//log.Print(krec)
			krecs = append(krecs, []string{k1, k2, k3, k4, k5, k6, k7})
		}
	}

	return krecs

}

func saveCsv(recs [][]string) {

	// 出力ファイル準備
	outfile, err := os.Create("./SIKAKU_1311131242_" + time.Now().Format("0102") + ".csv")
	failOnError(err)
	defer outfile.Close()

	writer := csv.NewWriter(transform.NewWriter(outfile, japanese.ShiftJIS.NewEncoder()))
	writer.Comma = ','
	writer.UseCRLF = true

	for _, recRow := range recs {
		writer.Write(recRow)
	}

	writer.Flush()

}

func readfile2(file1 string, file2 string) ([][]string, [][]string) {

	// 協会けんぽ結果ファイルか、予約台帳のファイルか判別する
	var fileK string // 協会けんぽのダウンロードファイル名を入れる
	var fileN string // 健診システムの予約台帳のファイル名を入れる

	infileChk, err := os.Open(file1)
	failOnError(err)
	defer infileChk.Close()

	readerChk := csv.NewReader(transform.NewReader(infileChk, japanese.ShiftJIS.NewDecoder()))
	readerChk.Comma = ','

	readChk, err := readerChk.Read()
	failOnError(err)

	if readChk[0] == "保険者番号（支部コード）" {
		fileK = file1
		fileN = file2
	} else {
		fileK = file2
		fileN = file1
	}

	// 協会けんぽ結果ファイルを読み込む
	infileK, err := os.Open(fileK)
	failOnError(err)
	defer infileK.Close()

	readerK := csv.NewReader(transform.NewReader(infileK, japanese.ShiftJIS.NewDecoder()))
	readerK.Comma = ','

	recordsK := make([][]string, 0)
	for {
		recordK, err := readerK.Read()
		if err == io.EOF {
			break
		} else {
			failOnError(err)
		}

		recordsK = append(recordsK, recordK)

	}

	// 予約台帳のファイルを読み込む
	infileN, err := os.Open(fileN)
	failOnError(err)
	defer infileN.Close()

	readerN := csv.NewReader(transform.NewReader(infileN, japanese.ShiftJIS.NewDecoder()))
	readerN.Comma = '\t'

	recordsN := make([][]string, 0)
	for {
		recordN, err := readerN.Read()
		if err == io.EOF {
			break
		} else {
			failOnError(err)
		}

		recordsN = append(recordsN, recordN)

	}

	return recordsK, recordsN

}

func saveExcel(dir string, recsK [][]string, recsN [][]string) {
	//エクセルファイル処理
	var cellK string
	var cell *xlsx.Cell

	day := time.Now()
	excelName := dir + "協会けんぽ受診資格結果" + day.Format("20060102") + ".xlsx"
	excelFile := xlsx.NewFile()
	xlsx.SetDefaultFont(11, "游ゴシック")
	sheet, err := excelFile.AddSheet("データ")
	failOnError(err)

	//予約台帳の項目確認
	var syo1pos int
	var namepos int
	var sexpos int
	var kigopos int
	var bangopos int

	for pos, recHead := range recsN[0] {
		switch recHead {
		case "所属名１":
			syo1pos = pos
		case "受診者名":
			namepos = pos
		case "性別":
			sexpos = pos
		case "健康保険記号":
			kigopos = pos
		case "健康保険番号":
			bangopos = pos
		}
	}

	//協会けんぽ結果と予約台帳のマッチング処理
	for i, rowK := range recsK {
		row := sheet.AddRow()
		if i == 0 { // タイトル行の処理
			cell = row.AddCell()
			cell.Value = "所属名１"
			cell = row.AddCell()
			cell.Value = "受診者名"
			cell = row.AddCell()
			cell.Value = "性別"

			for _, cellK = range rowK {
				cell = row.AddCell()
				cell.Value = cellK
			}

		} else { // タイトル行以外の処理
			var kigo string
			var bango string
			var syo1str string
			var namestr string
			var sexstr string

			for _, rowN := range recsN {
				kigo = right("00000000"+rowN[kigopos], 8)
				bango = right("0000000"+rowN[bangopos], 7)
				syo1str = ""
				namestr = ""
				sexstr = ""
				if (rowK[1] == kigo) && (rowK[2] == bango) { // 保険証記号と番号が一致
					syo1str = rowN[syo1pos]
					namestr = rowN[namepos]
					sexstr = rowN[sexpos]
					break
				}
			}

			cell = row.AddCell()
			cell.Value = syo1str
			cell = row.AddCell()
			cell.Value = namestr
			cell = row.AddCell()
			cell.Value = sexstr
			for j, cellK := range rowK {
				cell = row.AddCell()
				switch j {
				case 4, 5:
					cell.Value = cellK[0:4] + "/" + cellK[4:6] + "/" + cellK[6:]
				case 6:
					cell.Value = setCoseName(cellK)
				default:
					cell.Value = cellK
				}
			}
		}
	}

	err = excelFile.Save(excelName)
	failOnError(err)

}

func coseCheck(cose string) bool {
	// 協会けんぽのコースチェック
	var check bool

	if (cose == "19") || (cose == "20") || (cose == "21") {
		check = true
	} else {
		check = false
	}

	return check

}

func right(v string, i int) string {
	l := len(v)
	s := v[l-i:]
	return s
}

func setSeireki(v string) string {
	Wa := v[0:1]
	Y := v[1:3]
	iY, _ := strconv.Atoi(Y)
	M := v[4:6]
	D := v[7:9]

	switch Wa {
	case "M":
		iY = 1900 + iY - 33
	case "T":
		iY = 1900 + iY + 11
	case "S":
		iY = 1900 + iY + 25
	case "H":
		iY = 1900 + iY + 88
	}

	return strconv.Itoa(iY) + M + D
}

func yymmdd(v string) string {
	return v[0:4] + v[5:7] + v[8:10]
}

func setCose(v string) string {
	var s string

	switch v {
	case "19":
		s = "1"
	case "20":
		s = "2"
	case "21":
		s = "3"
	}

	return s
}

func setCoseName(v string) string {
	var s string

	switch v {
	case "1":
		s = "一般健診"
	case "2":
		s = "一般健診＋付加"
	case "3":
		s = "子宮がん単独"
	}

	return s
}
