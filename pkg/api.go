package pkg

import (
	"bytes"
	"fmt"
	"io"
	"mime/multipart"
	"net/http"
	"os"
	"path/filepath"
)

func Api(outputName string) (filePath string) {
	str1 := outputName + ".xlsx"
	url := "https://api.pspdfkit.com/build"
	method := "POST"

	payload := &bytes.Buffer{}
	writer := multipart.NewWriter(payload)
	//_ = writer.WriteField("instructions", "{\n  \"parts\": [\n    {\n      \"file\":"+str1+"\n    }\n  ]\n}\n")
	_ = writer.WriteField("instructions", "{\n  \"parts\": [\n    {\n      \"file\": \"piModle\"\n    }\n  ]\n}\n")
	file, errFile2 := os.Open(str1)
	defer file.Close()
	part2,
		errFile2 := writer.CreateFormFile("piModle", filepath.Base(str1))
	_, errFile2 = io.Copy(part2, file)
	if errFile2 != nil {
		fmt.Println(errFile2)
		return
	}
	err := writer.Close()
	if err != nil {
		fmt.Println(err)
		return
	}

	client := &http.Client{}
	req, err := http.NewRequest(method, url, payload)

	if err != nil {
		fmt.Println(err)
		return
	}
	req.Header.Add("Authorization", "Bearer pdf_live_bxnx3bbyLOOxkLW59E8yspzKcX1sa0go7J2tWGD5Dgh")
	req.Header.Add("Cookie", "AWSALB=zXv9IwcA9RTznlC5nNZ5fFM9KA3rmI7jTa7+WN4uPFfTGG/ivZJo50KTSFkKgUjoJdiCJtnydxPCyRqdJz+IaaM2LrMoRF0TqrsnlCz9jrw40XfSalNsqCEgRatn; AWSALBCORS=zXv9IwcA9RTznlC5nNZ5fFM9KA3rmI7jTa7+WN4uPFfTGG/ivZJo50KTSFkKgUjoJdiCJtnydxPCyRqdJz+IaaM2LrMoRF0TqrsnlCz9jrw40XfSalNsqCEgRatn")

	req.Header.Set("Content-Type", writer.FormDataContentType())
	res, err := client.Do(req)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer res.Body.Close()
	/*
		body, err := ioutil.ReadAll(res.Body)
		if err != nil {
			fmt.Println(err)
			return
		}*/
	output, err := os.Create("result.pdf")
	if err != nil {
		// handle error
	}
	defer output.Close()

	_, err = io.Copy(output, res.Body)
	if err != nil {
		// handle error
	}
	return outputName + ".xlsx"
}
