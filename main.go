package main

import (
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	http.HandleFunc("/xlsx", handler)
	log.Fatal(http.ListenAndServe(":51080", nil))
}

func handler(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Access-Control-Allow-Origin", "*")
	w.Header().Set("Access-Control-Expose-Headers", "*")
	if r.Method == http.MethodOptions {
		w.Header().Set("Access-Control-Allow-Methods", "POST")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type")
		w.WriteHeader(http.StatusNoContent)
		return
	}
	if r.Method != http.MethodPost {
		w.WriteHeader(http.StatusMethodNotAllowed)
		return
	}
	tab, err := parse(r.Body)
	if err != nil {
		log.Println(err)
		w.WriteHeader(http.StatusUnprocessableEntity)
		return
	}
	f, err := xlsx(tab)
	if err != nil {
		log.Println(err)
		w.WriteHeader(http.StatusInternalServerError)
		return
	}
	w.Header().Set("X-Filename", f.Path)
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	if err := f.Write(w); err != nil {
		log.Println(err)
		w.WriteHeader(http.StatusInternalServerError)
	}
}

func parse(body io.ReadCloser) (map[string][]interface{}, error) {
	var tab map[string][]interface{}
	dec := json.NewDecoder(body)
	err := dec.Decode(&tab)
	if err != nil {
		return nil, err
	}
	return tab, nil
}

func xlsx(tab map[string][]interface{}) (*excelize.File, error) {
	f := excelize.NewFile()
	if err := f.SetColWidth("Sheet1", "A", "N", 16); err != nil {
		return nil, err
	}
	for k, v := range tab {
		if err := f.SetSheetRow("Sheet1", k, &v); err != nil {
			return nil, err
		}
	}
	f.Path = fmt.Sprintf("%s.xlsx", time.Now().Format("20060102150405"))
	return f, nil
}
