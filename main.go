package main

import (
	"encoding/csv"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"

	"github.com/joho/godotenv"
	"github.com/tealeg/xlsx"
)

// DicResult represents the structure of the Yandex.Dictionary API JSON response.
type DicResult struct {
	Head any          `json:"head"`
	Def  []Definition `json:"def"`
}

type Definition struct {
	Text string        `json:"text"`
	Pos  string        `json:"pos"`
	Tr   []Translation `json:"tr"`
}

type Translation struct {
	Text string    `json:"text"`
	Pos  string    `json:"pos"`
	Syn  []Synonym `json:"syn,omitempty"`
	Mean []Meaning `json:"mean,omitempty"`
	Ex   []Example `json:"ex,omitempty"`
}

type Synonym struct {
	Text string `json:"text"`
}

type Meaning struct {
	Text string `json:"text"`
}

type Example struct {
	Text string        `json:"text"`
	Tr   []Translation `json:"tr"`
}

func main() {
	if len(os.Args) < 2 {
		fmt.Println("Usage: go run main.go <excel_file>")
		return
	}
	excelFile := os.Args[1]

	// Open the Excel file.
	xlFile, err := xlsx.OpenFile(excelFile)
	if err != nil {
		log.Fatalf("Failed to open Excel file: %v", err)
	}

	// Use the first sheet.
	if len(xlFile.Sheets) == 0 {
		log.Fatalf("No sheets found in the Excel file.")
	}
	sheet := xlFile.Sheets[0]

	// Prepare the CSV output file.
	outputFile, err := os.Create("output.csv")
	if err != nil {
		log.Fatalf("Failed to create output.csv: %v", err)
	}
	defer outputFile.Close()
	csvWriter := csv.NewWriter(outputFile)
	defer csvWriter.Flush()

	// Write header row.
	if err := csvWriter.Write([]string{"English Word", "Russian Translation", "Definition"}); err != nil {
		log.Fatalf("Error writing header to CSV: %v", err)
	}

	// Load environment variables
	if err := godotenv.Load(); err != nil {
		log.Printf("Warning: .env file not found")
	}

	apiKey := os.Getenv("YANDEX_API_KEY")
	if apiKey == "" {
		log.Fatal("YANDEX_API_KEY environment variable is required")
	}

	lang := "en-ru"
	baseURL := "https://dictionary.yandex.net/api/v1/dicservice.json/lookup"

	// Process each row in the Excel sheet.
	for _, row := range sheet.Rows {
		// Skip rows that do not have at least two cells.
		if len(row.Cells) < 2 {
			continue
		}

		// Read the English word and definition.
		word := row.Cells[0].String()
		definition := row.Cells[1].String()

		// Build the API request URL.
		url := fmt.Sprintf("%s?key=%s&lang=%s&text=%s", baseURL, apiKey, lang, word)
		resp, err := http.Get(url)
		if err != nil {
			log.Printf("Error fetching translation for %s: %v", word, err)
			continue
		}
		body, err := io.ReadAll(resp.Body)
		resp.Body.Close()
		if err != nil {
			log.Printf("Error reading response for %s: %v", word, err)
			continue
		}

		var result DicResult
		if err := json.Unmarshal(body, &result); err != nil {
			log.Printf("Error parsing JSON for %s: %v", word, err)
			continue
		}

		// Retrieve the first translation from the result, if available.
		russian := ""
		if len(result.Def) > 0 && len(result.Def[0].Tr) > 0 {
			russian = result.Def[0].Tr[0].Text
		}

		// Write the output row to the CSV.
		err = csvWriter.Write([]string{word, russian, definition})
		if err != nil {
			log.Printf("Error writing CSV row for %s: %v", word, err)
		}
	}

	fmt.Println("Processing complete. Output written to output.csv")
}
