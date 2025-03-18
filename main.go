package main

import (
	"bytes"
	"encoding/csv"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"path/filepath"

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

// ElevenLabsRequest represents the request structure for ElevenLabs TTS API
type ElevenLabsRequest struct {
	Text          string        `json:"text"`
	ModelID       string        `json:"model_id"`
	VoiceID       string        `json:"voice_id"`
	VoiceSettings VoiceSettings `json:"voice_settings"`
}

type VoiceSettings struct {
	Stability       float64 `json:"stability"`
	SimilarityBoost float64 `json:"similarity_boost"`
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
	if err := csvWriter.Write([]string{"English Word", "Example sentence", "Sound", "Russian Translation"}); err != nil {
		log.Fatalf("Error writing header to CSV: %v", err)
	}

	// Load environment variables
	if err := godotenv.Load(); err != nil {
		log.Printf("Warning: .env file not found")
	}

	// Yandex Dictionary API
	yandexAPIKey := os.Getenv("YANDEX_API_KEY")
	if yandexAPIKey == "" {
		log.Fatal("YANDEX_API_KEY environment variable is required")
	}

	// ElevenLabs API
	elevenLabsAPIKey := os.Getenv("ELEVENLABS_API_KEY")
	if elevenLabsAPIKey == "" {
		log.Fatal("ELEVENLABS_API_KEY environment variable is required")
	}

	// Create audio directory if it doesn't exist
	audioDir := "audio"
	if err := os.MkdirAll(audioDir, 0755); err != nil {
		log.Fatalf("Failed to create audio directory: %v", err)
	}

	lang := "en-ru"
	yandexBaseURL := "https://dictionary.yandex.net/api/v1/dicservice.json/lookup"
	elevenLabsBaseURL := "https://api.elevenlabs.io/v1/text-to-speech"

	// Default voice ID - you can change this to any voice from ElevenLabs
	voiceID := "21m00Tcm4TlvDq8ikWAM" // Default voice - Rachel

	// Process each row in the Excel sheet.
	for _, row := range sheet.Rows {
		// Skip rows that do not have at least two cells.
		if len(row.Cells) < 2 {
			continue
		}

		// Read the English word and definition.
		word := row.Cells[0].String()
		definition := row.Cells[1].String()

		// Get an example sentence (using the definition from Excel)
		exampleSentence := definition

		// Build the Yandex API request URL.
		url := fmt.Sprintf("%s?key=%s&lang=%s&text=%s", yandexBaseURL, yandexAPIKey, lang, word)
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

		// Generate audio with ElevenLabs API
		audioFilename := fmt.Sprintf("%s.mp3", word)
		audioPath := filepath.Join(audioDir, audioFilename)

		// Check if audio file already exists, generate only if needed
		if _, err := os.Stat(audioPath); os.IsNotExist(err) {
			// Prepare request for ElevenLabs
			elevenLabsReq := ElevenLabsRequest{
				Text:    word,
				ModelID: "eleven_multilingual_v2",
				VoiceID: voiceID,
				VoiceSettings: VoiceSettings{
					Stability:       0.5,
					SimilarityBoost: 0.5,
				},
			}

			reqBody, err := json.Marshal(elevenLabsReq)
			if err != nil {
				log.Printf("Error creating request for ElevenLabs for %s: %v", word, err)
				continue
			}

			// Create the HTTP request
			req, err := http.NewRequest("POST", fmt.Sprintf("%s/%s", elevenLabsBaseURL, voiceID), bytes.NewBuffer(reqBody))
			if err != nil {
				log.Printf("Error creating HTTP request for %s: %v", word, err)
				continue
			}

			req.Header.Set("Content-Type", "application/json")
			req.Header.Set("xi-api-key", elevenLabsAPIKey)

			// Execute the request
			client := &http.Client{}
			resp, err := client.Do(req)
			if err != nil {
				log.Printf("Error generating audio for %s: %v", word, err)
				continue
			}
			defer resp.Body.Close()

			if resp.StatusCode != http.StatusOK {
				responseBody, _ := io.ReadAll(resp.Body)
				log.Printf("ElevenLabs API error for %s: %d - %s", word, resp.StatusCode, string(responseBody))
				continue
			}

			// Save the audio file
			audioFile, err := os.Create(audioPath)
			if err != nil {
				log.Printf("Error creating audio file for %s: %v", word, err)
				continue
			}

			_, err = io.Copy(audioFile, resp.Body)
			audioFile.Close()
			if err != nil {
				log.Printf("Error saving audio file for %s: %v", word, err)
				continue
			}

			log.Printf("Created audio file for: %s", word)
		} else {
			log.Printf("Audio file for %s already exists, skipping generation", word)
		}

		// Format for Anki: [sound:filename.mp3]
		soundField := fmt.Sprintf("[sound:%s]", audioFilename)

		// Write the output row to the CSV.
		err = csvWriter.Write([]string{word, exampleSentence, soundField, russian})
		if err != nil {
			log.Printf("Error writing CSV row for %s: %v", word, err)
		}
	}

	fmt.Println("Processing complete. Output written to output.csv")
	fmt.Printf("Audio files saved to the '%s' directory\n", audioDir)
}
