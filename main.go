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

	xlFile, err := xlsx.OpenFile(excelFile)
	if err != nil {
		log.Fatalf("Failed to open Excel file: %v", err)
		return
	}

	if len(xlFile.Sheets) == 0 {
		log.Fatalf("No sheets found in the Excel file.")
		return
	}
	sheet := xlFile.Sheets[0]

	totalWords := 0
	for _, row := range sheet.Rows {
		if len(row.Cells) >= 2 {
			totalWords++
		}
	}

	outputFile, err := os.Create("output.csv")
	if err != nil {
		log.Fatalf("Failed to create output.csv: %v", err)
		return
	}
	defer outputFile.Close()

	csvWriter := csv.NewWriter(outputFile)
	csvWriter.Comma = ';'
	defer csvWriter.Flush()

	if err := godotenv.Load(); err != nil {
		log.Printf("Warning: .env file not found")
	}

	yandexAPIKey := os.Getenv("YANDEX_API_KEY")
	if yandexAPIKey == "" {
		log.Fatal("YANDEX_API_KEY environment variable is required")
		return
	}

	elevenLabsAPIKey := os.Getenv("ELEVENLABS_API_KEY")
	if elevenLabsAPIKey == "" {
		log.Fatal("ELEVENLABS_API_KEY environment variable is required")
		return
	}

	audioDir := "audio"
	if err := os.MkdirAll(audioDir, 0755); err != nil {
		log.Fatalf("Failed to create audio directory: %v", err)
		return
	}

	lang := "en-ru"
	yandexBaseURL := "https://dictionary.yandex.net/api/v1/dicservice.json/lookup"
	elevenLabsBaseURL := "https://api.elevenlabs.io/v1/text-to-speech"

	voiceID := "21m00Tcm4TlvDq8ikWAM"

	processedWords := 0

	for _, row := range sheet.Rows {
		// Skip rows that do not have at least two cells.
		if len(row.Cells) < 2 {
			continue
		}

		// Read the English word and definition.
		word := row.Cells[0].String()
		definition := row.Cells[1].String()

		// Print progress information
		fmt.Printf("\r\033[2KProcessing word: %s\n", word)
		fmt.Printf("Current progress: %d/%d", processedWords, totalWords)

		// Get an example sentence (using the definition from Excel)
		exampleSentence := definition

		// Build the Yandex API request URL.
		url := fmt.Sprintf("%s?key=%s&lang=%s&text=%s", yandexBaseURL, yandexAPIKey, lang, word)
		resp, err := http.Get(url)
		if err != nil {
			log.Printf("\r\033[2KError fetching translation for %s: %v", word, err)
			fmt.Printf("Current progress: %d/%d", processedWords, totalWords)
			continue
		}
		body, err := io.ReadAll(resp.Body)
		resp.Body.Close()
		if err != nil {
			log.Printf("\r\033[2KError reading response for %s: %v", word, err)
			fmt.Printf("Current progress: %d/%d", processedWords, totalWords)
			continue
		}

		var result DicResult
		if err := json.Unmarshal(body, &result); err != nil {
			log.Printf("\r\033[2KError parsing JSON for %s: %v", word, err)
			fmt.Printf("Current progress: %d/%d", processedWords, totalWords)
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
				log.Printf("\r\033[2KError creating request for ElevenLabs for %s: %v", word, err)
				fmt.Printf("Current progress: %d/%d", processedWords, totalWords)
				continue
			}

			// Create the HTTP request
			req, err := http.NewRequest("POST", fmt.Sprintf("%s/%s", elevenLabsBaseURL, voiceID), bytes.NewBuffer(reqBody))
			if err != nil {
				log.Printf("\r\033[2KError creating HTTP request for %s: %v", word, err)
				fmt.Printf("Current progress: %d/%d", processedWords, totalWords)
				continue
			}

			req.Header.Set("Content-Type", "application/json")
			req.Header.Set("xi-api-key", elevenLabsAPIKey)

			// Execute the request
			client := &http.Client{}
			resp, err := client.Do(req)
			if err != nil {
				log.Printf("\r\033[2KError generating audio for %s: %v", word, err)
				fmt.Printf("Current progress: %d/%d", processedWords, totalWords)
				continue
			}
			defer resp.Body.Close()

			if resp.StatusCode != http.StatusOK {
				responseBody, _ := io.ReadAll(resp.Body)
				log.Printf("\r\033[2KElevenLabs API error for %s: %d - %s", word, resp.StatusCode, string(responseBody))
				fmt.Printf("Current progress: %d/%d", processedWords, totalWords)
				continue
			}

			// Save the audio file
			audioFile, err := os.Create(audioPath)
			if err != nil {
				log.Printf("\r\033[2KError creating audio file for %s: %v", word, err)
				fmt.Printf("Current progress: %d/%d", processedWords, totalWords)
				continue
			}

			_, err = io.Copy(audioFile, resp.Body)
			audioFile.Close()
			if err != nil {
				log.Printf("\r\033[2KError saving audio file for %s: %v", word, err)
				fmt.Printf("Current progress: %d/%d", processedWords, totalWords)
				continue
			}

			log.Printf("\r\033[2KCreated audio file for: %s", word)
			fmt.Printf("Current progress: %d/%d", processedWords, totalWords)
		} else {
			log.Printf("\r\033[2KAudio file for %s already exists, skipping generation", word)
			fmt.Printf("Current progress: %d/%d", processedWords, totalWords)
		}

		// Format for Anki: [sound:filename.mp3]
		soundField := fmt.Sprintf("[sound:%s]", audioFilename)

		// Write the output row to the CSV, ensuring proper handling of fields with semicolons
		// The csv.Writer will automatically handle quoting and escaping when needed
		err = csvWriter.Write([]string{word, exampleSentence, soundField, russian})
		if err != nil {
			log.Printf("\r\033[2KError writing CSV row for %s: %v", word, err)
			fmt.Printf("Current progress: %d/%d", processedWords, totalWords)
		}

		// Update progress counter and display
		processedWords++
	}

	fmt.Printf("\r\033[2KProcessing %d words complete. Output written to output.csv\n", totalWords)
	fmt.Printf("Audio files saved to the '%s' directory\n", audioDir)
}
