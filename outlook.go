package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"log"
	"os"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func setupLogging() {
	//set logging
	file, err := os.OpenFile("_logs.txt", os.O_APPEND|os.O_CREATE|os.O_RDONLY, 0666)
	if err != nil {
		log.Fatal(err)
	}
	//if i do this only then the output will be in the txt file and not on the console
	//must use multiwriter
	// log.SetOutput(file)
	multiwriter := io.MultiWriter(file, os.Stdout)
	log.SetOutput(multiwriter)
}

func outlookFind() {
	setupLogging()
	fmt.Println("Logging initialized")

	// Initialize COM
	err := ole.CoInitialize(0)
	if err != nil {
		log.Printf("Failed to initialize COM library: %v", err)
		stopPrompt()
		return
	}
	defer ole.CoUninitialize()

	// Connect to Outlook
	outlook, err := oleutil.CreateObject("Outlook.Application")
	if err != nil {
		log.Printf("Failed to connect to Outlook: %v", err)
		stopPrompt()
		return
	}
	defer outlook.Release()

	// Convert *ole.IUnknown to *ole.IDispatch
	outlookDispatch, err := outlook.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Printf("Failed to get IDispatch: %v", err)
		stopPrompt()
		return
	}
	defer outlookDispatch.Release()

	// Get Namespace MAPI
	ns, err := oleutil.CallMethod(outlookDispatch, "GetNamespace", "MAPI")
	if err != nil {
		log.Printf("Failed to get MAPI namespace: %v", err)
		stopPrompt()
		return
	}
	mapi := ns.ToIDispatch()
	defer mapi.Release()

	// Get the Global Address List
	addressLists := oleutil.MustGetProperty(mapi, "AddressLists").ToIDispatch()
	defer addressLists.Release()

	gal := oleutil.MustGetProperty(addressLists, "Item", "Global Address List").ToIDispatch()
	defer gal.Release()

	addressEntries := oleutil.MustGetProperty(gal, "AddressEntries").ToIDispatch()
	defer addressEntries.Release()

	// Load CSV file with names
	file, err := os.Open(csvFile)
	if err != nil {
		fmt.Printf("Failed to open %s", csvFile)
		stopPrompt()
		return
	}
	defer file.Close()

	reader := csv.NewReader(file)
	records, err := reader.ReadAll()
	if err != nil {
		// i := 0
		fmt.Printf("Failed to read %s", csvFile)
		stopPrompt()
		return
	}

	// Create/Truncate the output file
	outputFile, err := os.Create(outputTxt)
	if err != nil {
		fmt.Printf("Failed to create %s", outputTxt)
		stopPrompt()
		return
	}
	defer outputFile.Close()

	fullSmtpList := map[string]bool{}

	// Iterate over CSV file
	for _, record := range records {
		nameX := strings.TrimSpace(record[0])
		log.Printf("Searching for %s...\n", nameX)

		// Search GAL
		entryResult, err := oleutil.CallMethod(addressEntries, "Item", nameX)
		if err != nil {
			log.Printf("Not in GAL, skipping %s...\n", nameX)
			continue
		}

		entry := entryResult.ToIDispatch()
		if entry == nil {
			log.Printf("Entry found for %s is nil, skipping...\n", nameX)
			continue
		}
		//possible to be more precise in the output found by outlook engine
		// Get the Display Name of the entry found
		// displayName := oleutil.MustGetProperty(entry, "Name").ToString()
		// log.Printf("Found->%s", displayName)

		// Get the Exchange User
		userEmailResult, err := oleutil.GetProperty(entry, "GetExchangeUser")
		if err != nil {
			log.Printf("Could not get Exchange User for %s, skipping...\n", nameX)
			entry.Release()
			continue
		}

		userEmail := userEmailResult.ToIDispatch()
		if userEmail == nil {
			log.Printf("Exchange User for %s is nil, skipping...\n", nameX)
			entry.Release()
			continue
		}

		// Get the Primary SMTP Address
		smtpAddressResult, err := oleutil.GetProperty(userEmail, "PrimarySmtpAddress")
		if err != nil {
			log.Printf("Could not get SMTP address for %s, skipping...\n", nameX)
			userEmail.Release()
			entry.Release()
			continue
		}

		smtpAddress := smtpAddressResult.ToString()
		if smtpAddress == "" {
			log.Printf("SMTP address for %s is empty, skipping...\n", nameX)
			smtpAddressResult.Clear()
			userEmail.Release()
			entry.Release()
			continue
		}

		log.Printf("Found email: %s\n", smtpAddress)

		// Add to Hash map. unique entries
		_, ok := fullSmtpList[smtpAddress]
		if !ok {
			fullSmtpList[smtpAddress] = true
		}

		// Release COM objects
		smtpAddressResult.Clear()
		userEmail.Release()
		entry.Release()
	}

	// Write Struct collection to txt file
	for key := range fullSmtpList {
		_, err = outputFile.WriteString(key + "\n")
		if err != nil {
			log.Fatalf("Failed to write to %s file: %v", outputTxt, err)
		}
	}

	fmt.Print("Successfully processed all records")
	stopPrompt()
}
