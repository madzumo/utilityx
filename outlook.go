package main

import (
	"bufio"
	"encoding/csv"
	"fmt"
	"log"
	"os"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func outlookFind(csvFile string, outputTxtFile string) {
	// Initialize COM
	err := ole.CoInitialize(0)
	if err != nil {
		log.Fatalf("Failed to initialize COM library: %v", err)
	}
	defer ole.CoUninitialize()

	// Connect to Outlook
	outlook, err := oleutil.CreateObject("Outlook.Application")
	if err != nil {
		log.Fatalf("Failed to connect to Outlook: %v", err)
	}
	defer outlook.Release()

	// Convert *ole.IUnknown to *ole.IDispatch
	outlookDispatch, err := outlook.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Fatalf("Failed to get IDispatch: %v", err)
	}
	defer outlookDispatch.Release()

	// Get Namespace MAPI
	ns, err := oleutil.CallMethod(outlookDispatch, "GetNamespace", "MAPI")
	if err != nil {
		log.Fatalf("Failed to get MAPI namespace: %v", err)
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
		fmt.Print("Failed to open CSV file:...Press Enter to continue.")
		reader := bufio.NewReader(os.Stdin)
		_, _ = reader.ReadString('\n')
		return // Exit function if file cannot be opened
	}
	defer file.Close()

	reader := csv.NewReader(file)
	records, err := reader.ReadAll()
	if err != nil {
		fmt.Print("Failed to read CSV file:...Press Enter to continue.")
		reader := bufio.NewReader(os.Stdin)
		_, _ = reader.ReadString('\n')
		return // Exit function if file cannot be read
	}

	// Create/Truncate the output file
	outputFile, err := os.Create(outputTxtFile)
	if err != nil {
		fmt.Print("Failed to create output file:...Press Enter to continue.")
		reader := bufio.NewReader(os.Stdin)
		_, _ = reader.ReadString('\n')
		return // Exit function if output file cannot be created
	}
	defer outputFile.Close()

	fullSmtpList := map[string]bool{}

	// Iterate over CSV file
	for _, record := range records {
		nameX := strings.TrimSpace(record[0])
		fmt.Printf("Searching for %s...\n", nameX)

		// Search GAL
		entryResult, err := oleutil.CallMethod(addressEntries, "Item", nameX)
		if err != nil {
			fmt.Printf("Not in GAL, skipping %s...\n", nameX)
			continue
		}

		entry := entryResult.ToIDispatch()
		if entry == nil {
			fmt.Printf("Entry found for %s is nil, skipping...\n", nameX)
			continue
		}

		// Get the Exchange User
		userEmailResult, err := oleutil.GetProperty(entry, "GetExchangeUser")
		if err != nil {
			fmt.Printf("Could not get Exchange User for %s, skipping...\n", nameX)
			entry.Release()
			continue
		}

		userEmail := userEmailResult.ToIDispatch()
		if userEmail == nil {
			fmt.Printf("Exchange User for %s is nil, skipping...\n", nameX)
			entry.Release()
			continue
		}

		// Get the Primary SMTP Address
		smtpAddressResult, err := oleutil.GetProperty(userEmail, "PrimarySmtpAddress")
		if err != nil {
			fmt.Printf("Could not get SMTP address for %s, skipping...\n", nameX)
			userEmail.Release()
			entry.Release()
			continue
		}

		smtpAddress := smtpAddressResult.ToString()
		if smtpAddress == "" {
			fmt.Printf("SMTP address for %s is empty, skipping...\n", nameX)
			smtpAddressResult.Clear()
			userEmail.Release()
			entry.Release()
			continue
		}

		fmt.Printf("Found email: %s\n", smtpAddress)

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
			log.Fatalf("Failed to write to output file: %v", err)
		}
	}

	// os.Stdout.Sync()
	// clearInputBuffer()
	// rox := bufio.NewReader(os.Stdin)

	r := bufio.NewReader(os.Stdin)
	fmt.Print("Successfully processed all records...")
	_, _ = r.ReadString('\n')

	// _, err = rox.ReadString('\n')
	// if err != nil {
	// 	fmt.Println("Error reading input:", err)
	// }
}
