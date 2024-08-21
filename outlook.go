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
	ole.CoInitialize(0)
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
		// log.Fatalf("Failed to open CSV file: %v", err)
		fmt.Print("Failed to open CSV file:...Press Enter to continue.")
		reader := bufio.NewReader(os.Stdin)
		_, _ = reader.ReadString('\n')
	}
	defer file.Close()

	reader := csv.NewReader(file)
	records, err := reader.ReadAll()
	if err != nil {
		// log.Fatalf("Failed to read CSV file: %v", err)
		fmt.Print("Failed to read CSV file:...Press Enter to continue.")
		reader := bufio.NewReader(os.Stdin)
		_, _ = reader.ReadString('\n')
	}

	// Create/Truncate the output file
	outputFile, err := os.Create(outputTxtFile)
	if err != nil {
		// log.Fatalf("Failed to create output file: %v", err)
		fmt.Print("Failed to create output file:...Press Enter to continue.")
		reader := bufio.NewReader(os.Stdin)
		_, _ = reader.ReadString('\n')
	}
	defer outputFile.Close()

	// Iterate over CSV records
	for _, record := range records {
		name := strings.TrimSpace(record[0])
		fmt.Printf("Searching for %s...\n", name)

		// Find the user in the Global Address List
		entry, err := oleutil.CallMethod(addressEntries, "Item", name)
		if err != nil {
			fmt.Printf("Could not find %s in Global Address List, skipping...\n", name)
			continue
		}
		defer entry.ToIDispatch().Release()

		// Get the user's SMTP address
		userEmail, err := oleutil.GetProperty(entry.ToIDispatch(), "GetExchangeUser")
		if err != nil {
			fmt.Printf("Could not get Exchange User for %s, skipping...\n", name)
			continue
		}
		defer userEmail.ToIDispatch().Release()

		smtpAddress, err := oleutil.GetProperty(userEmail.ToIDispatch(), "PrimarySmtpAddress")
		if err != nil {
			fmt.Printf("Could not get SMTP address for %s, skipping...\n", name)
			continue
		}

		email := smtpAddress.ToString()
		fmt.Printf("Found email: %s\n", email)

		// Write the found email address to the output file
		_, err = outputFile.WriteString(email + ";\n")
		if err != nil {
			log.Fatalf("Failed to write to output file: %v", err)
		}
		fmt.Print("Successfully processed all records...Press Enter to continue.")
		r := bufio.NewReader(os.Stdin)
		_, _ = r.ReadString('\n')
	}
}
