package main

import (
	"fmt"

	"github.com/fatih/color"
)

var csvFile string = "names.csv"
var outputTxt string = "emails.txt"
var color1 = color.New(color.BgBlue).Add(color.Bold)
var color2 = color.New(color.BgHiMagenta)

func main() {
	for {
		switch printMenu() {
		case 1:
			color2.Print("Enter CSV file name: ")
			fmt.Scan(&csvFile)
			clearScreen()
		case 2:
			color2.Print("Enter output TXT file: ")
			fmt.Scan(&outputTxt)
			clearScreen()
		case 3:
			clearScreen()
			fmt.Println("==========================================")
			color1.Printf("Csv file: %s\n", csvFile)
			color1.Printf("OutputTXt: %s\n", outputTxt)
			fmt.Println("==========================================")
			outlookFind(csvFile, outputTxt)
			clearScreen()
		case 4:
			return
		}
	}
}
