package main

import (
	"bufio"
	"fmt"
	"os"

	"github.com/fatih/color"
)

var csvFile string = "names.csv"
var outputTxt string = "emails_output.txt"
var colorPrompt = color.New(color.BgBlue).Add(color.Bold)
var colorExit = color.New(color.FgRed).Add(color.Bold)

func main() {
	for {
		switch printMenu() {
		case 1:
			scrapper()
			return
		case 2:
			clearScreen()
			fmt.Println("==========================================")
			colorPrompt.Printf("Csv file: %s\n", csvFile)
			colorPrompt.Printf("OutputTXt: %s\n", outputTxt)
			fmt.Println("==========================================")
			outlookFind()
			clearScreen()
		case 3:
			return
		}
	}
}

func stopPrompt() {
	colorExit.Print("\nType Q to continue\n")
	reader := bufio.NewReader(os.Stdin)
	_, _ = reader.ReadString('q')
}

func clearScreen() {
	fmt.Print("\033[H\033[2J")
}
