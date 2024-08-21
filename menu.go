package main

import (
	"fmt"

	"github.com/fatih/color"
)

var menuText = `
           +--------------+
          /|             /|
         / |            / |
        *--+-----------*  |
        |  |           |  |
        |  |           |  |
        |  |           |  |
        |  +-----------+--+
        | /            | / 
        |/             |/  
J.M.    *--------------*   
`

func printMenu() int {
	c1 := color.New(color.BgMagenta)
	c2 := color.New(color.FgGreen).Add(color.Bold)
	c3 := color.New(color.FgHiYellow).Add(color.Bold)
	c5 := color.New(color.FgRed).Add(color.Bold)
	c1.Println(menuText)
	fmt.Println("==========================================")
	c3.Println("1. csv file (default: names.csv)")
	c3.Println("2. output txt (default: emails.txt)")
	c2.Println("3. RUN Utility")
	c5.Println("4. QUIT")
	fmt.Println("==========================================")

	menuOption := 0
	fmt.Print("Enter Menu Option: ")
	fmt.Scan(&menuOption)
	return menuOption
}

func clearScreen() {
	fmt.Print("\033[H\033[2J")
}
