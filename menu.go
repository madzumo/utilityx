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
utility |/             |/  
   X    *--------------*   
`

func printMenu() int {
	cTitle := color.New(color.BgMagenta)
	cMenu := color.New(color.FgBlue).Add(color.Bold)
	cQuit := color.New(color.FgRed).Add(color.Bold)
	cTitle.Println(menuText)
	fmt.Println("==========================================")
	cMenu.Println("1. Web Scrapper")
	cMenu.Println("2. GAL Lookup -> EMail export")
	// cMenu.Println("3. output txt (default: emails.txt)")
	// c2.Println("4. RUN Utility")
	cQuit.Println("3. QUIT")
	fmt.Println("==========================================")

	menuOption := 0
	fmt.Print("Enter Menu Option: ")
	fmt.Scan(&menuOption)
	return menuOption
}
