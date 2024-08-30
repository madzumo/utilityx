package main

import (
	"fmt"

	"github.com/fatih/color"
)

var menuOptions = map[int]string{
	1: "Web Scrapper",
	2: "GAL Lookup -> EMail export",
	3: "Quit",
}

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
	// cQuit := color.New(color.FgRed).Add(color.Bold)
	cTitle.Println(menuText)
	fmt.Println("==========================================")
	for i, option := range menuOptions {
		cMenu.Printf("%d. %s\n", i, option)
	}
	fmt.Println("==========================================")

	menuOption := 0
	fmt.Print("Enter Menu Option: ")
	fmt.Scan(&menuOption)
	return menuOption
}
