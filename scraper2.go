package main

import (
	"bufio"
	"fmt"
	"os"

	"github.com/gocolly/colly"
	// "github.com/playwright-community/playwright-go"
)

func scrapperColly(targetURL string) {
	fmt.Print("Input URL to scrape all text\n")
	// fmt.Scan(&targetURL)
	r := bufio.NewReader(os.Stdin)
	targetURL, _ = r.ReadString('\n')

	c := colly.NewCollector(
		colly.AllowedDomains(targetURL),
	)

	// called before an HTTP request is triggered
	c.OnRequest(func(r *colly.Request) {
		fmt.Println("Visiting: ", r.URL)
	})

	// triggered when the scraper encounters an error
	c.OnError(func(_ *colly.Response, err error) {
		fmt.Println("Something went wrong: ", err)
	})

	// fired when the server responds
	c.OnResponse(func(r *colly.Response) {
		fmt.Println("Page visited: ", r.Request.URL)
	})

	// triggered when a CSS selector matches an element
	c.OnHTML("a", func(e *colly.HTMLElement) {
		// printing all URLs associated with the <a> tag on the page
		fmt.Printf("%v", e.Attr("href"))
	})

	// triggered once scraping is done (e.g., write the data to a CSV file)
	c.OnScraped(func(r *colly.Response) {
		fmt.Println(r.Request.URL, " scraped!")
	})
	c.Visit(targetURL)

	fmt.Print("finished")
	rx := bufio.NewReader(os.Stdin)
	_, _ = rx.ReadString('\n')
}
