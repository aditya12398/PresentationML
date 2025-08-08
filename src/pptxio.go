package main

import (
	"fmt"
	"github.com/moipa-cn/pptx"
	"regexp"
)

func main() {
	pptfile, _ := pptx.ReadPowerPoint("./test.pptx")
	fmt.Println("Presentation has been loaded successfully.")

	slideno := pptfile.GetSlideCount()
	fmt.Println("Total number of slides in the presentation: ", slideno)

	for slidepaths := range pptfile.Slides {
		fmt.Println(slidepaths)
	}
	allTags := getAllTags(pptfile.Slides, "(?s)<a(.*?)>")
	fmt.Println("Length of all the tags: ", len(allTags))
	for i := range allTags {
		fmt.Println("Tag ", i+1, ": ", allTags[i])
	}
}

func getAllTags(workstring map[string]string, pattern string) []string {
	var tags []string
	for key := range workstring {
		re := regexp.MustCompile(pattern)
		matches := re.FindAllStringSubmatch(workstring[key], -1)
		for _, m := range matches {
			tags = append(tags, m[0])
		}
		break
	}
	return tags
}
