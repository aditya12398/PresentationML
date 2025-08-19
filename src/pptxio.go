package main

import (
	"fmt"
	"github.com/moipa-cn/pptx"
	"os"
	"regexp"
	"strconv"
	"strings"
)

func main() {
	pptfile, _ := pptx.ReadPowerPoint("./test.pptx")
	fmt.Println("Presentation has been loaded successfully.")

	slideno := pptfile.GetSlideCount()
	fmt.Println("Total number of slides in the presentation: ", slideno)

	for slidepaths := range pptfile.Slides {
		fmt.Println(slidepaths)
	}

	for key, content := range pptfile.Slides {
		// Get the basename of the filepath from key
		filename := key[strings.LastIndex(key, "/")+1 : strings.LastIndex(key, ".")]
		fmt.Printf("Processing slide: %s\n", filename)

		allTags := getAllTags(content, "(?s)<[ap](.*?)>") // This regex matches all tags starting with 'a' or 'p'
		f1, err1 := os.Create("../data/p_" + filename + ".dat")
		f2, err2 := os.Create("../data/a_" + filename + ".dat")
		if err1 != nil || err2 != nil {
			fmt.Println("Error creating output files:", err1, err2)
			return
		}
		defer f1.Close()
		defer f2.Close()

		fmt.Printf("Parsing regex pattern (?s)<[ap](.*?)> for slide: %s\n", filename)
		// fmt.Println("Length of all the tags: ", len(allTags))
		for i := range allTags {
			// fmt.Println("Tag ", i+1, ": ", allTags[i])
			f1.WriteString(strconv.Itoa(i+1) + ": " + allTags[i] + "\n")
		}

		tagsWithFeatures := tagsWFeatures(content, "(?s)<[^<>]*?/>") // This regex matches self-closing tags
		fmt.Printf("Parsing regex pattern (?s)<[^<>]*?/> for slide: %s\n", filename)
		for i := range tagsWithFeatures {
			// fmt.Println("Tag with features ", i+1, ": ", tagsWithFeatures[i])
			f2.WriteString(strconv.Itoa(i+1) + ": " + tagsWithFeatures[i] + "\n")
		}
	}

}

func getAllTags(content string, pattern string) []string {
	var tags []string
	re := regexp.MustCompile(pattern)
	matches := re.FindAllStringSubmatch(content, -1)
	for _, m := range matches {
		tags = append(tags, m[0])
	}
	return tags
}

// This function extracts self closing tags that usually contain features like font size, color, etc.
func tagsWFeatures(content string, pattern string) []string {
	var tags []string
	re := regexp.MustCompile(pattern)
	matches := re.FindAllStringSubmatch(content, -1)
	for _, m := range matches {
		tags = append(tags, m[0])
	}
	return tags
}
