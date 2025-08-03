package main

import (
	"fmt"
	"github.com/moipa-cn/pptx"
)

func main() {
	pptfile, _ := pptx.ReadPowerPoint("./test.pptx")
	fmt.Println("Just loaded it.")
	pptfile.DeletePassWord()
	fmt.Println("Password deleted.")
	slideno := pptfile.GetSlideCount()
	fmt.Println("Total number of slides in the presentation: ", slideno)
}
