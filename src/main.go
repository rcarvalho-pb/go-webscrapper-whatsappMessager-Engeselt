package main

import (
	"fmt"
	"strconv"
	"time"
	"os"

	"github.com/tebeka/selenium"
	"github.com/tebeka/selenium/chrome"
	"github.com/xuri/excelize/v2"
)

// var CONSTANTS = map[string] string {
// 	"driverPath": "./chromedriver/chromedriver.exe",
// 	"sheetPath": "./modelo_importacao_lote.xlsx",
// 	"sheetName": "Solicitação",
// 	"whatsappURL": "https://web.whatsapp.com",
// 	"sideElementWhats": "//*[@id='side']",
// 	"notFoundContact": "//*[@id='app']/div/span[2]/div/span/div/div/div/div/div/div[1]",
// 	"inputMessage": "//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]",
// 	"sendButton": "//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]/div[2]/button",
// 	"optionsButton": "//*[@id='app']/div/div/div[4]/header/div[2]/div/span/div[4]/div",
// 	"disconnectButton": "//*[@id='app']/div/div/div[4]/header/div[2]/div/span/div[4]/span/div/ul/li[6]/div",
// 	"confirmDisconnectButton": "//*[@id='app']/div/span[2]/div/div/div/div/div/div/div[3]/div/button[2]",
// }

const logFile = "./log.xlsx"
const sheetPath = "./modelo_importacao_lote.xlsx"
const driverPath = "./chromedriver/chromedriver.exe"
const sheetName = "Solicitação"
const whatsappURL = "https://web.whatsapp.com"
const sideElementWhats = "//*[@id='side']"
const notFoundContact = "//*[@id='app']/div/span[2]/div/span/div/div/div/div/div/div[1]"
const inputMessage = "//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]"
const sendButton = "//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]/div[2]/button"
const optionsButton = "//*[@id='app']/div/div/div[4]/header/div[2]/div/span/div[4]/div"
const disconnectButton = "//*[@id='app']/div/div/div[4]/header/div[2]/div/span/div[4]/span/div/ul/li[6]/div"
const confirmDisconnectButton = "//*[@id='app']/div/span[2]/div/div/div/div/div/div/div[3]/div/button[2]"

type Info struct {
	link string
	os string
}

type Log struct {
	os string
	contact string
}

func main () {
	WhatsService()
}

func WhatsService() {
	fmt.Println("Starting the program")

	service, err := selenium.NewChromeDriverService(driverPath, 4444)
	if err != nil {
		fmt.Println ("Error: ", err)
	}

	browser := SendMessages(service)
	browser.Close()

	fmt.Println("Finishing the program")
}

func SendMessages(service *selenium.Service) selenium.WebDriver {
	fmt.Println("Starting the Sending Messages")

	var logs []Log

	logFile := CreateExcelFile()

	browser := GetWebDriver()
	
	browser.Get("https://www.google.com")
	browser.Get(whatsappURL)

	links := GetLinks()	

	if IsElementLoaded(browser, selenium.ByXPATH, sideElementWhats) {
		for i, link := range links {
			fmt.Printf("Enviando a %dª mensagem...\n", i + 1)

			browser.Get(link.link)
			
			if IsElementLoaded(browser, selenium.ByXPATH, notFoundContact) {
				fmt.Printf("OS %s - Não foi possível realizar o contato\n", link.os)
				logs = append(logs, Log{os: link.os, contact: "NÃO"})
				continue
			}
			if IsElementLoaded(browser, selenium.ByXPATH, inputMessage) {
				time.Sleep(5 * time.Second)
				EnterKey(browser)
				fmt.Printf("OS %s - Contato realizado com sucesso.\n", link.os)
				logs = append(logs, Log{os: link.os, contact: "SIM"})
			}

			if (i == len(links) - 1) {
				fmt.Printf("last message sended.\n")
				} else {
				fmt.Printf("%dª message sended.\n", i + 1)
			}
			time.Sleep(5 * time.Second)
		}

	} else {
		fmt.Println("Whatsapp couldn't open. Try Again.")
		return browser
	}

	if WriteLog(logs, logFile) {
		if err := logFile.SaveAs("log.xlsx"); err != nil {
			fmt.Println("Error while saving file")
		}
	}

	if err := logFile.SaveAs("log.xlsx"); err != nil {
		fmt.Println("Error while saving file")
	}
	CloseWhats(browser)
	fmt.Println("Finishing send messages")
	return browser
}

func WriteLog(logs []Log, file *excelize.File) bool {
	for i, log := range logs {
		err := file.SetCellStr("Sheet1", fmt.Sprintf("A%d", i+2), log.os)
		err = file.SetCellStr("Sheet1", fmt.Sprintf("B%d", i+2), log.contact)
		if err != nil {
			fmt.Println("Error while writing file")
			return false
		}
	}
	return true
}

func GetWebDriver() (selenium.WebDriver) {
	fmt.Println("Getting webDriver")

	caps := selenium.Capabilities {}
	caps.AddChrome (chrome.Capabilities{Args: []string {
		"window-size=1920x1080",
			"--no-sandbox",
			"--disable-dev-shm-usage",
			"disable-gpu",
			// "--headless",  // comment out this line to see the browser
	}})

	driver, err := selenium.NewRemote(caps, "")
	if err != nil {
		fmt.Println("Error: ", err)
	}

	fmt.Println("Returning webDriver")
	return driver
}

func IsElementLoaded(driver selenium.WebDriver, typeSearch,  xpath string)  bool {
	xpathName := GetXpathName(xpath)
	
	limitTime := GetLimitTime(xpath)
	fmt.Println("Time:", limitTime)
	
	fmt.Printf("Testing if element %s is loaded\n", xpathName)


	flag := 0
	element, err := driver.FindElements(typeSearch, xpath)
	if err != nil {
		fmt.Println("Error: ", err)
	}
	for len(element) < 1 {
		element, err = driver.FindElements(typeSearch, xpath)
		if err != nil {
			fmt.Println("Error: ", err)
		}
		flag++
		fmt.Print(".")
		time.Sleep(1*time.Second)

		if flag == limitTime || len(element) >= 1 {
			break
		}
	}
	if !(len(element) < 1) {
		fmt.Printf("\nElement %s is loaded\n", xpathName)
		return true
	}

	fmt.Printf("\nElement %s isn't loaded\n", xpathName)
	return false
}

func GetLimitTime(xpathName string) int {
	switch xpathName {
		case sideElementWhats:
			return 60
		case notFoundContact:
			return 15
		default:
			return 10
	}
}

func GetXpathName(xpath string) string {

	switch xpath {
		case sideElementWhats:
			return "sideElementWhats"
		case notFoundContact:
			return "notFoundContact"
		case inputMessage:
			return "inputMessage"
		case sendButton:
			return "sendButton"
		case optionsButton:
			return "optionsButton"
		case disconnectButton:
			return "disconnectButton"
		case confirmDisconnectButton:
			return "confirmDisconnectButton"
		default:
			return "UnknownElement"
	}
}

func EnterKey(browser selenium.WebDriver) {
	fmt.Println("KeyDown Enter started")

	enter, err := browser.FindElement(selenium.ByXPATH, sendButton)
	if err != nil {
		fmt.Println("Error: ", err)
	}

	time.Sleep(2 * time.Second)

	fmt.Println("Enter Key binding")
	enter.SendKeys(string('\ue007'))
	fmt.Println("KeyDown Enter finished")
}

func CloseWhats(browser selenium.WebDriver) {
	fmt.Println("Closing Web Whatsapp...")

	menu, err := browser.FindElement(selenium.ByXPATH, optionsButton)
	if err != nil {
		fmt.Println("Error: ", err)
	}
	menu.Click()
	time.Sleep(2 * time.Second)
	logout, err := browser.FindElement(selenium.ByXPATH, disconnectButton)
	if err != nil {
		fmt.Println("Error: ", err)
	}
	logout.Click()
	time.Sleep(2 * time.Second)
	getOut, err := browser.FindElement(selenium.ByXPATH, confirmDisconnectButton)
	if err != nil {
		fmt.Println("Error: ", err)
	}
	getOut.Click()

	time.Sleep(5 * time.Second)

	fmt.Println("Web Whatsapp closed")
}

func GetLinks() []Info {
	fmt.Println("Getting links")
	var res []Info
	rows := GetSheet(sheetPath, sheetName)
	fmt.Println("Listing rows")

	for i, row := range rows {
		if i != 0{
			name := row[4]
			phone := row[12]
			address := row[10]
			os := row[0]

			msg := strconv.Quote(fmt.Sprintf("Olá, %s. Tudo bem? Queria confirmar se o seu endereço realmente é %s", name, address))

			link := fmt.Sprintf("https://web.whatsapp.com/send?phone=%s&text=%s", phone, msg)
			info := Info{link, os}
			res = append(res, info)
		}
	}

	fmt.Println("Returning links")
	return res
}

func GetSheet(sheetPath, sheetName string) ([][]string) {
	fmt.Println("Getting sheets")

	f, err := excelize.OpenFile(sheetPath)
	if err != nil {
		fmt.Println("Error: ", err)
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		fmt.Println("Error: ", err)
	}

	fmt.Println("Resturning sheets")
	return rows
}

func CreateExcelFile() (log *excelize.File) {
	if _, err := os.Stat(logFile); err != nil {
		if os.Remove(logFile) != nil {
			fmt.Println("Error while deleting log file")
		}
	}
	
	log = excelize.NewFile()
	defer func() {
		if err := log.Close(); err != nil {
			fmt.Println("Error while closing log file. Error: ", err)
		}
	}()

	log.SetCellValue("Sheet1", "A1", "Ordem de Serviço")
	log.SetCellValue("Sheet1", "B1", "Contato com Cliente")

	return log
}