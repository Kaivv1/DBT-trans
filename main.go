package main

import (
	// "fmt"
	"fmt"
	"io/fs"
	"log"
	"os"
	"path/filepath"
	"strings"

	excel "github.com/xuri/excelize/v2"
)

func main() {
	currDir, err := os.Getwd()
	if err != nil {
		log.Println("Cant get current dir")
		return
	}
	parentDir := filepath.Dir(currDir)
	filename := "DBT_s_imeili.xlsx"
	var filePath string

	err = filepath.Walk(parentDir, func(path string, info fs.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if !info.IsDir() && strings.EqualFold(info.Name(), filename) {
			log.Printf("Found file: %s\n", info.Name())
			filePath = path
			log.Printf("Path to file is: %s", filePath)
			return filepath.SkipDir
		}
		return nil
	})
	if err != nil {
		log.Printf("Error walking path: %s\n", parentDir)
		return
	}
	excelFile, err := excel.OpenFile(filePath)
	if err != nil {
		log.Printf("Error opening %s: %s", filename, err.Error())
		return
	}
	defer func() {
		if err := excelFile.Close(); err != nil {
			log.Printf("Error while closing file(%s) with error: %s\n", filename, err.Error())
		}
	}()
	sheets := excelFile.GetSheetList()
	neededSheet := sheets[0]
	rows, err := excelFile.GetRows(neededSheet)
	if err != nil {
		log.Printf("Can't get rows on file: %s", filename)
		return
	}
	DBT_FOLDER := filepath.Join(parentDir, "ДБТ")

	if _, err = os.Stat(DBT_FOLDER); err == nil {
		log.Println("Folder exists")
		log.Fatalln("Script is being canceled")
		return
	} else if os.IsNotExist(err) {
		log.Println("Folder does not exist")
		log.Println("Folder is being created")
	} else {
		log.Println("Error searching for folder")
		log.Fatalln("Script is being canceled")
		return
	}

	err = os.MkdirAll(DBT_FOLDER, os.ModePerm)
	if err != nil {
		log.Println("Error creating DBT folder")
		return
	}
	log.Println("ДБТ folder created")

	emailFilePath := filepath.Join(DBT_FOLDER, "ДБТ-имейли.txt")

	emailsFile, err := os.OpenFile(emailFilePath, os.O_RDWR|os.O_CREATE|os.O_APPEND|os.O_WRONLY, 0666)
	emailsFileName := filepath.Base(emailFilePath)
	if err != nil {
		log.Fatalf("Cannot open %s\n", emailsFileName)
	}

	defer func() {
		if err := emailsFile.Close(); err != nil {
			log.Fatalf("Error closing %s\n", emailsFileName)
		}
	}()

	log.Println("ДБТ-имейли.txt created")

	for _, row := range rows {
		DBT := row[1]
		email := row[2]

		if DBT == "" || email == "" {
			continue
		}

		if email != "" {
			_, err := emailsFile.Write([]byte(fmt.Sprintf("%s - %s\n", DBT, email)))
			if err != nil {
				log.Printf("Error while adding email: %s for DBT: %s\n", email, DBT)
			} else {
				// log.Printf("Email: %s for %s is added in %s\n", email, DBT, emailsFileName)
			}
		}
		slice1 := strings.Split(email, "-")
		slice2 := strings.Split(slice1[1], "@")
		num := slice2[0]

		cityDBTPath := filepath.Join(DBT_FOLDER, fmt.Sprintf("%s-%s", num, DBT))
		err = os.Mkdir(cityDBTPath, os.ModePerm)
		folder := filepath.Base(cityDBTPath)
		if err != nil {
			log.Printf("Error creating: %s", folder)
		}
		log.Printf("Folder: %s created", folder)
	}
	AZ_Folder_Path := filepath.Join(DBT_FOLDER, "AZ")
	err = os.Mkdir(AZ_Folder_Path, os.ModePerm)
	if err != nil {
		log.Println("Error creating AZ folder")
	}
}
