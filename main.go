package main

import (
	"archive/zip"
	"bufio"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
)

func main() {
	// 1. read docx file
	// 2. display docx file
	// 3. modify docx file
	// 4. save changed file
	// 1. read docx file as bytes
	// 2. return srialized data

	// make docx file readable
	unzipDocx := func(dist, filename string) []string {
		dst := dist
		archive, err := zip.OpenReader(filename)
		if err != nil {
			panic(err)
		}
		defer archive.Close()

		files := make([]string, 0)

		for _, f := range archive.File {
			filePath := filepath.Join(dst, f.Name)
			// fmt.Println("unzipping file ", filePath)

			if !strings.HasPrefix(filePath, filepath.Clean(dst)+string(os.PathSeparator)) {
				// fmt.Println("invalid file path")
				return files
			}
			if f.FileInfo().IsDir() {
				// fmt.Println("creating directory...")
				os.MkdirAll(filePath, os.ModePerm)
				continue
			}

			if err := os.MkdirAll(filepath.Dir(filePath), os.ModePerm); err != nil {
				panic(err)
			}

			dstFile, err := os.OpenFile(filePath, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, f.Mode())
			if err != nil {
				panic(err)
			}

			fileInArchive, err := f.Open()
			if err != nil {
				panic(err)
			}

			if _, err := io.Copy(dstFile, fileInArchive); err != nil {
				panic(err)
			}

			dstFile.Close()
			fileInArchive.Close()

			files = append(files, filePath)
			fmt.Println(filePath)
		}
		return files
	}
	unzipDocx("output", "document.docx")

	// open xml file
	xmlFile, err := os.Open("output/word/document.xml")
	if err != nil {
		fmt.Println(err)
	}

	defer xmlFile.Close()

	// create new xml file
	scanner := bufio.NewScanner(xmlFile)
	outputFile, err := os.Create("document.xml")
	if err != nil {
		fmt.Println(err)
	}
	defer outputFile.Close()

	for scanner.Scan() {
		text := scanner.Text()

		if strings.Contains(text, "leva.kondratev") {
			text = strings.Replace(text, "leva", "Anna", -1)
		}
		// fmt.Println(text)
		outputFile.WriteString(text)
	}

	// replace old xml with new one
	os.Remove("output/word/document.xml")
	os.Rename("document.xml", "output/word/document.xml")

	// Convert to docx
	// recursiveZip("output/", "output.docx")
	os.Remove("output/.DS_Store")
	recursiveZip("output", "doc.docx")

	os.RemoveAll("output")
}
