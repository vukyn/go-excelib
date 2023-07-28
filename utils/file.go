package utils

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"path"
	"path/filepath"
)

func CreateFilePath(filePath string) error {
	path, _ := filepath.Split(filePath)
	if len(path) == 0 {
		return nil
	}

	_, err := os.Stat(path)
	if err != nil || os.IsExist(err) {
		err = os.MkdirAll(path, os.ModePerm)
	}
	return err
}

// Write text to a file.
// Create if not exist, or overwrite the existing file.
//
// Example:
//
//	WriteFile("Hello World", "temp/output.txt")
func WriteFile(input string, filePath string) error {
	data := []byte(input)
	dir, _ := filepath.Split(filePath)

	if _, err := os.Stat(dir); err == nil {
		os.Remove(filePath)
	} else {
		if err := os.MkdirAll(dir, os.ModePerm); err != nil {
			return err
		}
	}
	if err := os.WriteFile(filePath, data, 0); err != nil {
		return err
	}
	return nil
}

// Clear file removes the named files and directories, if any.
func ClearFile(files ...string) error {
	for _, file := range files {
		if err := os.Remove(file); err != nil {
			return err
		}
	}
	return nil
}

// ZipSingleFile zip the sourceFile path to outputFile path
func ZipSingleFile(sourceFile, outputFile string) error {
	if err := CreateFilePath(outputFile); err != nil {
		return err
	}

	archive, err := os.Create(outputFile)
	if err != nil {
		return err
	}
	defer archive.Close()

	zipWriter := zip.NewWriter(archive)
	defer zipWriter.Close()

	f, err := os.Open(sourceFile)
	if err != nil {
		return err
	}
	defer f.Close()

	w, err := zipWriter.Create(path.Base(f.Name()))
	if err != nil {
		return err
	}
	if _, err := io.Copy(w, f); err != nil {
		return err
	}

	return nil
}

// ZipMultipleFile zip list of sourceFile path to outputFile path
func ZipMultipleFile(outputFile string, sourceFiles ...string) error {
	flags := os.O_WRONLY | os.O_CREATE | os.O_TRUNC
	file, err := os.OpenFile(outputFile, flags, 0644)
	if err != nil {
		return fmt.Errorf("failed to open zip for writing: %s", err)
	}
	defer file.Close()

	zipw := zip.NewWriter(file)
	defer zipw.Close()

	for _, filename := range sourceFiles {
		if err := appendFiles(filename, zipw); err != nil {
			return fmt.Errorf("failed to add file %s to zip: %s", filename, err)
		}
	}
	return err
}

func appendFiles(filename string, zipw *zip.Writer) error {
	file, err := os.Open(filename)
	if err != nil {
		return fmt.Errorf("failed to open %s: %s", filename, err)
	}
	defer file.Close()

	wr, err := zipw.Create(filename)
	if err != nil {
		msg := "failed to create entry for %s in zip file: %s"
		return fmt.Errorf(msg, filename, err)
	}

	if _, err := io.Copy(wr, file); err != nil {
		return fmt.Errorf("failed to write %s to zip: %s", filename, err)
	}

	return nil
}
