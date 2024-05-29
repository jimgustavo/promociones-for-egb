package main

import (
	"fmt"
	"log"
	"strconv"
	"time"

	"github.com/jung-kurt/gofpdf"
	"github.com/xuri/excelize/v2"
)

// Count the filled rows in the students list
func countFilledRows(f *excelize.File) (int, error) {
	// Find the last non-empty row in column F
	var numRows int
	for i := 2; i <= 50; i++ { // Starting from F2 to the last possible row, to maximum 50
		cellValue, err := f.GetCellValue("DATA", fmt.Sprintf("F%d", i))
		if err != nil {
			return 0, err
		}
		if cellValue == "" {
			break
		}
		numRows++
	}

	return numRows, nil
}

func getCellValueWithCheck(row []string, index int) string {
	if index < len(row) {
		return row[index]
	}
	return ""
}

func main() {
	// Open the Excel file
	f, err := excelize.OpenFile("consolidado.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	//Get the number of filled rows
	numRows, err := countFilledRows(f)
	if err != nil {
		log.Fatal(err)
	}

	// Create a new PDF
	pdf := gofpdf.New("L", "mm", "A4", "")

	//In case is needed to use special charaters such as: ñ, é, ó and so forth.
	tr := pdf.UnicodeTranslatorFromDescriptor("")

	// Get data from specific cells in the DATA sheet
	dataSheet := "DATA"
	var institution, class, school_year, workday, city, principal, secretary string

	institution, err = f.GetCellValue(dataSheet, "B2")
	if err != nil {
		log.Fatal(err)
	}
	class, err = f.GetCellValue(dataSheet, "B3")
	if err != nil {
		log.Fatal(err)
	}
	school_year, err = f.GetCellValue(dataSheet, "B6")
	if err != nil {
		log.Fatal(err)
	}
	workday, err = f.GetCellValue(dataSheet, "B8")
	if err != nil {
		log.Fatal(err)
	}
	city, err = f.GetCellValue(dataSheet, "B9")
	if err != nil {
		log.Fatal(err)
	}
	principal, err = f.GetCellValue(dataSheet, "B10")
	if err != nil {
		log.Fatal(err)
	}
	secretary, err = f.GetCellValue(dataSheet, "B11")
	if err != nil {
		log.Fatal(err)
	}

	// Loop through each student
	rows, err := f.GetRows("DATA")
	if err != nil {
		log.Fatal(err)
	}

	// Get the current date and time
	currentTime := time.Now()
	// Truncate the time to seconds
	truncatedTime := currentTime.Truncate(time.Second)

	for i, row := range rows[1 : numRows+1] { // Assuming student list starts from row F2 to the last filled row.
		// Extract student name
		studentName := getCellValueWithCheck(row, 5) //// Assuming student list starts in F2 row or 5 row

		// Add a new page for each student
		pdf.AddPage()

		// Add logo image
		logoPath := "ue12f_logo.jpeg"
		pdf.Image(logoPath, 10, 5, 20, 0, false, "", 0, "ue12f_logo")

		pdf.SetFont("Arial", "", 13)
		// Add title
		pdf.CellFormat(280, 10, institution, "0", 0, "C", false, 0, "")
		pdf.Ln(5)

		pdf.SetFont("Arial", "", 9)
		// Add city
		pdf.CellFormat(280, 10, city, "0", 0, "C", false, 0, "")
		pdf.Ln(5)
		// Add school year
		pdf.CellFormat(280, 10, school_year, "0", 0, "C", false, 0, "")
		pdf.Ln(10)

		pdf.SetFont("Arial", "", 13)
		// Add title
		pdf.CellFormat(280, 10, tr("CERTIFICADO DE PROMOCIÓN "), "0", 0, "C", false, 0, "")
		pdf.Ln(10)

		pdf.SetFont("Arial", "", 10)
		// Write specific data from the DATA sheet
		pdf.Cell(40, 10, tr("De conformidad con los prescrito en el Art. 187 del Reglamento General a la Ley Orgánica de Educación Intercultural y demas normativas vigentes, certifica que"))
		pdf.Ln(5)
		pdf.Cell(40, 10, tr("el/la estudiante: "+studentName+", paralelo "+class+", modalidad "+workday+", especialidad Ciencia Generales, obtuvo"))
		pdf.Ln(5)
		pdf.Cell(40, 10, tr("las siguientes calificaciones durante el presente año lectivo:"))
		pdf.Ln(5)

		// Extract grades for different subjects
		subjects := []string{"Matematicas", "Lenguaje", "CCNN", "EESS", "Ingles", "EEFF", "ECA"}
		grades := make(map[string][][]string)

		for _, subject := range subjects {
			grades[subject], err = f.GetRows(subject)
			if err != nil {
				log.Fatal(err)
			}
		}

		// Variable to track if the student is promoted
		isPromoted := true
		hasSupletorioGrades := false

		// Write the labels
		pdf.SetFont("Helvetica", "", 9) // Set font and size for the table
		pdf.Ln(5)
		pdf.Cell(60, 10, "")
		pdf.Cell(40, 10, tr("Calificación Anual"))
		pdf.Cell(40, 10, "Supletorio")
		pdf.Cell(40, 10, "Comportamiento")
		pdf.Ln(10)

		// Calculate the X coordinate for the start and end of the line
		startX := 175.0         // Right margin of the page
		endX := startX - 164    // Adjust as needed for the length of the line
		lineY := pdf.GetY() + 7 // Adjust as needed for the vertical position of the lines

		for _, subject := range subjects {
			// Write subject name
			switch subject {
			case "Matematicas":
				pdf.Cell(60, 10, "MATEMATICAS:")
			case "Lenguaje":
				pdf.Cell(60, 10, "LENGUAJE:")
			case "CCNN":
				pdf.Cell(60, 10, "CIENCIAS NATURALES:")
			case "EESS":
				pdf.Cell(60, 10, "ESTUDIOS SOCIALES:")
			case "Ingles":
				pdf.Cell(60, 10, "INGLES:")
			case "EEFF":
				pdf.Cell(60, 10, "CULTURA FISICA:")
			case "ECA":
				pdf.Cell(60, 10, "ECA:")
			}

			calificacionAnualStr := getCellValueWithCheck(grades[subject][i+6], 16)
			calificacionAnual, _ := strconv.ParseFloat(calificacionAnualStr, 64)

			supletorioStr := getCellValueWithCheck(grades[subject][i+6], 17)
			supletorio, err := strconv.ParseFloat(supletorioStr, 64)
			if err == nil {
				hasSupletorioGrades = true
			}

			// Check promotion status
			if calificacionAnual < 7 {
				if err != nil || supletorio < 7 {
					isPromoted = false
				}
			}

			if err == nil && supletorio < 7 {
				isPromoted = false
			}

			pdf.Cell(40, 10, calificacionAnualStr)
			pdf.Cell(40, 10, supletorioStr)
			pdf.Cell(40, 10, getCellValueWithCheck(grades[subject][i+6], 19))
			pdf.Ln(5)

			// Add a line below the row
			pdf.Line(startX, lineY, endX, lineY)
			//pdf.Ln(5) // Move to the next line
			// Update the Y-coordinate for the next line
			lineY = pdf.GetY() + 7
		}

		// Set font and size for the report closing section
		pdf.SetFont("Arial", "", 10)

		// Add closing paragraph with promotion status
		promotionStatus := "ES PROMOVIDO/A"
		if !isPromoted || (hasSupletorioGrades && !isPromoted) {
			promotionStatus = "NO ES PROMOVIDO/A"
		}

		// Add closing paragraph
		pdf.Cell(40, 20, "Por lo tanto el estudiante "+promotionStatus+". Para certificar, suscriben en unidad de acto el Rector/a con la/el Secretario/a General del Plantel el "+truncatedTime.Format("2006-01-02 15:04:05"))
		pdf.Ln(5)
		pdf.Cell(40, 10, "")
		pdf.Ln(5)

		// Add teacher signature
		pdf.Cell(75, 10, "")
		pdf.CellFormat(40, 45, "________________", "0", 0, "C", false, 0, "")
		pdf.CellFormat(70, 45, "________________", "0", 0, "C", false, 0, "") //In case it's needed authority signature
		pdf.Ln(5)
		pdf.Cell(75, 10, "")
		pdf.CellFormat(40, 45, principal, "0", 0, "C", false, 0, "")
		pdf.CellFormat(70, 45, secretary, "0", 0, "C", false, 0, "") //In case it's needed authority signature
		pdf.Ln(5)
		pdf.Cell(75, 10, "")
		pdf.CellFormat(40, 45, "RECTOR(A)", "0", 0, "C", false, 0, "")
		pdf.CellFormat(70, 45, "SECRETARIA(O)", "0", 0, "C", false, 0, "") //In case it's needed authority signature
	}

	// Save the PDF
	err = pdf.OutputFileAndClose("promociones.pdf")
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println("Report generated successfully.")
}

/*
package main

import (
	"fmt"
	"log"
	"time"

	"github.com/jung-kurt/gofpdf"
	"github.com/xuri/excelize/v2"
)

// Count the filled rows in the students list
func countFilledRows(f *excelize.File) (int, error) {
	// Find the last non-empty row in column F
	var numRows int
	for i := 2; i <= 50; i++ { // Starting from F2 to the last possible row, to maximum 50
		cellValue, err := f.GetCellValue("DATA", fmt.Sprintf("F%d", i))
		if err != nil {
			return 0, err
		}
		if cellValue == "" {
			break
		}
		numRows++
	}

	return numRows, nil
}

func main() {
	// Open the Excel file
	f, err := excelize.OpenFile("consolidado.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	//Get the number of filled rows
	numRows, err := countFilledRows(f)
	if err != nil {
		log.Fatal(err)
	}
	//fmt.Println("Number of filled rows:", numRows)   //Print the number of filled rows in column F to the console

	// Create a new PDF
	pdf := gofpdf.New("L", "mm", "A4", "")

	//In case is needed to use special charaters such as: ñ, é, ó and so forth.
	tr := pdf.UnicodeTranslatorFromDescriptor("")

	// Get data from specific cells in the DATA sheet
	dataSheet := "DATA"
	var institution, class, school_year, workday, city, principal, secretary string

	institution, err = f.GetCellValue(dataSheet, "B2")
	if err != nil {
		log.Fatal(err)
	}
	class, err = f.GetCellValue(dataSheet, "B3")
	if err != nil {
		log.Fatal(err)
	}
	school_year, err = f.GetCellValue(dataSheet, "B6")
	if err != nil {
		log.Fatal(err)
	}
	workday, err = f.GetCellValue(dataSheet, "B8")
	if err != nil {
		log.Fatal(err)
	}
	city, err = f.GetCellValue(dataSheet, "B9")
	if err != nil {
		log.Fatal(err)
	}
	principal, err = f.GetCellValue(dataSheet, "B10")
	if err != nil {
		log.Fatal(err)
	}
	secretary, err = f.GetCellValue(dataSheet, "B11")
	if err != nil {
		log.Fatal(err)
	}

	// Loop through each student
	rows, err := f.GetRows("DATA")
	if err != nil {
		log.Fatal(err)
	}

	// Get the current date and time
	currentTime := time.Now()
	// Truncate the time to seconds
	truncatedTime := currentTime.Truncate(time.Second)

	for i, row := range rows[1 : numRows+1] { // Assuming student list starts from row F2 to the last filled row.
		// Extract student name
		studentName := row[5] //// Assuming student list starts in F2 row or 5 row

		// Add a new page for each student
		pdf.AddPage()

		// Add logo image
		logoPath := "ue12f_logo.jpeg"
		pdf.Image(logoPath, 10, 5, 20, 0, false, "", 0, "ue12f_logo")

		pdf.SetFont("Arial", "", 13)
		// Add title
		pdf.CellFormat(280, 10, institution, "0", 0, "C", false, 0, "")
		pdf.Ln(5)

		pdf.SetFont("Arial", "", 9)
		// Add city
		pdf.CellFormat(280, 10, city, "0", 0, "C", false, 0, "")
		pdf.Ln(5)
		// Add school year
		pdf.CellFormat(280, 10, school_year, "0", 0, "C", false, 0, "")
		pdf.Ln(10)

		pdf.SetFont("Arial", "", 13)
		// Add title
		pdf.CellFormat(280, 10, tr("CERTIFICADO DE PROMOCIÓN "), "0", 0, "C", false, 0, "")
		pdf.Ln(10)

		pdf.SetFont("Arial", "", 10)
		// Write specific data from the DATA sheet
		pdf.Cell(40, 10, tr("De conformidad con los prescrito en el Art. 187 del Reglamento General a la Ley Orgánica de Educación Intercultural y demas normativas vigentes, certifica que"))
		pdf.Ln(5)
		pdf.Cell(40, 10, tr("el/la estudiante: "+studentName+", paralelo "+class+", modalidad "+workday+", especialidad Ciencia Generales, obtuvo"))
		pdf.Ln(5)
		pdf.Cell(40, 10, tr("las siguientes calificaciones durante el presente año lectivo:"))
		pdf.Ln(5)

		// Extract math grades
		mathGrades, err := f.GetRows("Matematicas")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		languageGrades, err := f.GetRows("Lenguaje")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		scienceGrades, err := f.GetRows("CCNN")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		social_studiesGrades, err := f.GetRows("EESS")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		englishGrades, err := f.GetRows("Ingles")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		physical_cultureGrades, err := f.GetRows("EEFF")
		if err != nil {
			log.Fatal(err)
		}

		// Extract science grades
		art_cultureGrades, err := f.GetRows("ECA")
		if err != nil {
			log.Fatal(err)
		}

		// Write the labels
		pdf.SetFont("Helvetica", "", 9) // Set font and size for the table
		pdf.Ln(5)
		pdf.Cell(60, 10, "")
		pdf.Cell(40, 10, tr("Calificación Anual"))
		pdf.Cell(40, 10, "Supletorio")
		pdf.Cell(40, 10, "Comportamiento")
		pdf.Ln(10)

		// Write math grades
		pdf.Cell(60, 10, "MATEMATICAS:")

		pdf.Cell(40, 10, mathGrades[i+6][16])
		pdf.Cell(40, 10, mathGrades[i+6][17])
		pdf.Cell(40, 10, mathGrades[i+6][19])
		pdf.Ln(5)

		// Write language grades
		pdf.Cell(60, 10, "LENGUAJE:")

		pdf.Cell(40, 10, languageGrades[i+6][16])
		pdf.Cell(40, 10, languageGrades[i+6][17])
		pdf.Cell(40, 10, languageGrades[i+6][19])
		pdf.Ln(5)

		// Write science grades
		pdf.Cell(60, 10, "CIENCIAS NATURALES:")

		pdf.Cell(40, 10, scienceGrades[i+6][16])
		pdf.Cell(40, 10, scienceGrades[i+6][17])
		pdf.Cell(40, 10, scienceGrades[i+6][19])
		pdf.Ln(5)

		// Write social studies grades
		pdf.Cell(60, 10, "ESTUDIOS SOCIALES:")

		pdf.Cell(40, 10, social_studiesGrades[i+6][16])
		pdf.Cell(40, 10, social_studiesGrades[i+6][17])
		pdf.Cell(40, 10, social_studiesGrades[i+6][19])
		pdf.Ln(5)

		// Write english grades
		pdf.Cell(60, 10, "INGLES:")

		pdf.Cell(40, 10, englishGrades[i+6][16])
		pdf.Cell(40, 10, englishGrades[i+6][17])
		pdf.Cell(40, 10, englishGrades[i+6][19])
		pdf.Ln(5)

		// Write physical culture grades
		pdf.Cell(60, 10, "CULTURA FISICA:")

		pdf.Cell(40, 10, physical_cultureGrades[i+6][16])
		pdf.Cell(40, 10, physical_cultureGrades[i+6][17])
		pdf.Cell(40, 10, physical_cultureGrades[i+6][19])
		pdf.Ln(5)

		// Write art culture grades
		pdf.Cell(60, 10, "ECA:")

		pdf.Cell(40, 10, art_cultureGrades[i+6][16])
		pdf.Cell(40, 10, art_cultureGrades[i+6][17])
		pdf.Cell(40, 10, art_cultureGrades[i+6][19])
		pdf.Ln(10)

		// Set font and size for the report closing section
		pdf.SetFont("Arial", "", 10)

		// Add closing paragraph
		pdf.Cell(40, 10, "Por lo tanto el estudiante ES PROMOVIDO/A. Para certificar, suscriben en unidad de acto el Rector/a con la/el Secretario/a General del Plantel el "+truncatedTime.Format("2006-01-02 15:04:05"))
		pdf.Ln(5)
		pdf.Cell(40, 10, "")
		pdf.Ln(5)

		// Add teacher signature
		pdf.Cell(75, 10, "")
		pdf.CellFormat(40, 45, "________________", "0", 0, "C", false, 0, "")
		pdf.CellFormat(70, 45, "________________", "0", 0, "C", false, 0, "") //In case it'needed authority signature
		pdf.Ln(5)
		pdf.Cell(75, 10, "")
		pdf.CellFormat(40, 45, principal, "0", 0, "C", false, 0, "")
		pdf.CellFormat(70, 45, secretary, "0", 0, "C", false, 0, "") //In case it'needed authority signature
		pdf.Ln(5)
		pdf.Cell(75, 10, "")
		pdf.CellFormat(40, 45, "RECTOR(A)", "0", 0, "C", false, 0, "")
		pdf.CellFormat(70, 45, "SECRETARIA(O)", "0", 0, "C", false, 0, "") //In case it'needed authority signature
	}

	// Save PDF to files
	err = pdf.OutputFileAndClose("promociones.pdf")
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("Report generated successfully.")
}
*/
//------------For Linux-------------------//
//GOOS=windows GOARCH=amd64 go build -o reportes-individuales    //Replace linux with windows or darwin depending on the target platform. Replace amd64 with other architectures if needed.

//------------For Windows----------------//
//set GOOS=windows
//set GOARCH=amd64
//go build -o reportes-individuales.exe
