package ExcelOperations;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.InputMismatchException;
import java.util.List;
import java.util.Scanner;

class Implementor extends DataProvider implements ExcelOperations {
    String userDirectory = System.getProperty("user.dir");
    String pathSeparator = System.getProperty("file.separator");
    String excelFilePath = userDirectory + pathSeparator + "src" + pathSeparator + "files" + pathSeparator + "StudentRecords.xlsx";
    Scanner scannerObject = new Scanner(System.in);

    public void readTheExcelFile() {
        List<String> headerNames = new ArrayList<>();
        List<String> nameOfTheStudents = new ArrayList<>();
        List<Integer> ageOfTheStudents = new ArrayList<>();
        List<Integer> totalMarksOfTheStudents = new ArrayList<>();
        int numberOfRows;
        int numberOfColumns;
        try {
            File fileObject = new File(excelFilePath);
            FileInputStream fileInputStreamObject = new FileInputStream(fileObject);
            System.out.println("Reading Excel File..");
            XSSFWorkbook workBook = new XSSFWorkbook(fileInputStreamObject);
            XSSFSheet sheet = workBook.getSheetAt(0);
            numberOfRows = sheet.getPhysicalNumberOfRows();
            Row firstRow = sheet.getRow(0);
            numberOfColumns = firstRow.getPhysicalNumberOfCells();
            for (Cell cellData : firstRow) {
                headerNames.add(cellData.toString());
            }
            for (int rowIndex = 1; rowIndex < numberOfRows; rowIndex++) {
                Row individualRecordOfStudent = sheet.getRow(rowIndex);
                for (int columnIndex = 0; columnIndex < numberOfColumns; columnIndex++) {
                    if (columnIndex == 0) {
                        Cell columnData = individualRecordOfStudent.getCell(columnIndex);
                        nameOfTheStudents.add(columnData.toString());
                    } else if (columnIndex == 1) {
                        Cell columnData = individualRecordOfStudent.getCell(columnIndex);
                        ageOfTheStudents.add(Integer.parseInt(columnData.toString().replace(".0", "")));
                    } else if (columnIndex == 2) {
                        Cell columnData = individualRecordOfStudent.getCell(columnIndex);
                        totalMarksOfTheStudents.add(Integer.parseInt(columnData.toString().replace(".0", "")));
                    }
                }
            }
            System.out.println("Read Excel file successfully.");
            setNumberOfRows(numberOfRows);
            setNumberOfColumns(numberOfColumns);
            setHeaderNames(headerNames);
            setNameOfTheStudents(nameOfTheStudents);
            setAgeOfTheStudents(ageOfTheStudents);
            setTotalMarksOfTheStudents(totalMarksOfTheStudents);
        } catch (IOException exception) {
            System.out.println("Check the file in the specified path.");
        }
    }

    public void checkTheGivenDataInTheGivenColumn() {
        try {
            System.out.println("\nColumn Data Checker....");
            System.out.println("Enter the column number from 1 to 3, in which you want to check the value:");
            int columnNumber = scannerObject.nextInt();
            if (columnNumber == 1) {
                System.out.println("Enter the name, which you want to check in the column " + columnNumber);
                String nameToCheck = scannerObject.next();
                int count = 0;
                for (String nameData : getNameOfTheStudents()) {
                    if (nameData.equalsIgnoreCase(nameToCheck)) {
                        count++;
                        break;
                    }
                }
                if (count > 0) {
                    System.out.println("The given name " + nameToCheck + " is present in the column " + columnNumber);
                } else {
                    System.out.println("The given name " + nameToCheck + " is not present in the column " + columnNumber);
                }
            } else if (columnNumber == 2) {
                System.out.println("Enter the age, which you want to check in the column " + columnNumber);
                int ageToCheck = scannerObject.nextInt();
                int count = 0;
                for (int ageData : getAgeOfTheStudents()) {
                    if (ageData == ageToCheck) {
                        count++;
                        break;
                    }
                }
                if (count > 0) {
                    System.out.println("The given age " + ageToCheck + " is present in the column " + columnNumber);
                } else {
                    System.out.println("The given age " + ageToCheck + " is not present in the column " + columnNumber);
                }
            } else if (columnNumber == 3) {
                System.out.println("Enter the total, which you want to check in the column " + columnNumber);
                int totalMarksToCheck = scannerObject.nextInt();
                int count = 0;
                for (int marksData : getTotalMarksOfTheStudents()) {
                    if (marksData == totalMarksToCheck) {
                        count++;
                        break;
                    }
                }
                if (count > 0) {
                    System.out.println("The given total marks " + totalMarksToCheck + " is present in the column " + columnNumber);
                } else {
                    System.out.println("The given total marks " + totalMarksToCheck + " is not present in the column " + columnNumber);
                }
            } else {
                System.out.println("Oops, enter the column number from 1 to 3.");
            }
        } catch (InputMismatchException exception) {
            System.out.println("Oops, enter the current format as input");
        }
    }

    public void checkTheGivenDataInTheGivenRow() {
        try {
            System.out.println("\nRow Data Checker....");
            System.out.println("Enter the row number from 1 to 15, in which you want to check the value:");
            int rowNumber = scannerObject.nextInt();
            if (rowNumber > 0 && rowNumber <= 15) {
                System.out.println("Enter the data which you want to the check in the row " + rowNumber);
                String dataToCheck = scannerObject.next();
                String nameData = getNameOfTheStudents().get(rowNumber - 1);
                String ageData = getAgeOfTheStudents().get(rowNumber - 1).toString();
                String markData = getTotalMarksOfTheStudents().get(rowNumber - 1).toString();
                if (dataToCheck.equalsIgnoreCase(nameData) || dataToCheck.equals(ageData) || dataToCheck.equals(markData)) {
                    System.out.println("The given data " + dataToCheck + " is present in the row " + rowNumber);
                } else {
                    System.out.println("The given data " + dataToCheck + " is not present in the row " + rowNumber);
                }
            } else {
                System.out.println("Oops, enter the row number from 1 to 15.");
            }
        } catch (InputMismatchException exception) {
            System.out.println("Oops, enter the current format as input");
        }
    }

    public void checkIfTheGivenHeaderIsPresentOrNot() {
        System.out.println("\nHeader Checker....");
        System.out.println("Enter the header name of column 1, which you want to check:");
        String headerColumnOne = scannerObject.next();
        System.out.println("Enter the header name of column 2, which you want to check:");
        String headerColumnTwo = scannerObject.next();
        System.out.println("Enter the header name of column 3, which you want to check:");
        String headerColumnThree = scannerObject.next();
        if (headerColumnOne.equals(getHeaderNames().get(0)) && headerColumnTwo.equals(getHeaderNames().get(1)) && headerColumnThree.equals(getHeaderNames().get(2))) {
            System.out.println("The given header matches with the actual header in the excel file.");
        } else {
            System.out.println("Header Mismatch.");
        }
    }

    public void getTheValuesFromTheGivenColumnNumber() {
        try {
            System.out.println("\nColumn Value Getter....");
            System.out.println("Enter the column number from 1 to 3, in which you want to get the values:");
            int columnNumber = scannerObject.nextInt();
            if (columnNumber == 1) {
                System.out.println(getHeaderNames().get(0));
                for (String nameData : getNameOfTheStudents()) {
                    System.out.println(nameData);
                }
            } else if (columnNumber == 2) {
                System.out.println(getHeaderNames().get(1));
                for (int ageData : getAgeOfTheStudents()) {
                    System.out.println(ageData);
                }
            } else if (columnNumber == 3) {
                System.out.println(getHeaderNames().get(2));
                for (int marksData : getTotalMarksOfTheStudents()) {
                    System.out.println(marksData);
                }
            } else {
                System.out.println("Oops, enter the column number from 1 to 3.");
            }
        } catch (InputMismatchException exception) {
            System.out.println("Oops, enter the current format as input");
        }
    }

    public void getTheValuesFromTheGivenRowNumber() {
        try {
            System.out.println("\nRow Value Getter....");
            System.out.println("Enter the row number from 1 to 15, in which you want to get the values:");
            int rowNumber = scannerObject.nextInt();
            if (rowNumber > 0 && rowNumber <= 15) {
                String nameData = getNameOfTheStudents().get(rowNumber - 1);
                String ageData = getAgeOfTheStudents().get(rowNumber - 1).toString();
                String markData = getTotalMarksOfTheStudents().get(rowNumber - 1).toString();
                System.out.println(getHeaderNames().get(0) + " : " + nameData);
                System.out.println(getHeaderNames().get(1) + " : " + ageData);
                System.out.println(getHeaderNames().get(2) + " : " + markData);
            } else {
                System.out.println("Oops, enter the row number from 1 to 15.");
            }
        } catch (InputMismatchException exception) {
            System.out.println("Oops, enter the current format as input");
        }
    }

    public void getTheNumberOfColumnsInTheFile() {
        System.out.println();
        System.out.println("The total number of Columns in the given excel file: " + getNumberOfColumns());
    }

    public void getTheNumberOfRowsInTheFile() {
        System.out.println();
        System.out.println("The total number of Rows in the given excel file: " + getNumberOfRows());
    }
}
