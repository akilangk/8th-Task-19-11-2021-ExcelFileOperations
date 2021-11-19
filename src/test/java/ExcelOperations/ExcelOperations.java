package ExcelOperations;

interface ExcelOperations {
    void readTheExcelFile();

    void checkTheGivenDataInTheGivenColumn();

    void checkTheGivenDataInTheGivenRow();

    void checkIfTheGivenHeaderIsPresentOrNot();

    void getTheValuesFromTheGivenColumnNumber();

    void getTheValuesFromTheGivenRowNumber();

    void getTheNumberOfColumnsInTheFile();

    void getTheNumberOfRowsInTheFile();
}
