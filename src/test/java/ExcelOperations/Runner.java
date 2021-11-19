package ExcelOperations;

public class Runner {
    public static void main(String[] args) {
        ExcelOperations run = new Implementor();
        run.readTheExcelFile();
        run.checkTheGivenDataInTheGivenColumn();
        run.checkTheGivenDataInTheGivenRow();
        run.checkIfTheGivenHeaderIsPresentOrNot();
        run.getTheValuesFromTheGivenColumnNumber();
        run.getTheValuesFromTheGivenRowNumber();
        run.getTheNumberOfColumnsInTheFile();
        run.getTheNumberOfRowsInTheFile();
    }
}
