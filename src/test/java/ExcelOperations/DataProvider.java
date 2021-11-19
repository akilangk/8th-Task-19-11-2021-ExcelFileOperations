package ExcelOperations;

import java.util.ArrayList;
import java.util.List;

class DataProvider {
    int numberOfRows;
    int numberOfColumns;
    List<String> headerNames = new ArrayList<>();
    List<String> nameOfTheStudents = new ArrayList<>();
    List<Integer> ageOfTheStudents = new ArrayList<>();
    List<Integer> totalMarksOfTheStudents = new ArrayList<>();

    public int getNumberOfRows() {
        return numberOfRows;
    }

    public void setNumberOfRows(int numberOfRows) {
        this.numberOfRows = numberOfRows;
    }

    public int getNumberOfColumns() {
        return numberOfColumns;
    }

    public void setNumberOfColumns(int numberOfColumns) {
        this.numberOfColumns = numberOfColumns;
    }

    public List<String> getHeaderNames() {
        return headerNames;
    }

    public void setHeaderNames(List<String> headerNames) {
        this.headerNames = headerNames;
    }

    public List<String> getNameOfTheStudents() {
        return nameOfTheStudents;
    }

    public void setNameOfTheStudents(List<String> nameOfTheStudents) {
        this.nameOfTheStudents = nameOfTheStudents;
    }

    public List<Integer> getAgeOfTheStudents() {
        return ageOfTheStudents;
    }

    public void setAgeOfTheStudents(List<Integer> ageOfTheStudents) {
        this.ageOfTheStudents = ageOfTheStudents;
    }

    public List<Integer> getTotalMarksOfTheStudents() {
        return totalMarksOfTheStudents;
    }

    public void setTotalMarksOfTheStudents(List<Integer> totalMarksOfTheStudents) {
        this.totalMarksOfTheStudents = totalMarksOfTheStudents;
    }
}
