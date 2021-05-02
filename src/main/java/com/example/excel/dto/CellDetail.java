package com.example.excel.dto;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class CellDetail {

    private String columnName;
    private String cellValue;

    public CellDetail(Sheet sheet, Cell cell, Integer currentColumnIndex) {
        this.cellValue = this.getValue(cell);
        Row row = sheet.getRow(0);
        this.columnName = row.getCell(currentColumnIndex).getRichStringCellValue().toString();
    }

    public String getColumnName() {
        return this.columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }

    public String getCellValue() {
        return this.cellValue;
    }

    public void setCellValue(String cellValue) {
        this.cellValue = cellValue;
    }

    private String getValue(Cell cell) {
        String value = "";
        switch (cell.getCellType()) {
            case BOOLEAN:
                value = String.valueOf(cell.getBooleanCellValue());
                break;
            case ERROR:
                value = String.valueOf(cell.getErrorCellValue());
                break;
            case FORMULA:
                value = String.valueOf(cell.getCellFormula());
                break;
            case STRING:
                value = String.valueOf(cell.getStringCellValue());
                break;
            case NUMERIC:
                value = String.valueOf(cell.getNumericCellValue());
                break;
            default:
                value = "";
        }
        return value;
    }

}
