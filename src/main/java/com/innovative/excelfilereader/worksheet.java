package com.innovative.excelfilereader;

import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;


public class worksheet {

    private int rowCounts;
    private int coloumnCounts;
    private ArrayList<String> coloumnDataTypes;
    private Cell currentCell;
    private Sheet datasheet;

    public worksheet(Sheet datasheet){
        this.rowCounts = datasheet.getPhysicalNumberOfRows();;
        this.coloumnCounts = datasheet.getRow(0).getPhysicalNumberOfCells();;
        this.coloumnDataTypes = this.getDataTypes(datasheet);
    }

    private ArrayList<String> getDataTypes(Sheet datasheet){
        ArrayList<String> result = new ArrayList<String>();
        Row row = datasheet.getRow(1);
        int count = 1;

        Iterator<Cell> cellIterator = row.iterator();

        while (cellIterator.hasNext()) {  // Loop for Column
                Cell currentCell = cellIterator.next();
            result.add(getDataType(currentCell));
            count++;
        }
        return result;
    }

    public int getColumnCountsOfWorksheet() {
        return this.coloumnCounts;
    }

    public int getRowCountsOfWorksheet() {
        return this.rowCounts;
    }


    public String getDataType(Cell currentCell) {
        if (currentCell.getCellTypeEnum() == CellType.STRING) {
            return String.class.getSimpleName();

        } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(currentCell)) {  // Checking if NUMERIC type is Date or not
                return Date.class.getSimpleName();
            } else {
                if (currentCell.getNumericCellValue() == Math.ceil(currentCell.getNumericCellValue())) { // checking if the value is float
                    return int.class.getSimpleName();
                } else {
                    return float.class.getSimpleName();
                }

            }
        } else if (currentCell.getCellTypeEnum() == CellType.FORMULA) {
            if (currentCell.getCachedFormulaResultTypeEnum() == CellType.NUMERIC) {
                if (HSSFDateUtil.isCellDateFormatted(currentCell)) {
                    return Date.class.getSimpleName();
                } else {
                    return float.class.getSimpleName();
                }
            } else if (currentCell.getCachedFormulaResultTypeEnum() == CellType.STRING) {
                return String.class.getSimpleName();
            }
        } else if (currentCell.getCellTypeEnum() == CellType.BOOLEAN) {
            return boolean.class.getSimpleName();
        }
        return "";
    }


    public int getRowCounts(){
        return rowCounts;
    }

    public int getColoumnCounts() {
        return coloumnCounts;
    }

    public ArrayList<String> getColoumnDataTypes(){
        return this.coloumnDataTypes;
    }

}

