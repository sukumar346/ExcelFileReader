package com.innovative.excelfilereader;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.text.SimpleDateFormat;


public class ExcelReader {


    public ExcelInfo readExcel(String filePath, int sheetIndex) throws
            FileNotFoundException,
            IllegalArgumentException,
            IndexOutOfBoundsException,
            IOException,
            Exception{
        ArrayList<RowString> rs = new ArrayList<RowString>();
        ExcelInfo excelInfo = new ExcelInfo();
        try {
            sheetIndex = sheetIndex - 1;

            // Getting Workbook objects and Row Iterator
            FileInputStream excelFile = new FileInputStream(new File(filePath));
            Workbook workbook = new HSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(sheetIndex);
            long rowCounts = datatypeSheet.getPhysicalNumberOfRows();
            excelInfo.setRowCounts(rowCounts);
            int noofWorksheets = workbook.getNumberOfSheets();
            excelInfo.setNoofWorksheets(noofWorksheets);
            int coloumnCounts = datatypeSheet.getRow(0).getPhysicalNumberOfCells();
            excelInfo.setColoumnCounts(coloumnCounts);
            ArrayList<String> coloumnDataTypes = this.getDataTypes(datatypeSheet);
            excelInfo.setColoumnDataTypes(coloumnDataTypes);

            Iterator<Row> iterator = datatypeSheet.iterator();
            rs = getRows(iterator);
            excelInfo.setResultSet(rs);


        } catch (FileNotFoundException fe) {
            throw fe;
        } catch (IllegalArgumentException iae) {
            throw iae;
        } catch (IndexOutOfBoundsException ioe){
            throw ioe;
        }catch (IOException io) {
            throw io;
        }catch (Exception e) {
            throw e;
        }
        return excelInfo;
    }

    public ArrayList<RowString> getRows(Iterator<Row> iterator){
        ArrayList<RowString> result = new ArrayList<RowString>();
        int count = 0; // for Row counts

        while (iterator.hasNext()) {   // Loop for Row
            Row currentRow = iterator.next();
            Iterator<Cell> cellIterator = currentRow.iterator();
            result.add(loopForColumn(cellIterator));

        }
        return result;
    }
    public void getDashLine(int columnCounts){
        for(int i=0; i<(columnCounts*40)-4; i++){
            System.out.print("-");

        }
        System.out.println();

    }


    //Method to loop over the columns in worksheet
    public RowString loopForColumn(Iterator<Cell> cellIterator) {
        ArrayList<String> result = new ArrayList<String>();
        while (cellIterator.hasNext()) {  // Loop for Column
            Cell currentCell = cellIterator.next();
            result.add(getCell(currentCell));

        }
        RowString rs = new RowString();
        rs.setRow(result);
        return rs;
    }

    //Method for Getting Cells according to their type
    public String getCell(Cell currentCell) {
        if (currentCell.getCellTypeEnum() == CellType.STRING) {
            return wraptext(currentCell.getStringCellValue());

        } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(currentCell)) {  // Checking if NUMERIC type is Date or not
               return new SimpleDateFormat("MM/dd/yyyy").format(currentCell.getDateCellValue());
                
            } else {
                if (currentCell.getNumericCellValue() == Math.ceil(currentCell.getNumericCellValue())) { // checking if the value is float
                    return ((int) currentCell.getNumericCellValue()) + "";
                } else {
                    return currentCell.getNumericCellValue() + "";
                }

            }
        } else if (currentCell.getCellTypeEnum() == CellType.FORMULA) {
            if (currentCell.getCachedFormulaResultTypeEnum() == CellType.NUMERIC) {
                if (HSSFDateUtil.isCellDateFormatted(currentCell)) {
                    return new SimpleDateFormat("MM/dd/yyyy").format(currentCell.getDateCellValue());
                } else {
                    double db = currentCell.getNumericCellValue();
                    String str = new Double(db).toString();
                    return str;
                }
            } else if (currentCell.getCachedFormulaResultTypeEnum() == CellType.STRING) {
                return currentCell.getStringCellValue();
            }
            return currentCell.getCellFormula();
        } else if (currentCell.getCellTypeEnum() == CellType.BOOLEAN) {
            Boolean b = currentCell.getBooleanCellValue();
            String str2 = String.valueOf(b);
            return str2;
        } else if (currentCell.getCellTypeEnum() == CellType.BLANK) {
            return "NULL";
        }
        return "|";

    }

   public static String wraptext(String str){
        final int FIXED_WIDTH = 30;
        String temp = "";
        if(str !=null && str.length() > FIXED_WIDTH) {
            temp = str.substring(0, FIXED_WIDTH) + "...";
        } else {
            temp = str;
        }
        return temp;
    }



    public ArrayList<String> getDataTypes(Sheet datasheet){
        ArrayList<String> result = new ArrayList<String>();
        Row row = datasheet.getRow(1);

        Iterator<Cell> cellIterator = row.iterator();

        while (cellIterator.hasNext()) {  // Loop for Column
            Cell currentCell = cellIterator.next();
            result.add(getDataType(currentCell));
        }
        return result;
    }

    public String getDataType(Cell currentCell) {
        if (currentCell.getCellTypeEnum() == CellType.STRING) {
            return String.class.getSimpleName();

        } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(currentCell)) {  // Checking if NUMERIC type is Date or not
                return Date.class.getSimpleName().toString();
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



}



