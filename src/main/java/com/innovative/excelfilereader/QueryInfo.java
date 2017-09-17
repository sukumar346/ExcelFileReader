package com.innovative.excelfilereader;


import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;

public class QueryInfo {

    public ExcelInfo queryInfo(String filePath, int sheetIndex,int columnNum, char operator, String Operand) throws
                                                        FileNotFoundException,
                                                        IllegalArgumentException,
                                                        IndexOutOfBoundsException,
                                                        IOException,
                                                        Exception {
        ArrayList<RowString> result = new ArrayList<RowString>();
        ExcelReader rs = new ExcelReader();
        ExcelInfo excelInfo = new ExcelInfo();


        try {
            sheetIndex = sheetIndex - 1;
            columnNum = columnNum - 1;
            MetaData md = new MetaData(filePath);
            ArrayList<worksheet> wsList = md.getSheets();
            // Getting Workbook objects and Row Iterator
            FileInputStream excelFile = new FileInputStream(new File(filePath));
            Workbook workbook = new HSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(sheetIndex);
            Iterator<Row> iterator = datatypeSheet.iterator();

            // Getting Worksheet information from Worksheet's object
            int noOfCol = datatypeSheet.getRow(0).getPhysicalNumberOfCells();

            excelInfo.setColoumnCounts(noOfCol);
            int count = 0; // for Row counts
            int resultCount = 0;
            RowString rowString = new RowString();

            while (iterator.hasNext()) {   // Loop for Row

                Row currentRow = iterator.next();
                Cell currentCell = CellUtil.getCell(currentRow, columnNum);
                Iterator<Cell> cellIterator = currentRow.iterator();
                if (count < 1) {

                    result.add(rs.loopForColumn(cellIterator));

                } else {
                    try {
                        double num = Double.parseDouble(Operand);
                        // is an integer!
                        switch (operator) {
                            case '=':
                                if (currentCell.getNumericCellValue() == num) {
                                    result.add(rs.loopForColumn(cellIterator));
                                    resultCount++;
                                }
                                break;
                            case '<':
                                if (currentCell.getNumericCellValue() < num) {
                                    result.add(rs.loopForColumn(cellIterator));
                                    resultCount++;
                                }
                                break;
                            case '>':
                                if (currentCell.getNumericCellValue() > num) {
                                    result.add(rs.loopForColumn(cellIterator));
                                    resultCount++;
                                }
                                break;
                            default:
                                throw new IllegalArgumentException();
                        }

                    } catch (NumberFormatException e) {
                        // not an integer!
                        switch (operator) {
                            case '=':
                                if (currentCell.getStringCellValue().equals(Operand)) {
                                    result.add(rs.loopForColumn(cellIterator));
                                    resultCount++;
                                }
                                break;
                            default:
                                throw new IllegalArgumentException();
                        }
                    }

                }
                count++;
            }
            excelInfo.setRowCounts(resultCount);
            excelInfo.setResultSet(result);
            ArrayList<String> coloumnDataTypes = rs.getDataTypes(datatypeSheet);
            excelInfo.setColoumnDataTypes(coloumnDataTypes);
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

}


