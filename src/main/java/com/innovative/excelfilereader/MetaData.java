package com.innovative.excelfilereader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;

public class MetaData {

    private ArrayList<worksheet> worksheets;
    private Workbook workbook;
    public int noOfWorksheets;

    public MetaData(String filePath) throws
            FileNotFoundException,
            IllegalArgumentException,
            IndexOutOfBoundsException,
            IOException,
            Exception {
            try {
                FileInputStream excelFile = new FileInputStream(new File(filePath));
                this.workbook = new HSSFWorkbook(excelFile);
                this.noOfWorksheets = workbook.getNumberOfSheets();
            } catch(IllegalArgumentException iae){
            throw iae;
            }catch(FileNotFoundException fe) {
                throw fe;
            }catch(IndexOutOfBoundsException ioe){
            throw ioe;
            }catch(IOException io){
            throw io;
            } catch(Exception e){
            throw e;
        }
    }

    public int getNoOfWorksheets() {
        return noOfWorksheets;
    }

    public ArrayList<String> getSheetNames(){
        ArrayList<String> result = new ArrayList<String>();
        for (int i=0; i<this.workbook.getNumberOfSheets(); i++) {
            result.add(i+1+". "+this.workbook.getSheetName(i));
        }
        return result;
    }

    public void getDashLine(int columnCounts){
        for(int i=0; i<(columnCounts*40)-4; i++){
            System.out.print("-");

        }
        System.out.println();

    }
    public static void printUsage(){
        System.out.println("give valid Worksheet Number. \n");

    }


    public ArrayList<worksheet> getSheets(){
        ArrayList<worksheet> sheets = new ArrayList<worksheet>();
        for (int i=0; i<this.workbook.getNumberOfSheets(); i++) {
            worksheet w = new worksheet(this.workbook.getSheetAt(i));
            sheets.add(w);
        }
        return sheets;
    }
}
