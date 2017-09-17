package com.innovative.excelfilereader;

import java.util.ArrayList;

public class ExcelInfo {


    long rowCounts = 0;
    int noofWorksheets = 0;
    int coloumnCounts = 0;
    ArrayList<String> coloumnDataTypes = new ArrayList<String>();
    ArrayList<RowString> resultSet = new ArrayList<RowString>();


    public int getNoofWorksheets(){
        return noofWorksheets;
    }
    public void setNoofWorksheets(int noofWorksheets) {
        this.noofWorksheets = noofWorksheets;
    }

    public long getRowCounts() {
        return rowCounts;
    }

    public void setRowCounts(long rowCounts) {
        this.rowCounts = rowCounts;
    }

    public int getColoumnCounts() {
        return coloumnCounts;
    }

    public void setColoumnCounts(int coloumnCounts) {
        this.coloumnCounts = coloumnCounts;
    }

    public ArrayList<String> getColoumnDataTypes() {
        return coloumnDataTypes;
    }

    public void setColoumnDataTypes(ArrayList<String> coloumnDataTypes) {
        this.coloumnDataTypes = coloumnDataTypes;
    }

    public ArrayList<RowString> getResultSet() {
        return resultSet;
    }

    public void setResultSet(ArrayList<RowString> resultSet) {
        this.resultSet = resultSet;
    }
}

