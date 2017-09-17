package com.innovative.excelfilereader;

import java.util.ArrayList;

public class RowString {

    private ArrayList<String> row;
    public RowString(){};
    public RowString(ArrayList<String> str){
        this.row = str;
    }
    public ArrayList<String> getRow(){
    return row;
    }
    public void setRow(ArrayList<String> row){
        this.row = row;

    }
}
