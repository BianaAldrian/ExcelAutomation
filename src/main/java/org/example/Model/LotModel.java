package org.example.Model;

import java.util.ArrayList;

public class LotModel {
    private String LotTitle;
    private ArrayList<String> items;
    private ArrayList<String> QTY;

    public LotModel(String lotTitle, ArrayList<String> items, ArrayList<String> QTY) {
        LotTitle = lotTitle;
        this.items = items;
        this.QTY = QTY;
    }

    public String getLotTitle() {
        return LotTitle;
    }

    public ArrayList<String> getItems() {
        return items;
    }

    public ArrayList<String> getQTY() {
        return QTY;
    }
}
