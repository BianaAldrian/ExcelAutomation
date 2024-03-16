package org.example.Model;

import java.util.ArrayList;

public class ItemsQtyModel {

    private ArrayList<String> items;
    private ArrayList<String> QTY;

    public ItemsQtyModel(ArrayList<String> items, ArrayList<String> QTY) {
        this.items = items;
        this.QTY = QTY;
    }

    public ArrayList<String> getItems() {
        return items;
    }

    public ArrayList<String> getQTY() {
        return QTY;
    }
}
