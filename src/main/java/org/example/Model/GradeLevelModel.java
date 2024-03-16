package org.example.Model;

import java.util.ArrayList;

public class GradeLevelModel {
    private String gradeLevel;
    private ArrayList<LotModel> lotHolder;

    public GradeLevelModel(String gradeLevel, ArrayList<LotModel> lotHolder) {
        this.gradeLevel = gradeLevel;
        this.lotHolder = lotHolder;
    }

    public String getGradeLevel() {
        return gradeLevel;
    }

    public ArrayList<LotModel> getLotHolder() {
        return lotHolder;
    }
}
