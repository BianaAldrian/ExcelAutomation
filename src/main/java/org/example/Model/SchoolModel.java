package org.example.Model;

import java.util.ArrayList;

public class SchoolModel {
    private String region, division, schoolID, schoolName;

    private ArrayList<GradeLevelModel> gradeLevelHolder;

    public SchoolModel(String region, String division, String schoolID, String schoolName, ArrayList<GradeLevelModel> gradeLevelHolder) {
        this.region = region;
        this.division = division;
        this.schoolID = schoolID;
        this.schoolName = schoolName;
        this.gradeLevelHolder = gradeLevelHolder;
    }

    public String getRegion() {
        return region;
    }

    public String getDivision() {
        return division;
    }

    public String getSchoolID() {
        return schoolID;
    }

    public String getSchoolName() {
        return schoolName;
    }

    public ArrayList<GradeLevelModel> getGradeLevelHolder() {
        return gradeLevelHolder;
    }
}
