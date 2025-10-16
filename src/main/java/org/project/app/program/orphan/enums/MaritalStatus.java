package org.project.app.program.orphan.enums;

import org.project.app.converter.BengaliEnum;

public enum MaritalStatus implements BengaliEnum {
    MARRIED("বিবাহিত"),
    UNMARRIED("অবিবাহিত"),
    DIVORCED("তালাকপ্রাপ্ত"),
    WIDOWED("বিধবা");

    private final String bengaliValue;

    MaritalStatus(String bengaliValue) {
        this.bengaliValue = bengaliValue;
    }

    //@JsonValue // Serialize using Bengali value
    @Override
    public String getBengaliValue() {
        return bengaliValue;
    }
}
