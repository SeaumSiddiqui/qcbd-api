package org.project.app.program.orphan.enums;

import org.project.app.converter.BengaliEnum;

public enum MothersStatus implements BengaliEnum {
    WIDOWED("বিধবা"),
    REMARRIED("পুনর্বিবাহিত"),
    DEAD("মৃত");


    private final String bengaliValue;

    MothersStatus(String bengaliValue) {
        this.bengaliValue = bengaliValue;
    }

    //@JsonValue // Serialize using Bengali value
    @Override
    public String getBengaliValue() {
        return bengaliValue;
    }
}
