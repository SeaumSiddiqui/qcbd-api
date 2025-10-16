package org.project.app.program.orphan.enums;

import org.project.app.converter.BengaliEnum;

public enum Gender implements BengaliEnum {
    MALE("ছেলে"),
    FEMALE("মেয়ে");

    private final String bengaliValue;

    Gender(String bengaliValue) {
        this.bengaliValue = bengaliValue;
    }

    //@JsonValue // Serialize using Bengali value
    @Override
    public String getBengaliValue() {
        return bengaliValue;
    }
}
