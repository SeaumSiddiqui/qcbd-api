package org.project.app.program.orphan.enums;

import org.project.app.converter.BengaliEnum;

public enum ResidenceStatus implements BengaliEnum {
    OWN("নিজ"),
    RENTED("ভাড়া বাসা"),
    SHELTERED("আত্নীয়ের বাসা"),
    HOMELESS("গৃহহীন");

    private final String bengaliValue;

    ResidenceStatus(String bengaliValue) {
        this.bengaliValue = bengaliValue;
    }

    //@JsonValue // Serialize using Bengali value
    @Override
    public String getBengaliValue() {
        return bengaliValue;
    }
}
