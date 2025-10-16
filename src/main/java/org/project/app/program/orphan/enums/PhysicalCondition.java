package org.project.app.program.orphan.enums;

import org.project.app.converter.BengaliEnum;

public enum PhysicalCondition implements BengaliEnum {
    HEALTHY("সুস্থ"),
    SICK("অসুস্থ"),
    DISABLED("প্রতিবন্ধী");

    private final String bengaliValue;

    PhysicalCondition(String bengaliValue) {
        this.bengaliValue = bengaliValue;
    }

    //@JsonValue // Serialize using Bengali value
    @Override
    public String getBengaliValue() {
        return bengaliValue;
    }
}
