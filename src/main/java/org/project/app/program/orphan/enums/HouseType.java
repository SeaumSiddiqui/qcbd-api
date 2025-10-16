package org.project.app.program.orphan.enums;

import org.project.app.converter.BengaliEnum;

public enum HouseType implements BengaliEnum {
    CONCRETE_HOUSE("পাকা"),
    SEMI_CONCRETE_HOUSE("আধপাকা"),
    MUD_HOUSE("কাঁচা");

    private final String bengaliValue;

    HouseType(String bengaliValue) {
        this.bengaliValue = bengaliValue;
    }

    // @JsonValue // Serialize using Bengali value
    @Override
    public String getBengaliValue() {
        return bengaliValue;
    }
}
