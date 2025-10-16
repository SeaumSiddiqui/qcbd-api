package org.project.app.converter;

import jakarta.persistence.AttributeConverter;

public abstract class AbstractBengaliEnumConverter<E extends Enum<E> & BengaliEnum> implements AttributeConverter<E, String>{

    private final Class<E> enumType;

    protected AbstractBengaliEnumConverter(Class<E> enumType) {
        this.enumType = enumType;
    }

    @Override
    public String convertToDatabaseColumn(E attribute) {
        return attribute != null ? attribute.getBengaliValue() : null;
    }

    @Override
    public E convertToEntityAttribute(String dbData) {
        if (dbData == null) return null;

        for (E constant : enumType.getEnumConstants()) {
            if (constant.getBengaliValue().equals(dbData)) {
                return constant;
            }
        }
        throw new IllegalArgumentException("Invalid Bengali value for " + enumType.getSimpleName() + ": " + dbData);
    }

}
