package org.project.app.program.orphan.converter;

import jakarta.persistence.Converter;
import org.project.app.converter.AbstractBengaliEnumConverter;
import org.project.app.program.orphan.enums.HouseType;

@Converter(autoApply = true)
public class HouseTypeConverter extends AbstractBengaliEnumConverter<HouseType> {
    public HouseTypeConverter() {
        super(HouseType.class);
    }
}
