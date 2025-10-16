package org.project.app.program.orphan.converter;

import jakarta.persistence.Converter;
import org.project.app.converter.AbstractBengaliEnumConverter;
import org.project.app.program.orphan.enums.Gender;

@Converter(autoApply = true)
public class GenderConverter extends AbstractBengaliEnumConverter<Gender> {
    public GenderConverter() {
        super(Gender.class);
    }
}
