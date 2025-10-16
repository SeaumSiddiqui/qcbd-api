package org.project.app.program.orphan.converter;

import jakarta.persistence.Converter;
import org.project.app.converter.AbstractBengaliEnumConverter;
import org.project.app.program.orphan.enums.MaritalStatus;

@Converter(autoApply = true)
public class MaritalStatusConverter extends AbstractBengaliEnumConverter<MaritalStatus> {
    public MaritalStatusConverter() {
        super(MaritalStatus.class);
    }
}
