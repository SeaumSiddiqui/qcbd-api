package org.project.app.program.orphan.converter;

import jakarta.persistence.Converter;
import org.project.app.converter.AbstractBengaliEnumConverter;
import org.project.app.program.orphan.enums.ResidenceStatus;

@Converter(autoApply = true)
public class ResidenceStatusConverter extends AbstractBengaliEnumConverter<ResidenceStatus> {
    public ResidenceStatusConverter() {
        super(ResidenceStatus.class);
    }
}
