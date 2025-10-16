package org.project.app.program.orphan.converter;

import jakarta.persistence.Converter;
import org.project.app.converter.AbstractBengaliEnumConverter;
import org.project.app.program.orphan.enums.MothersStatus;

@Converter(autoApply = true)
public class MotherStatusConverter  extends AbstractBengaliEnumConverter<MothersStatus> {
    public MotherStatusConverter() {
        super(MothersStatus.class);
    }
}
