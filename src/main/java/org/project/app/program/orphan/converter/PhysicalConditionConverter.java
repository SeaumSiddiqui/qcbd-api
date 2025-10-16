package org.project.app.program.orphan.converter;

import jakarta.persistence.Converter;
import org.project.app.converter.AbstractBengaliEnumConverter;
import org.project.app.program.orphan.enums.PhysicalCondition;

@Converter(autoApply = true)
public class PhysicalConditionConverter  extends AbstractBengaliEnumConverter<PhysicalCondition> {
    public PhysicalConditionConverter() {
        super(PhysicalCondition.class);
    }
}
