package org.project.app.converter;

import org.project.app.program.orphan.enums.*;

public class BengaliEnumHelper {
    static final Class<? extends BengaliEnum>[] SUPPORTED_ENUMS =
            new Class[] {
                    Gender.class,
                    HouseType.class,
                    MaritalStatus.class,
                    MothersStatus.class,
                    PhysicalCondition.class,
                    ResidenceStatus.class,
            };
}
