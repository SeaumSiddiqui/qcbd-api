package org.project.app.program.orphan.dto;

import lombok.Data;
import org.project.app.program.orphan.enums.HouseType;
import org.project.app.program.orphan.enums.PhysicalCondition;
import org.project.app.program.orphan.enums.ResidenceStatus;

@Data
public class BasicInformationDTO {
    private PhysicalCondition physicalCondition;
    private boolean hasCriticalIllness;
    private String typeOfIllness;
    private boolean isResident;

    private ResidenceStatus residenceStatus;
    private HouseType houseType;

    // household details
    private int bedroom;
    private boolean balcony;
    private boolean kitchen;
    private boolean store;
    private boolean hasTubeWell;
    private boolean toilet;

    // guardian
    private String guardiansName;
    private String guardiansRelation;
    private String NID;
    private String cell1;
    private String cell2;
}
