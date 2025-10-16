package org.project.app.program.orphan.domain;

import jakarta.persistence.*;
import jakarta.validation.constraints.NotNull;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.hibernate.annotations.DynamicInsert;
import org.hibernate.annotations.DynamicUpdate;
import org.project.app.program.orphan.enums.HouseType;
import org.project.app.program.orphan.enums.PhysicalCondition;
import org.project.app.program.orphan.enums.ResidenceStatus;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@DynamicInsert
@DynamicUpdate
@Entity
public class BasicInformation {
    @Id
    @GeneratedValue(strategy = GenerationType.UUID)
    private String id;

    @NotNull(message = "physical condition information is required")
    private PhysicalCondition physicalCondition; //  HEALTHY, SICK, DISABLED

    private boolean hasCriticalIllness;
    private String typeOfIllness;
    private boolean isResident;

    private ResidenceStatus residenceStatus;  // OWN, RENTED, SHELTERED, HOMELESS
    private HouseType houseType; // CONCRETE_HOUSE, SEMI_CONCRETE_HOUSE, MUD_HOUSE

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
