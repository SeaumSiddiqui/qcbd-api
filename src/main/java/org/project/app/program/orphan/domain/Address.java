package org.project.app.program.orphan.domain;

import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import jakarta.validation.constraints.NotBlank;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.hibernate.annotations.DynamicInsert;
import org.hibernate.annotations.DynamicUpdate;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@DynamicInsert
@DynamicUpdate
@Entity
public class Address {
    @Id
    @GeneratedValue(strategy = GenerationType.UUID)
    private String id;
    private boolean isSameAsPermanent;

    /** permanent address fields **/
    @NotBlank
    private String permanentDistrict;
    @NotBlank
    private String permanentSubDistrict;
    private String permanentUnion;
    private String permanentVillage;
    private String permanentArea;

    /** present address fields **/
    private String presentDistrict;
    private String presentSubDistrict;
    private String presentUnion;
    private String presentVillage;
    private String presentArea;

}
