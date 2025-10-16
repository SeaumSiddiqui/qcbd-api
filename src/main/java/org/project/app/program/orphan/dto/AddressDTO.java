package org.project.app.program.orphan.dto;

import lombok.Data;

@Data
public class AddressDTO {
    private boolean isSameAsPresent;

    /** present address fields **/
    private String presentDistrict;
    private String presentSubDistrict;
    private String presentUnion;
    private String presentVillage;
    private String presentArea;

    /** permanent address fields **/
    private String permanentDistrict;
    private String permanentSubDistrict;
    private String permanentUnion;
    private String permanentVillage;
    private String permanentArea;
}
