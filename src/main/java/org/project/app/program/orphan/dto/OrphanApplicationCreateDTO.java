package org.project.app.program.orphan.dto;

import jakarta.validation.constraints.NotBlank;
import lombok.Data;
import org.project.app.program.orphan.enums.ApplicationStatus;

import java.util.List;

@Data
public class OrphanApplicationCreateDTO {
    @NotBlank
    private String beneficiaryUserId;

    private ApplicationStatus status;

    private PrimaryInformationDTO primaryInformation;
    private AddressDTO address;
    private List<FamilyMemberDTO> familyMembers;
    private BasicInformationDTO basicInformation;

}
