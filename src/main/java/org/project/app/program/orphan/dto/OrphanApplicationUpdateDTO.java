package org.project.app.program.orphan.dto;

import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.NotNull;
import lombok.Data;
import org.project.app.enrollment.dto.ProgramEnrollmentUpdateDTO;
import org.project.app.program.orphan.enums.ApplicationStatus;

import java.util.List;

@Data
public class OrphanApplicationUpdateDTO {
    @NotBlank
    private String userId;

    private ApplicationStatus status;
    private String rejectionMessage;

    private ProgramEnrollmentUpdateDTO programEnrollmentUpdateDto;

    private PrimaryInformationDTO primaryInformation;
    private AddressDTO address;
    private List<FamilyMemberDTO> familyMembers;
    private BasicInformationDTO basicInformation;
    private VerificationDTO verification;

    @NotNull
    private Long version;
}
