package org.project.app.program.orphan.dto;

import lombok.Builder;
import lombok.Data;
import org.project.app.program.orphan.domain.*;
import org.project.app.program.orphan.enums.ApplicationStatus;

import java.time.LocalDateTime;
import java.util.List;

@Data
@Builder
public class OrphanApplicationResponseDTO {
    private String id;
    private ApplicationStatus status; //INCOMPLETE, COMPLETE, PENDING, REJECTED, ACCEPTED, GRANTED
    private String rejectionMessage;

    // Audit fields
    private Long version;

    private String createdBy;
    private String lastReviewedBy;

    private LocalDateTime createdAt;
    private LocalDateTime lastModifiedAt;

    // Submodules
    private PrimaryInformation primaryInformation;
    private Address address;
    private List<FamilyMember> familyMembers;
    private BasicInformation basicInformation;
    private Verification verification;
}
