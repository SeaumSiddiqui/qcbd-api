package org.project.app.program;

import lombok.Builder;
import lombok.Data;
import org.project.app.enrollment.dto.ProgramEnrollmentDTO;
import org.project.app.program.orphan.enums.ApplicationStatus;

@Data
@Builder
public class ProgramDTO {
    private String id;
    private ApplicationStatus status;
    private String rejectionMessage;
    private ProgramEnrollmentDTO enrollment;
}
