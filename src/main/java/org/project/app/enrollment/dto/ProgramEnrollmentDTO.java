package org.project.app.enrollment.dto;

import lombok.Builder;
import lombok.Data;
import org.project.app.enrollment.enums.ProgramStatus;

import java.math.BigDecimal;
import java.time.LocalDateTime;

@Data
@Builder
public class ProgramEnrollmentDTO {
    private String id;
    private ProgramStatus status;
    private BigDecimal monthlyAmount;
    private boolean isPaid;
    private LocalDateTime enrolledAt;
    private LocalDateTime lastUpdatedAt;
}
