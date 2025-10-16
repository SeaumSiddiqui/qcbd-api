package org.project.app.enrollment.dto;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Data;
import org.project.app.enrollment.enums.ProgramStatus;

import java.math.BigDecimal;

@Data
public class ProgramEnrollmentUpdateDTO {
    private ProgramStatus status;
    private BigDecimal monthlyAmount;
    @JsonProperty("isPaid")
    private boolean isPaid;
}
