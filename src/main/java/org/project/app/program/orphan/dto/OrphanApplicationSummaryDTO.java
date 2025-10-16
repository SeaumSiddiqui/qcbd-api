package org.project.app.program.orphan.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Builder;
import lombok.Data;
import org.project.app.program.orphan.enums.ApplicationStatus;
import org.project.app.program.orphan.enums.Gender;

import java.time.LocalDate;

@Builder
@Data
@JsonInclude(JsonInclude.Include.NON_NULL)
@JsonIgnoreProperties(ignoreUnknown = true)
public class OrphanApplicationSummaryDTO {
    private String id;
    private String fullName;
    private String fathersName;
    private LocalDate dateOfBirth;
    private Gender gender;
    private String permanentDistrict;
    private ApplicationStatus status;
}
