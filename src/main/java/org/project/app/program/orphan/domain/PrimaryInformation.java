package org.project.app.program.orphan.domain;

import jakarta.persistence.*;
import jakarta.validation.constraints.Max;
import jakarta.validation.constraints.Min;
import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.NotNull;
import lombok.*;
import org.project.app.program.orphan.enums.Gender;
import org.project.app.program.orphan.enums.MothersStatus;

import java.time.LocalDate;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@Entity
public class PrimaryInformation {
    @Id
    @GeneratedValue(strategy = GenerationType.UUID)
    private String id;

    @NotBlank(message = "name is required")
    private String fullName;
    @NotNull(message = "date of birth is required")
    private LocalDate dateOfBirth;
    @NotBlank(message = "registration number is required")
    private String bcRegistration;
    private String nationality;
    private String placeOfBirth;
    @Max(11)
    private int age;
    @NotNull(message = "gender is required")
    private Gender gender; // MALE, FEMALE -No BS
    private String religion;

    @NotBlank(message = "fathers name is required")
    private String fathersName;
    private LocalDate dateOfDeath;
    private String causeOfDeath;

    private String mothersName;
    private String mothersOccupation;
    private MothersStatus mothersStatus; // WIDOW, REMARRIED, DEAD

    private String fixedAssets;
    private String annualIncome;
    @Min(0)
    private int numOfSiblings; // Dynamic row for family member table based on the value in front end

    private String academicInstitution;
    @Min(0)
    private int grade;
}
