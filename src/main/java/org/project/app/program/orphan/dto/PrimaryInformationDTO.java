package org.project.app.program.orphan.dto;

import lombok.Data;
import org.project.app.program.orphan.enums.Gender;
import org.project.app.program.orphan.enums.MothersStatus;

import java.time.LocalDate;

@Data
public class PrimaryInformationDTO {
    // Orphan
    private String fullName;
    private LocalDate dateOfBirth;
    private String bcRegistration;
    private String placeOfBirth;
    private int age;
    private Gender gender;

    // Father
    private String fathersName;
    private LocalDate dateOfDeath;
    private String causeOfDeath;

    // Mother
    private String mothersName;
    private String mothersOccupation;
    private MothersStatus mothersStatus;

    // Family
    private String fixedAssets;
    private String annualIncome;
    private int numOfSiblings;

    // Academics
    private String academicInstitution;
    private int grade;
}
