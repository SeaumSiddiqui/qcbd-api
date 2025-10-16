package org.project.app.program.orphan.dto;

import lombok.Data;
import org.project.app.program.orphan.enums.Gender;
import org.project.app.program.orphan.enums.MaritalStatus;

@Data
public class FamilyMemberDTO {
    private String id;

    private String name;
    private int age;
    private int siblingsGrade;
    private String occupation;
    private Gender siblingsGender;
    private MaritalStatus maritalStatus;
}
