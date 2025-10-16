package org.project.app.program.orphan.domain;

import com.fasterxml.jackson.annotation.JsonBackReference;
import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.project.app.program.orphan.enums.Gender;
import org.project.app.program.orphan.enums.MaritalStatus;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@Entity
public class FamilyMember {
    @Id
    @GeneratedValue(strategy = GenerationType.UUID)
    private String id;

    private String name;
    private int age;
    private int siblingsGrade;
    private String occupation;

    private Gender siblingsGender;
    private MaritalStatus maritalStatus;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "application_id")
    @JsonBackReference("application-familyMembers")
    private OrphanApplication application;

}