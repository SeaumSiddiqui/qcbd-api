package org.project.app.program.orphan.domain;

import com.fasterxml.jackson.annotation.JsonManagedReference;
import jakarta.persistence.*;
import lombok.*;
import org.project.app.program.Program;

import java.util.ArrayList;
import java.util.List;

@EqualsAndHashCode(callSuper = true)
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@Entity
@DiscriminatorValue("ORPHAN")
@Table(name = "orphan_applications")
public class OrphanApplication extends Program {

    @OneToOne(cascade = CascadeType.ALL, orphanRemoval = true)
    private PrimaryInformation primaryInformation;

    @OneToOne(cascade = CascadeType.ALL, orphanRemoval = true)
    private Address address;

    @OneToMany(mappedBy = "application", cascade = CascadeType.ALL, fetch = FetchType.LAZY, orphanRemoval = true)
    @JsonManagedReference("application-familyMembers")
    @Builder.Default
    private List<FamilyMember> familyMembers = new ArrayList<>();

    @OneToOne(cascade = CascadeType.ALL, orphanRemoval = true)
    private BasicInformation basicInformation;

    @OneToOne(cascade = CascadeType.ALL, orphanRemoval = true)
    @Builder.Default
    private Verification verification = new Verification();

}
