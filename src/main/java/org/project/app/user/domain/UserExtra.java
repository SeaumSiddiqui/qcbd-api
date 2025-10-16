package org.project.app.user.domain;

import com.fasterxml.jackson.annotation.JsonManagedReference;
import jakarta.persistence.*;
import lombok.*;
import org.project.app.enrollment.domain.ProgramEnrollment;
import org.project.app.program.Program;
import org.springframework.lang.NonNull;

import java.util.ArrayList;
import java.util.List;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@RequiredArgsConstructor
@Entity
@Table(name = "user_extras")
public class UserExtra {
    @Id
    @NonNull
    private String userId; // Form keycloak userId

    private String cell;
    private String address;

    @OneToMany(mappedBy = "beneficiary", cascade = CascadeType.ALL, orphanRemoval = true)
    @JsonManagedReference
    private List<Program> programs = new ArrayList<>();

    @OneToMany(mappedBy = "beneficiary", cascade = CascadeType.ALL, orphanRemoval = true)
    @JsonManagedReference
    private List<ProgramEnrollment> enrollments = new ArrayList<>();

    @Embedded
    private BeneficiaryExtra beneficiaryExtra;
}
