package org.project.app.program;

import com.fasterxml.jackson.annotation.JsonBackReference;
import com.fasterxml.jackson.annotation.JsonIgnore;
import jakarta.persistence.*;
import jakarta.validation.constraints.NotNull;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;
import org.project.app.enrollment.domain.ProgramEnrollment;
import org.project.app.program.orphan.enums.ApplicationStatus;
import org.project.app.user.domain.UserExtra;
import org.springframework.data.annotation.CreatedBy;
import org.springframework.data.annotation.CreatedDate;
import org.springframework.data.annotation.LastModifiedBy;
import org.springframework.data.annotation.LastModifiedDate;
import org.springframework.data.jpa.domain.support.AuditingEntityListener;

import java.time.LocalDateTime;

@EqualsAndHashCode(onlyExplicitlyIncluded = true)
@Data
@NoArgsConstructor
@AllArgsConstructor
@Entity
@EntityListeners(AuditingEntityListener.class)
@Inheritance(strategy = InheritanceType.JOINED)
@DiscriminatorColumn(name = "program_type", discriminatorType = DiscriminatorType.STRING)
@Table(name = "programs")
public abstract class Program {
    @Id
    @GeneratedValue(strategy = GenerationType.UUID)
    private String id;

    // Reference to enrollment (only when application is granted)
    @OneToOne(mappedBy = "program", cascade = CascadeType.ALL, orphanRemoval = true)
    private ProgramEnrollment enrollment;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "beneficiary_user_id", referencedColumnName = "userId")
    @JsonBackReference
    @JsonIgnore
    private UserExtra beneficiary;

    @NotNull
    @Enumerated(EnumType.STRING)
    private ApplicationStatus status; //INCOMPLETE, COMPLETE, PENDING, REJECTED, ACCEPTED, GRANTED
    private String rejectionMessage;


    // Audit fields
    @Version
    private Long version;

    @CreatedBy
    @Column(nullable = false, updatable = false)
    private String createdBy;
    @LastModifiedBy
    @Column(insertable = false)
    private String lastReviewedBy;

    @CreatedDate
    @Column(nullable = false, updatable = false)
    private LocalDateTime createdAt;
    @LastModifiedDate
    @Column(insertable = false)
    private LocalDateTime lastModifiedAt;
}
