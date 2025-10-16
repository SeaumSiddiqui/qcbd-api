package org.project.app.enrollment.domain;

import com.fasterxml.jackson.annotation.JsonBackReference;
import com.fasterxml.jackson.annotation.JsonIgnore;
import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.project.app.program.Program;
import org.project.app.enrollment.enums.ProgramStatus;
import org.project.app.user.domain.UserExtra;
import org.springframework.data.annotation.CreatedDate;
import org.springframework.data.annotation.LastModifiedDate;
import org.springframework.data.jpa.domain.support.AuditingEntityListener;

import java.math.BigDecimal;
import java.time.LocalDateTime;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@Entity
@EntityListeners(AuditingEntityListener.class)
@Table(name = "program_enrollments")
public class ProgramEnrollment {
    @Id
    @GeneratedValue(strategy = GenerationType.UUID)
    private String id;

    @Enumerated(EnumType.STRING)
    private ProgramStatus status; // ACTIVE, SUSPENDED, COMPLETED
    private BigDecimal monthlyAmount;
    private boolean isPaid;

    @ManyToOne(fetch = FetchType.LAZY, optional = false)
    @JoinColumn(name = "userId", referencedColumnName = "userId")
    @JsonBackReference
    @JsonIgnore
    private UserExtra beneficiary;

    @ManyToOne(fetch = FetchType.LAZY, optional = false)
    @JoinColumn(name = "program_id")
    @JsonBackReference
    @JsonIgnore
    private Program program;

    @CreatedDate
    @Column(nullable = false, updatable = false)
    private LocalDateTime enrolledAt;
    @LastModifiedDate
    @Column(insertable = false)
    private LocalDateTime lastUpdatedAt;
}