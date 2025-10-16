package org.project.app.program.orphan.domain;

import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.project.app.program.orphan.enums.DocumentType;

import java.time.Instant;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@Entity
@Table(name = "application_media", uniqueConstraints = {
        @UniqueConstraint(columnNames = {"orphan_application_id", "type"})
})
public class ApplicationMedia {

    @Id
    @GeneratedValue(strategy = GenerationType.UUID)
    private String id;

    @ManyToOne(optional = false)
    @JoinColumn(name = "orphan_application_id")
    private OrphanApplication orphanApplication;

    @Enumerated(EnumType.STRING)
    @Column(nullable = false)
    private DocumentType type;

    @Column(name = "s3_key", nullable = false)
    private String s3Key;

    @Column(name = "uploaded_at", nullable = false)
    private Instant uploadedAt;
}

