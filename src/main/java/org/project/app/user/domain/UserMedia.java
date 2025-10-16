package org.project.app.user.domain;

import jakarta.persistence.*;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.project.app.user.enums.UserMediaType;

import java.time.Instant;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@Entity
@Table(name = "user_media", uniqueConstraints = {
        @UniqueConstraint(columnNames = {"user_id", "type"})
})
public class UserMedia {

    @Id
    @GeneratedValue(strategy = GenerationType.UUID)
    private String id;

    @ManyToOne(optional = false)
    @JoinColumn(name = "user_id")
    private UserExtra user;

    @Enumerated(EnumType.STRING)
    @Column(nullable = false)
    private UserMediaType type; // AVATAR, SIGNATURE

    @Column(name = "s3_key", nullable = false)
    private String s3Key;

    @Column(name = "uploaded_at", nullable = false)
    private Instant uploadedAt;
}

