package org.project.app.program.orphan.domain;

import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.hibernate.annotations.DynamicInsert;
import org.hibernate.annotations.DynamicUpdate;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
@DynamicInsert
@DynamicUpdate
@Entity
public class Verification {
    @Id
    @GeneratedValue(strategy = GenerationType.UUID)
    private String id;
    private String agentUserId;
    private String authenticatorUserId;
    private String investigatorUserId;
    private String qcSwdUserId;
}
