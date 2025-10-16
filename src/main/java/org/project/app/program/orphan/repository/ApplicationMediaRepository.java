package org.project.app.program.orphan.repository;

import org.project.app.program.orphan.domain.ApplicationMedia;
import org.project.app.program.orphan.enums.DocumentType;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;
import java.util.Optional;

public interface ApplicationMediaRepository extends JpaRepository<ApplicationMedia, String> {
    Optional<ApplicationMedia> findByOrphanApplicationIdAndType(String applicationId, DocumentType type);
    List<ApplicationMedia> findAllByOrphanApplicationId(String applicationId);
    void deleteByOrphanApplicationId(String orphanApplicationId);
    void deleteByOrphanApplicationIdAndType(String applicationId, DocumentType type);
}
