package org.project.app.program.orphan.repository;

import org.project.app.program.orphan.domain.OrphanApplication;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.JpaSpecificationExecutor;

public interface OrphanApplicationRepository extends JpaRepository<OrphanApplication, String>, JpaSpecificationExecutor<OrphanApplication> {
    boolean existsByPrimaryInformationBcRegistration(String bcRegistration);
}
