package org.project.app.user.repository;

import org.project.app.user.domain.UserExtra;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.UUID;

public interface UserExtraRepository extends JpaRepository<UserExtra, String> {
    boolean existsByBeneficiaryExtraBcRegistration(String bcRegistration);
}
