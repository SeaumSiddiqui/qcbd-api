package org.project.app.user.repository;

import org.project.app.user.domain.UserMedia;
import org.project.app.user.enums.UserMediaType;
import org.springframework.data.jpa.repository.JpaRepository;

import java.util.List;
import java.util.Optional;

public interface UserMediaRepository extends JpaRepository<UserMedia, String> {
    Optional<UserMedia> findByUser_UserIdAndType(String userId, UserMediaType type);
    List<UserMedia> findAllByUser_UserId(String userId);
    void deleteByUser_UserId(String userId);
}
