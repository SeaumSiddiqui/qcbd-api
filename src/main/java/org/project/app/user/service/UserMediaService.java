package org.project.app.user.service;

import jakarta.persistence.EntityNotFoundException;
import jakarta.transaction.Transactional;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.project.app.exception.MediaNotFoundException;
import org.project.app.storage.FileStorageService;
import org.project.app.user.domain.UserExtra;
import org.project.app.user.domain.UserMedia;
import org.project.app.user.enums.UserMediaType;
import org.project.app.user.repository.UserExtraRepository;
import org.project.app.user.repository.UserMediaRepository;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.time.Instant;
import java.util.Map;
import java.util.stream.Collectors;

@Slf4j
@RequiredArgsConstructor
@Transactional
@Service
public class UserMediaService {

    private final FileStorageService fileStorageService;
    private final UserExtraRepository userRepository;
    private final UserMediaRepository userMediaRepository;

    public void uploadUserMedia(String userId, MultipartFile file, UserMediaType type) {
        String key = generateStorageKey(userId, type, file);

        // Update if same type of media already exist else create
        userMediaRepository.findByUser_UserIdAndType(userId, type)
                .ifPresentOrElse(existing -> {
                    // Delete from s3
                    fileStorageService.deleteObject(existing.getS3Key());

                    // Update metadata in DB
                    existing.setS3Key(key);
                    existing.setUploadedAt(Instant.now());

                    log.info("Updating existing user media: user={}, type={}", userId, type);
                    userMediaRepository.save(existing);
                }, () -> {
                    UserExtra user = userRepository.findById(userId)
                            .orElseThrow(() -> new EntityNotFoundException("User not found: " + userId));

                    // Save metadata in DB
                    UserMedia media = new UserMedia();
                    media.setUser(user);
                    media.setType(type);
                    media.setS3Key(key);
                    media.setUploadedAt(Instant.now());

                    log.info("Uploading new user media: user={}, type={}", userId, type);
                    userMediaRepository.save(media);
                });

        // Upload new media to S3
        fileStorageService.uploadFile(key, file);
    }

    public String getMediaUrl(String userId, UserMediaType type) {
        UserMedia media = userMediaRepository.findByUser_UserIdAndType(userId, type)
                .orElseThrow(() -> new MediaNotFoundException("Media not found for type: " + type));

        return fileStorageService.generatePreSignedUrl(media.getS3Key(), false);
    }

    public Map<UserMediaType, String> getAllMediaUrls(String userId) {
        return userMediaRepository.findAllByUser_UserId(userId).stream()
                .collect(Collectors.toMap(
                        UserMedia::getType,
                        media -> fileStorageService.generatePreSignedUrl(media.getS3Key(), false)
                ));
    }

    public void deleteAllMediaForUser(String userId) {
        String prefix = getUserMediaFolder(userId);

        // Delete from S3
        fileStorageService.deleteObjectsWithPrefix(prefix);

        // Delete from DB
        userMediaRepository.deleteByUser_UserId(userId);
    }

    private String generateStorageKey(String userId, UserMediaType type, MultipartFile file) {
        String extension = fileStorageService.getFileExtension(file);
        return String.format("user-media/%s/%s-%d.%s",
                userId,
                type.name().toLowerCase(),
                System.currentTimeMillis(),
                extension
        );
    }

    private String getUserMediaFolder(String userId) {
        return "user-media/" + userId;
    }
}
