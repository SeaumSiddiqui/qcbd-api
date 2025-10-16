package org.project.app.program.orphan.service;

import jakarta.persistence.EntityNotFoundException;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.project.app.exception.DocumentNotFoundException;
import org.project.app.program.orphan.domain.ApplicationMedia;
import org.project.app.program.orphan.domain.OrphanApplication;
import org.project.app.program.orphan.enums.DocumentType;
import org.project.app.program.orphan.repository.ApplicationMediaRepository;
import org.project.app.program.orphan.repository.OrphanApplicationRepository;
import org.project.app.storage.FileStorageService;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import java.time.Instant;
import java.util.Map;
import java.util.stream.Collectors;

@Slf4j
@RequiredArgsConstructor
@Transactional
@Service
public class OrphanMediaService {

    private final FileStorageService fileStorageService;
    private final ApplicationMediaRepository mediaRepository;
    private final OrphanApplicationRepository orphanAppRepository;

    public void uploadDocument(String applicationId, MultipartFile file, DocumentType docType) {
        String key = generateStorageKey(applicationId, docType, file);

        // Update if same type of media already exist else create
        mediaRepository.findByOrphanApplicationIdAndType(applicationId, docType)
                .ifPresentOrElse(existing -> {
                    // Delete form s3
                    fileStorageService.deleteObject(existing.getS3Key());

                    // Update metadata in DB
                    existing.setS3Key(key);
                    existing.setUploadedAt(Instant.now());

                    log.info("Updating existing document for: Application ID = {}, type = {}", applicationId, docType);
                    mediaRepository.save(existing);
                }, () -> {
                    OrphanApplication application = orphanAppRepository.findById(applicationId)
                            .orElseThrow(() -> new EntityNotFoundException("Application not found: " + applicationId));

                    // Save metadata in DB
                    ApplicationMedia media = new ApplicationMedia();
                    media.setOrphanApplication(application);
                    media.setType(docType);
                    media.setS3Key(key);
                    media.setUploadedAt(Instant.now());

                    log.info("Uploading new document for: Application ID = {}, type = {}", applicationId, docType);
                    mediaRepository.save(media);
                });

        // Upload new document to s3
        fileStorageService.uploadFile(key, file);
    }

    public String getDocumentUrl(String applicationId, DocumentType docType) {
        ApplicationMedia media = mediaRepository
                .findByOrphanApplicationIdAndType(applicationId, docType)
                .orElseThrow(() -> new DocumentNotFoundException("Document not found for type " + docType));

        return fileStorageService.generatePreSignedUrl(media.getS3Key(), true);
    }

    public Map<DocumentType, String> getAllDocumentUrls(String applicationId) {
        return mediaRepository.findAllByOrphanApplicationId(applicationId).stream()
                .collect(Collectors.toMap(
                        ApplicationMedia::getType,
                        media -> fileStorageService.generatePreSignedUrl(media.getS3Key(), true)
                ));
    }

    private String generateStorageKey(String applicationId, DocumentType docType, MultipartFile file) {
        String extension = fileStorageService.getFileExtension(file);
        return String.format("applications/OrphanApplication/%s/%s-%d.%s",
                applicationId,
                docType.name().toLowerCase(),
                System.currentTimeMillis(),
                extension
        );
    }

    public void deleteAllMediaForApplication(String applicationId) {
        String prefix = getApplicationFolder(applicationId);

        // Delete from S3
        fileStorageService.deleteObjectsWithPrefix(prefix);

        // Delete from DB
        mediaRepository.deleteByOrphanApplicationId(applicationId);
    }

    private String getApplicationFolder(String applicationId) {
        return "applications/OrphanApplication/" + applicationId;
    }

}
