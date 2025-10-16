package org.project.app.storage;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.project.app.exception.FileSizeExceededException;
import org.project.app.exception.FileUploadException;
import org.project.app.exception.InvalidFileTypeException;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import software.amazon.awssdk.core.sync.RequestBody;
import software.amazon.awssdk.services.s3.S3Client;
import software.amazon.awssdk.services.s3.model.*;
import software.amazon.awssdk.services.s3.paginators.ListObjectsV2Iterable;
import software.amazon.awssdk.services.s3.presigner.S3Presigner;
import software.amazon.awssdk.services.s3.presigner.model.GetObjectPresignRequest;
import java.io.IOException;
import java.net.URI;
import java.time.Duration;
import java.time.Instant;
import java.util.*;
import java.util.stream.Collectors;

@Slf4j
@RequiredArgsConstructor
@Service
public class FileStorageService {
    private static final Duration DEFAULT_EXPIRATION = Duration.ofMinutes(30);
    private static final Duration SENSITIVE_EXPIRATION = Duration.ofMinutes(5);
    private static final Set<String> ALLOWED_MIME_TYPES = Set.of("image/jpeg", "image/png", "application/pdf");
    private static final long MAX_FILE_SIZE = 10 * 1024 * 1024; // 10MB

    private final S3Client s3Client;
    private final S3Presigner s3Presigner;

    @Value("${aws.s3.bucket-name}")
    private String bucketName;

    // Uploads a file after validating and sanitizing
    public void uploadFile(String key, MultipartFile file) {
        try {
            validateFile(file);
            String sanitizedKey = sanitizeKey(key);

            s3Client.putObject(
                    PutObjectRequest.builder()
                            .bucket(bucketName)
                            .key(sanitizedKey)
                            .contentType(file.getContentType())
                            .serverSideEncryption(ServerSideEncryption.AES256)
                            .build(),
                    RequestBody.fromInputStream(file.getInputStream(), file.getSize())
            );

            log.debug("Uploaded file to S3 at key: {}", sanitizedKey);
        } catch (IOException e) {
            throw new FileUploadException("Failed to upload file", e);
        }
    }

    // Generates a pre-signed URL to access a file (with sensitivity awareness)
    public String generatePreSignedUrl(String key, boolean isSensitive) {
        String sanitizedKey = sanitizeKey(key);
        Duration expiration = isSensitive ? SENSITIVE_EXPIRATION : DEFAULT_EXPIRATION;

        GetObjectRequest getObjectRequest = GetObjectRequest.builder()
                .bucket(bucketName)
                .key(sanitizedKey)
                .responseCacheControl("max-age=86400, private")
                .build();

        return s3Presigner.presignGetObject(
                GetObjectPresignRequest.builder()
                        .signatureDuration(expiration)
                        .getObjectRequest(getObjectRequest)
                        .build()
        ).url().toString();
    }

    // Deletes a single object
    public void deleteObject(String key) {
        try {
            String sanitizedKey = sanitizeKey(key);
            s3Client.deleteObject(DeleteObjectRequest.builder()
                    .bucket(bucketName)
                    .key(sanitizedKey)
                    .build()
            );
            log.debug("Deleted object from S3: {}", sanitizedKey);
        } catch (Exception e) {
            log.warn("Failed to delete S3 object: {}: {}", key, e.getMessage());
        }
    }

    // Lists all objects with a given prefix (internally used only)
    private List<S3Object> listObjectsWithPrefix(String prefix) {
        List<S3Object> allObjects = new ArrayList<>();
        ListObjectsV2Request request = ListObjectsV2Request.builder()
                .bucket(bucketName)
                .prefix(sanitizeKey(prefix))
                .maxKeys(500)
                .build();

        for (ListObjectsV2Response response : s3Client.listObjectsV2Paginator(request)) {
            allObjects.addAll(response.contents());
        }
        return allObjects;
    }

    // Returns the latest (most recently modified) object's key
    public Optional<String> findLatestObject(String prefix) {
        return listObjectsWithPrefix(prefix).stream()
                .max(Comparator.comparing(S3Object::lastModified))
                .map(S3Object::key);
    }

    // Deletes all objects with a given prefix
    public void deleteObjectsWithPrefix(String prefix) {
        List<String> keysToDelete = listObjectsWithPrefix(prefix).stream()
                .map(S3Object::key)
                .collect(Collectors.toList());

        if (!keysToDelete.isEmpty()) {
            deleteBatch(keysToDelete);
        }
    }

    // Bulk delete helper
    private void deleteBatch(List<String> keys) {
        List<ObjectIdentifier> identifiers = keys.stream()
                .map(key -> ObjectIdentifier.builder().key(key).build())
                .collect(Collectors.toList());

        s3Client.deleteObjects(DeleteObjectsRequest.builder()
                .bucket(bucketName)
                .delete(Delete.builder().objects(identifiers).build())
                .build()
        );

        log.debug("Deleted {} objects in batch", keys.size());
    }

    // Validates file content, size, and name
    private void validateFile(MultipartFile file) {
        String contentType = file.getContentType();
        if (!ALLOWED_MIME_TYPES.contains(contentType)) {
            throw new InvalidFileTypeException("Invalid file type: " + contentType);
        }

        if (file.getSize() > MAX_FILE_SIZE) {
            throw new FileSizeExceededException("File size exceeds 10MB");
        }

        String filename = file.getOriginalFilename();
        if (filename != null && (filename.contains("..") || filename.contains("/") || filename.contains("\\"))) {
            throw new SecurityException("Invalid file name");
        }
    }

    // Extracts extension with fallback
    public String getFileExtension(MultipartFile file) {
        String originalFilename = file.getOriginalFilename();
        if (originalFilename != null && originalFilename.contains(".")) {
            return originalFilename.substring(originalFilename.lastIndexOf('.') + 1).toLowerCase();
        }

        String contentType = file.getContentType();
        if (contentType != null && contentType.contains("/")) {
            return contentType.split("/")[1].toLowerCase();
        }

        return "bin";
    }

    // Cleans up any unsafe or malformed keys
    private String sanitizeKey(String key) {
        String sanitized = key.replace("..", "").replaceAll("/{2,}", "/");
        return sanitized.startsWith("/") ? sanitized.substring(1) : sanitized;
    }

}
