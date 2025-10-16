package org.project.app.program.orphan.controller;

import lombok.RequiredArgsConstructor;
import org.project.app.program.orphan.enums.DocumentType;
import org.project.app.program.orphan.service.OrphanMediaService;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.util.Map;

@RestController
@RequestMapping("/api/orphan/documents")
@RequiredArgsConstructor
public class OrphanMediaController {

    private final OrphanMediaService orphanMediaService;

    @GetMapping("/{id}")
    public ResponseEntity<Map<DocumentType, String>> getAllDocumentUrls(@PathVariable String id) {
        return ResponseEntity.ok(orphanMediaService.getAllDocumentUrls(id));
    }

    @GetMapping("/{id}/document/{docType}")
    public ResponseEntity<String> getDocumentUrl(@PathVariable String id, @PathVariable DocumentType docType) {
        return ResponseEntity.ok(orphanMediaService.getDocumentUrl(id, docType));
    }

    @PostMapping("/{id}/document/{docType}")
    public ResponseEntity<String> uploadDocument(@PathVariable String id, @PathVariable DocumentType docType, @RequestParam MultipartFile file) {
        orphanMediaService.uploadDocument(id, file, docType);
        return ResponseEntity.ok("Document uploaded");
    }
}

