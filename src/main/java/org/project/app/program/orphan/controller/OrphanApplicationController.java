package org.project.app.program.orphan.controller;

import lombok.RequiredArgsConstructor;
import org.project.app.program.orphan.domain.OrphanApplication;
import org.project.app.program.orphan.dto.OrphanApplicationCreateDTO;
import org.project.app.program.orphan.dto.OrphanApplicationResponseDTO;
import org.project.app.program.orphan.dto.OrphanApplicationSummaryDTO;
import org.project.app.program.orphan.dto.OrphanApplicationUpdateDTO;
import org.project.app.program.orphan.service.OrphanApplicationService;
import org.springframework.data.domain.Page;
import org.springframework.http.ResponseEntity;
import org.springframework.security.core.Authentication;
import org.springframework.web.bind.annotation.*;

import java.time.LocalDate;
import java.time.LocalDateTime;

@RequestMapping("/api/applications/orphan")
@RequiredArgsConstructor
@RestController
public class OrphanApplicationController {
    private final OrphanApplicationService orphanService;


    @GetMapping
    public ResponseEntity<Page<OrphanApplicationSummaryDTO>> getAllApplication (Authentication currentUser,
                                                                                @RequestParam(required = false) String status,
                                                                                @RequestParam(required = false) String createdBy,
                                                                                @RequestParam(required = false) String lastReviewedBy,

                                                                                @RequestParam(required = false) LocalDateTime createdStartDate,
                                                                                @RequestParam(required = false) LocalDateTime createdEndDate,
                                                                                @RequestParam(required = false) LocalDateTime lastModifiedStartDate,
                                                                                @RequestParam(required = false) LocalDateTime lastModifiedEndDate,
                                                                                @RequestParam(required = false) LocalDate dateOfBirthStartDate,
                                                                                @RequestParam(required = false) LocalDate dateOfBirthEndDate,

                                                                                @RequestParam(required = false) String id,
                                                                                @RequestParam(required = false) String fullName,
                                                                                @RequestParam(required = false) String bcRegistration,
                                                                                @RequestParam(required = false) String fathersName,
                                                                                @RequestParam(required = false) String gender,
                                                                                @RequestParam(required = false) String physicalCondition,

                                                                                @RequestParam(required = false) String permanentDistrict,
                                                                                @RequestParam(required = false) String permanentSubDistrict,

                                                                                @RequestParam(required = false, defaultValue = "createdAt") String sortField,
                                                                                @RequestParam(required = false, defaultValue = "DESC") String sortDirection,
                                                                                @RequestParam(defaultValue = "0")int page,
                                                                                @RequestParam(defaultValue = "10") int size) {
        return ResponseEntity.ok(orphanService.getAllApplication(
                currentUser, status, createdBy, lastReviewedBy, createdStartDate, createdEndDate, lastModifiedStartDate,
                lastModifiedEndDate, dateOfBirthStartDate, dateOfBirthEndDate, id, fullName, bcRegistration, fathersName,
                gender, physicalCondition, permanentDistrict, permanentSubDistrict, sortField, sortDirection, page, size));
    }

    @GetMapping("/{id}")
    public ResponseEntity<OrphanApplication> getApplicationById(@PathVariable String id) {
        return ResponseEntity.ok(orphanService.getApplicationById(id));
    }

    @PostMapping
    public ResponseEntity<OrphanApplicationResponseDTO> createApplication(@RequestBody OrphanApplicationCreateDTO application, Authentication currentUser) {
        return ResponseEntity.ok(orphanService.createApplication(application, currentUser));
    }

    @PutMapping("/{id}")
    public ResponseEntity<OrphanApplication> updateApplication(@PathVariable String id, @RequestBody OrphanApplicationUpdateDTO application, Authentication currentUser) {
        return ResponseEntity.ok(orphanService.updateApplication(id, application, currentUser));
    }

    @DeleteMapping("/{id}")
    public ResponseEntity<String> deleteApplication(@PathVariable String id) {
        orphanService.deleteApplication(id);
        return ResponseEntity.ok("Deleted!");
    }

}
