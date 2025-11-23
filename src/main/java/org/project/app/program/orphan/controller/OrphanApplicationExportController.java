package org.project.app.program.orphan.controller;

import lombok.RequiredArgsConstructor;
import org.project.app.export.ExcelExportRequest;
import org.project.app.program.orphan.service.OrphanApplicationExportService;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.security.core.Authentication;
import org.springframework.web.bind.annotation.*;

import java.io.ByteArrayInputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;

@RequestMapping("/api/orphan/export")
@RequiredArgsConstructor
@RestController
public class OrphanApplicationExportController {
    private final OrphanApplicationExportService orphanApplicationExportService;

    @PostMapping
    public ResponseEntity<Resource> exportToExcel (Authentication currentUser,
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

                                                   @RequestBody ExcelExportRequest excelExportRequest) {

        ByteArrayInputStream excelStream = orphanApplicationExportService.findAllFilteredForExport(currentUser, status, createdBy, lastReviewedBy, createdStartDate, createdEndDate,
                lastModifiedStartDate, lastModifiedEndDate, dateOfBirthStartDate, dateOfBirthEndDate, id, fullName, bcRegistration, fathersName, gender,
                physicalCondition, permanentDistrict, permanentSubDistrict, sortField, sortDirection, excelExportRequest.getHeaders());

        InputStreamResource resource = new InputStreamResource(excelStream);
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=orphans.xlsx")
                .contentType(
                        MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                )
                .body(resource);
    }

}
