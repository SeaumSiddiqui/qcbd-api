package org.project.app.program.orphan.service;

import jakarta.transaction.Transactional;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.project.app.export.ExcelExportService;
import org.project.app.program.orphan.domain.OrphanApplication;
import org.project.app.program.orphan.repository.OrphanApplicationRepository;
import org.project.app.program.orphan.repository.OrphanApplicationSpecification;
import org.springframework.data.domain.Sort;
import org.springframework.data.jpa.domain.Specification;
import org.springframework.security.core.Authentication;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;

@Slf4j
@RequiredArgsConstructor
@Transactional
@Service
public class OrphanApplicationExportService {
    private final ExcelExportService exportService;
    private final OrphanApplicationService applicationService;
    private final OrphanApplicationRepository applicationRepository;

    public ByteArrayInputStream findAllFilteredForExport(Authentication currentUser, String status, String createdBy, String lastReviewedBy, LocalDateTime createdStartDate, LocalDateTime createdEndDate, LocalDateTime lastModifiedStartDate, LocalDateTime lastModifiedEndDate, LocalDate dateOfBirthStartDate, LocalDate dateOfBirthEndDate, String id, String fullName, String bcRegistration, String fathersName, String gender, String physicalCondition, String permanentDistrict, String permanentSubDistrict, String sortField, String sortDirection, List<String> headers) {
        // Build specification
        Specification<OrphanApplication> searchSpecification = applicationService.buildSpecification(currentUser, status, createdBy, lastReviewedBy, createdStartDate, createdEndDate, lastModifiedStartDate,
                        lastModifiedEndDate, dateOfBirthStartDate, dateOfBirthEndDate, id, fullName, bcRegistration,
                        fathersName, gender, physicalCondition, permanentDistrict, permanentSubDistrict);

        // Log the specification
        log.info("Search specification: {}", searchSpecification);

        // sort direction: (default)DESC, ASC
        Sort.Direction direction = Sort.Direction.fromString(sortDirection);
        Sort sort = Sort.by(direction, sortField);

        List<OrphanApplication> data = applicationRepository.findAll(searchSpecification, sort);
        log.info("Export rows count: {}", data.size());

        return exportService.exportAsStream(data, headers);
    }
}
