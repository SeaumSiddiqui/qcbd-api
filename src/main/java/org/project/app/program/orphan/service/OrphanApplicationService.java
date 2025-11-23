package org.project.app.program.orphan.service;


import jakarta.persistence.OptimisticLockException;
import jakarta.transaction.Transactional;
import jakarta.validation.constraints.NotNull;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.modelmapper.ModelMapper;
import org.project.app.enrollment.service.ProgramEnrollmentService;
import org.project.app.exception.ApplicationDeletionException;
import org.project.app.exception.ResourceNotFoundException;
import org.project.app.exception.UserNotFoundException;
import org.project.app.program.orphan.domain.*;
import org.project.app.program.orphan.dto.*;
import org.project.app.program.orphan.enums.ApplicationStatus;
import org.project.app.program.orphan.repository.OrphanApplicationRepository;
import org.project.app.program.orphan.repository.OrphanApplicationSpecification;
import org.project.app.user.service.UserService;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.data.domain.Pageable;
import org.springframework.data.domain.Sort;
import org.springframework.data.jpa.domain.Specification;
import org.springframework.security.access.AccessDeniedException;
import org.springframework.security.core.Authentication;
import org.springframework.security.oauth2.server.resource.authentication.JwtAuthenticationToken;
import org.springframework.stereotype.Service;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

import static org.project.app.program.orphan.enums.ApplicationStatus.*;
import static org.project.app.security.SecurityConfig.*;

@Slf4j
@RequiredArgsConstructor
@Transactional
@Service
public class OrphanApplicationService {
    private final OrphanApplicationRepository applicationRepository;
    private final OrphanMediaService mediaService;
    private final ProgramEnrollmentService enrollmentService;
    private final UserService userService;
    private final ModelMapper modelMapper;

    public Specification<OrphanApplication> buildSpecification(Authentication currentUser, String status, String createdBy, String lastReviewedBy, LocalDateTime createdStartDate, LocalDateTime createdEndDate, LocalDateTime lastModifiedStartDate, LocalDateTime lastModifiedEndDate, LocalDate dateOfBirthStartDate, LocalDate dateOfBirthEndDate, String id, String fullName, String bcRegistration, String fathersName, String gender, String physicalCondition, String permanentDistrict, String permanentSubDistrict) {

        // Check if the user has ADMIN or MANAGEMENT role to determine access permissions.
        if(!isAdminOrManagement(currentUser)) createdBy = currentUser.getName();
        log.info("Requested username: {}", currentUser.getName());

        // Build specification
        return OrphanApplicationSpecification.buildSearchSpecification
                (status, createdBy, lastReviewedBy, createdStartDate, createdEndDate, lastModifiedStartDate,
                        lastModifiedEndDate, dateOfBirthStartDate, dateOfBirthEndDate, id, fullName, bcRegistration,
                        fathersName, gender, physicalCondition, permanentDistrict, permanentSubDistrict);
    }

    public Page<OrphanApplicationSummaryDTO> getAllApplication(Authentication currentUser, String status, String createdBy, String lastReviewedBy,
                                                               LocalDateTime createdStartDate, LocalDateTime createdEndDate,
                                                               LocalDateTime lastModifiedStartDate, LocalDateTime lastModifiedEndDate,
                                                               LocalDate dateOfBirthStartDate, LocalDate dateOfBirthEndDate, String id,
                                                               String fullName, String bcRegistration, String fathersName, String gender,
                                                               String physicalCondition, String permanentDistrict, String permanentSubDistrict,
                                                               String sortField, String sortDirection, int page, int size) {

        // Build specification
        Specification<OrphanApplication> searchSpecification = buildSpecification(currentUser, status, createdBy, lastReviewedBy, createdStartDate, createdEndDate, lastModifiedStartDate,
                        lastModifiedEndDate, dateOfBirthStartDate, dateOfBirthEndDate, id, fullName, bcRegistration,
                        fathersName, gender, physicalCondition, permanentDistrict, permanentSubDistrict);

        // Log the specification
        log.info("Search specification: {}", searchSpecification);
        // sort direction: (default)DESC, ASC
        Sort.Direction direction = Sort.Direction.fromString(sortDirection);

        // sort field: (default)createdAt, lastModifiedAt, status, physicalCondition
        Pageable pageable = PageRequest.of(page, size, Sort.by(direction, sortField));

        // Fetch and log the results
        Page<OrphanApplication> resultPage = applicationRepository.findAll(searchSpecification, pageable);
        log.info("Fetched {} records", resultPage.getTotalElements());

        return resultPage.map(this::formattedSummary);
        //return applicationRepository.findAll(searchSpecification, pageable).map(this::formattedSummary);
    }

    /**
     * Validates if the authenticated user has the role of ADMIN or MANAGEMENT.
     * Role USER has limited access to their own applications.
     * Roles ADMIN and MANAGEMENT have full access to all applications.
     */
    private boolean isAdminOrManagement(Authentication currentUser) {
        return getUserRole(currentUser).contains(QC_SERVER_ADMIN) || getUserRole(currentUser).contains(QC_SERVER_AUTHENTICATOR);
    }

    private List<String> getUserRole(Authentication currentUser) {
        // Extract claims form jwt
        JwtAuthenticationToken jwtAuth = (JwtAuthenticationToken) currentUser;
        Map<String, Object> claims = jwtAuth.getToken().getClaims();

        // Extract client roles from resource_access
        Map<String, Map<String, List<String>>> resourceAccess =
                (Map<String, Map<String, List<String>>>) claims.get("resource_access");
        // TODO-> move the client-id into yml prop
        return resourceAccess.get("qc-api").get("roles");
    }

    // TODO-> null check on field mapping
    private OrphanApplicationSummaryDTO formattedSummary(OrphanApplication application) {
        return OrphanApplicationSummaryDTO.builder()
                .id(application.getId())
                .fullName(application.getPrimaryInformation().getFullName())
                .fathersName(application.getPrimaryInformation().getFathersName())
                .dateOfBirth(application.getPrimaryInformation().getDateOfBirth())
                .gender(application.getPrimaryInformation().getGender())
                .permanentDistrict(application.getAddress().getPermanentDistrict())
                .status(application.getStatus())
                .build();
    }

    public OrphanApplication getApplicationById(String id) {
        return applicationRepository.findById(id)
                .orElseThrow(()-> {
                    log.warn("Application with ID ({}): not found", id);
                    return new ResourceNotFoundException("Application with ID: {%s} no found".formatted(id));
                });
    }

    public OrphanApplicationResponseDTO createApplication(OrphanApplicationCreateDTO dto, Authentication currentUser) {
        // Check permission
        if (!getUserRole(currentUser).contains(QC_SERVER_AGENT)) throw new AccessDeniedException("Insufficient permission for role: " + getUserRole(currentUser));
        //if (isAdminOrManagement(currentUser)) throw new AccessDeniedException("Creation forbidden for " + getUserRole(currentUser));;

        // Check if already exist
        String bcRegistration = dto.getPrimaryInformation().getBcRegistration();
        if (applicationRepository.existsByPrimaryInformationBcRegistration(bcRegistration)) {
            throw new DuplicateFormatFlagsException("An application for this user is already in progress");
        }

        // Map DTO to Entity
        OrphanApplication application = modelMapper.map(dto, OrphanApplication.class);

        // Set beneficiary
        application.setBeneficiary(userService.getUserExtraById(dto.getBeneficiaryUserId())
                .orElseThrow(()-> new UserNotFoundException("Beneficiary user not found")));

        return formattedResponse(applicationRepository.save(application));
    }

    private OrphanApplicationResponseDTO formattedResponse(OrphanApplication application) {
        return OrphanApplicationResponseDTO.builder()
                .id(application.getId())
                .status(application.getStatus())
                .rejectionMessage(application.getRejectionMessage())
                .version(application.getVersion())
                .createdBy(application.getCreatedBy())
                .lastReviewedBy(application.getLastReviewedBy())
                .createdAt(application.getCreatedAt())
                .lastModifiedAt(application.getLastModifiedAt())
                .primaryInformation(application.getPrimaryInformation())
                .address(application.getAddress())
                .familyMembers(application.getFamilyMembers())
                .basicInformation(application.getBasicInformation())
                .verification(application.getVerification())
                .build();
    }

    public OrphanApplication updateApplication(String id, OrphanApplicationUpdateDTO updateDTO, Authentication currentUser) {
        OrphanApplication existing = getApplicationById(id);

        verifyRolePermission(currentUser, existing.getStatus());
        checkVersion(existing, updateDTO.getVersion());

        // Update native fields
        modelMapper.map(updateDTO, existing);

        // Update nested fields
        updateNestedEntities(updateDTO, existing);

        // Update collections
        updateFamilyMembers(existing, updateDTO.getFamilyMembers());
        //updateDocuments(existing, updateDTO.getDocuments());

        return applicationRepository.save(existing);
    }

    private void updateNestedEntities(OrphanApplicationUpdateDTO updateDTO, OrphanApplication existing) {
        if (updateDTO.getStatus() == GRANTED && existing.getEnrollment() == null) {
            enrollmentService.createEnrollment(existing);
        }

        if (updateDTO.getPrimaryInformation() != null) {
            if (existing.getPrimaryInformation() == null) {
                existing.setPrimaryInformation(new PrimaryInformation());
            }
            modelMapper.map(updateDTO.getPrimaryInformation(), existing.getPrimaryInformation());
        }
        if (updateDTO.getAddress() != null) {
            if (existing.getAddress() == null) {
                existing.setAddress(new Address());
            }
            modelMapper.map(updateDTO.getAddress(), existing.getAddress());
        }
        if (updateDTO.getBasicInformation() != null) {
            if (existing.getBasicInformation() == null) {
                existing.setBasicInformation(new BasicInformation());
            }
            modelMapper.map(updateDTO.getBasicInformation(), existing.getBasicInformation());
        }
        if (updateDTO.getVerification() != null) {
            if (existing.getVerification() == null) {
                existing.setVerification(new Verification());
            }
            modelMapper.map(updateDTO.getVerification(), existing.getVerification());
        }
    }

    private void updateFamilyMembers(OrphanApplication app, List<FamilyMemberDTO> updates) {
        List<FamilyMember> existingMembers = app.getFamilyMembers();
        Set<String> updatedIds = new HashSet<>();

        // Handle null updates
        if (updates == null) updates = Collections.emptyList();

        // Process UPDATES (existing members with IDs)
        for (FamilyMemberDTO dto : updates) {
            if (dto.getId() != null) {
                existingMembers.stream()
                        .filter(m -> m.getId().equals(dto.getId()))
                        .findFirst()
                        .ifPresent(m -> modelMapper.map(dto, m));
                updatedIds.add(dto.getId());
            }
        }

        // Process ADDITIONS (new members without IDs)
        updates.stream()
                .filter(dto -> dto.getId() == null)
                .forEach(dto -> {
                    FamilyMember newMember = modelMapper.map(dto, FamilyMember.class);
                    newMember.setApplication(app);
                    existingMembers.add(newMember);
                });

        // Handle DELETIONS (only existing members with IDs not in updates)
        Iterator<FamilyMember> iterator = existingMembers.iterator();
        while (iterator.hasNext()) {
            FamilyMember member = iterator.next();
            String memberId = member.getId();
            if (memberId != null && !updatedIds.contains(memberId)) {
                member.setApplication(null);
                iterator.remove();
            }
        }
    }

    /**
     * Checks user permissions for updating applications based on roles and status.

     * - **USER Role**:
     *   - Can update their own applications.
     *   - Applicable only when the application's status is either **INCOMPLETE** **COMPLETE** or **REJECTED**.

     * - **MANAGEMENT Role**:
     *   - Can update all applications except those with an **INCOMPLETE** status.

     * - **ADMIN Role**:
     *   - Can update all applications regardless of their status.
     */
    private void verifyRolePermission(Authentication currentUser, @NotNull ApplicationStatus status) {
        if (getUserRole(currentUser).contains(QC_SERVER_ADMIN)) return;
        if (getUserRole(currentUser).contains(QC_SERVER_AUTHENTICATOR) && status.equals(PENDING)) return;
        if (getUserRole(currentUser).contains(QC_SERVER_AGENT) && (status.equals(INCOMPLETE) || (status.equals(COMPLETE)) || (status.equals(REJECTED)))) return;

        throw new AccessDeniedException("Update forbidden for " + getUserRole(currentUser) + " on status " + status);
    }

    private void checkVersion(OrphanApplication existing, @NotNull Long version) {
        if (!existing.getVersion().equals(version)) {
            throw new OptimisticLockException("Application has been modified by another user");
        }
    }

    public void deleteApplication(String id) {

        try {
            mediaService.deleteAllMediaForApplication(id);
            applicationRepository.delete(getApplicationById(id));

            log.info("Application with ID: {} deleted", id);
        } catch (Exception e) {
            throw new ApplicationDeletionException("Failed to delete application: " + id, e);
        }
    }

}

