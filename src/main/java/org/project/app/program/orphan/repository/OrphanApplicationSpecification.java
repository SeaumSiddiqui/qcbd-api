package org.project.app.program.orphan.repository;


import jakarta.persistence.criteria.Join;
import jakarta.persistence.criteria.JoinType;
import jakarta.persistence.criteria.Root;
import org.project.app.program.orphan.domain.Address;
import org.project.app.program.orphan.domain.BasicInformation;
import org.project.app.program.orphan.domain.OrphanApplication;
import org.project.app.program.orphan.domain.PrimaryInformation;
import org.project.app.program.orphan.enums.ApplicationStatus;
import org.project.app.program.orphan.enums.Gender;
import org.project.app.program.orphan.enums.PhysicalCondition;
import org.springframework.data.jpa.domain.Specification;

import java.time.LocalDate;
import java.time.LocalDateTime;

public class OrphanApplicationSpecification {

    public static Specification<OrphanApplication> id(String id) {
        return (root, query, criteriaBuilder) ->
                id != null ? criteriaBuilder.equal(root.get("id"), id) : criteriaBuilder.conjunction();
    }

    public static Specification<OrphanApplication> status(String applicationStatus) {
        return (root, query, criteriaBuilder) -> {
            if (applicationStatus != null) {
                ApplicationStatus status = ApplicationStatus.valueOf(applicationStatus.toUpperCase());
                return criteriaBuilder.equal(root.get("status"), status);
            }
            return criteriaBuilder.conjunction();
        };
    }

    public static Specification<OrphanApplication> createdBy(String createdBy) {
        return (root, query, criteriaBuilder) ->
                createdBy != null ? criteriaBuilder.equal(root.get("createdBy"), createdBy) : criteriaBuilder.conjunction();
    }

    public static Specification<OrphanApplication> lastReviewedBy(String lastReviewedBy) {
        return (root, query, criteriaBuilder) ->
                lastReviewedBy != null ? criteriaBuilder.equal(root.get("lastReviewedBy"), lastReviewedBy) : criteriaBuilder.conjunction();
    }

    public static Specification<OrphanApplication> createdAt(LocalDateTime start, LocalDateTime end) {
        return (root, query, criteriaBuilder) -> {
            if (start != null && end != null) {
                return criteriaBuilder.between(root.get("createdAt"), start, end);
            } else if (start != null) {
                return criteriaBuilder.greaterThanOrEqualTo(root.get("createdAt"), start);
            } else if (end != null) {
                return criteriaBuilder.lessThanOrEqualTo(root.get("createdAt"), end);
            }
            return criteriaBuilder.conjunction();
        };
    }

    public static Specification<OrphanApplication> lastModifiedAt(LocalDateTime start, LocalDateTime end) {
        return (root, query, criteriaBuilder) -> {
            if (start != null && end != null) {
                return criteriaBuilder.between(root.get("lastModifiedAt"), start, end);
            } else if (start != null) {
                return criteriaBuilder.greaterThanOrEqualTo(root.get("lastModifiedAt"), start);
            } else if (end != null) {
                return criteriaBuilder.lessThanOrEqualTo(root.get("lastModifiedAt"), end);
            }
            return criteriaBuilder.conjunction();
        };
    }

    private static Join<OrphanApplication, PrimaryInformation> joinPrimaryInfo(Root<OrphanApplication> root) {
        return root.join("primaryInformation", JoinType.LEFT);
    }

    public static Specification<OrphanApplication> personalInformationDateOfBirth(LocalDate dobStart, LocalDate dobEnd) {
        return (root, query, criteriaBuilder) -> {
            Join<OrphanApplication, PrimaryInformation> personalInfoJoin = joinPrimaryInfo(root);
            if (dobStart != null && dobEnd != null) {
                return criteriaBuilder.between(personalInfoJoin.get("dateOfBirth"), dobStart, dobEnd);
            } else if (dobStart != null) {
                return criteriaBuilder.greaterThanOrEqualTo(personalInfoJoin.get("dateOfBirth"), dobStart);
            } else if (dobEnd != null) {
                return criteriaBuilder.lessThanOrEqualTo(personalInfoJoin.get("dateOfBirth"), dobEnd);
            }
            return criteriaBuilder.conjunction();
        };
    }

    public static Specification<OrphanApplication> personalInformationFullName(String fullName) {
        return (root, query, criteriaBuilder) -> {
            if (fullName != null) {
                Join<OrphanApplication, PrimaryInformation> personalInfoJoin = joinPrimaryInfo(root);
                return criteriaBuilder.like(criteriaBuilder.lower(personalInfoJoin.get("fullName")), "%" + fullName.toLowerCase() + "%");
            }
            return criteriaBuilder.conjunction();
        };
    }

    public static Specification<OrphanApplication> personalInformationBcRegistration(String bcRegistration) {
        return (root, query, criteriaBuilder) -> {
            if (bcRegistration != null) {
                Join<OrphanApplication, PrimaryInformation> personalInfoJoin = joinPrimaryInfo(root);
                return criteriaBuilder.like(personalInfoJoin.get("bcRegistration"), String.format("%%%s%%", bcRegistration));
            }
            return criteriaBuilder.conjunction();
        };
    }

    public static Specification<OrphanApplication> personalInformationFathersName(String fathersName) {
        return (root, query, criteriaBuilder) -> {
            if (fathersName != null) {
                Join<OrphanApplication, PrimaryInformation> personalInfoJoin = joinPrimaryInfo(root);
                return criteriaBuilder.like(criteriaBuilder.lower(personalInfoJoin.get("fathersName")), "%" + fathersName.toLowerCase() + "%");
            }
            return criteriaBuilder.conjunction();
        };
    }

    public static Specification<OrphanApplication> personalInformationGender(String genderSpec) {
        return (root, query, criteriaBuilder) -> {
            if (genderSpec == null || genderSpec.isEmpty()) {
                return criteriaBuilder.conjunction();
            }

            try {
                // Convert input to enum
                Gender gender = Gender.valueOf(genderSpec.toUpperCase());

                Join<OrphanApplication, PrimaryInformation> personalInfoJoin = joinPrimaryInfo(root);
                return criteriaBuilder.equal(
                        personalInfoJoin.get("gender"),
                        gender
                );
            } catch (IllegalArgumentException e) {
                return criteriaBuilder.disjunction();
            }
        };
    }

    private static Join<OrphanApplication, BasicInformation> joinBasicInfo(Root<OrphanApplication> root) {
        return root.join("basicInformation", JoinType.LEFT);
    }
    public static Specification<OrphanApplication> basicInformationPhysicalCondition(String condition) {
        return (root, query, criteriaBuilder) -> {
            if (condition == null || condition.isEmpty()) {
                return criteriaBuilder.conjunction();
            }

            try {
                // Convert input to enum
                PhysicalCondition physicalCondition = PhysicalCondition.valueOf(condition.toUpperCase());

                Join<OrphanApplication, BasicInformation> basicInfoJoin = joinBasicInfo(root);
                return criteriaBuilder.equal(
                        basicInfoJoin.get("physicalCondition"),
                        physicalCondition
                );
            } catch (IllegalArgumentException e) {
                return criteriaBuilder.disjunction();
            }
        };
    }

    private static Join<OrphanApplication, Address> joinAddress(Root<OrphanApplication> root) {
        return root.join("address", JoinType.LEFT);
    }

    public static Specification<OrphanApplication> addressPermanentDistrict(String permanentDistrict) {
        return (root, query, criteriaBuilder) -> {
            if (permanentDistrict != null) {
                Join<OrphanApplication, Address> addressJoin = joinAddress(root);
                return criteriaBuilder.like(addressJoin.get("permanentDistrict"), String.format("%%%s%%", permanentDistrict));
            }
            return criteriaBuilder.conjunction();
        };
    }

    public static Specification<OrphanApplication> addressPermanentSubDistrict(String permanentSubDistrict) {
        return (root, query, criteriaBuilder) -> {
            if (permanentSubDistrict != null) {
                Join<OrphanApplication, Address> addressJoin = joinAddress(root);
                return criteriaBuilder.like(addressJoin.get("permanentSubDistrict"), String.format("%%%s%%", permanentSubDistrict));
            }
            return criteriaBuilder.conjunction();
        };
    }

    public static Specification<OrphanApplication> buildSearchSpecification(String status, String createdBy, String lastReviewedBy, LocalDateTime createdStartDate, LocalDateTime createdEndDate, LocalDateTime lastModifiedStartDate, LocalDateTime lastModifiedEndDate, LocalDate dateOfBirthStartDate, LocalDate dateOfBirthEndDate, String id, String fullName, String bcRegistration, String fathersName, String gender, String physicalCondition, String permanentDistrict, String permanentSubDistrict) {
        return Specification.where(status(status)
                .and(createdBy(createdBy))
                .and(lastReviewedBy(lastReviewedBy))
                .and(createdAt(createdStartDate, createdEndDate))
                .and(lastModifiedAt(lastModifiedStartDate, lastModifiedEndDate))
                .and(personalInformationDateOfBirth(dateOfBirthStartDate, dateOfBirthEndDate))
                .and(id(id))
                .and(personalInformationFullName(fullName))
                .and(personalInformationBcRegistration(bcRegistration))
                .and(personalInformationFathersName(fathersName))
                .and(personalInformationGender(gender))
                .and(basicInformationPhysicalCondition(physicalCondition))
                .and(addressPermanentDistrict(permanentDistrict))
                .and(addressPermanentSubDistrict(permanentSubDistrict))
        );
    }

}

