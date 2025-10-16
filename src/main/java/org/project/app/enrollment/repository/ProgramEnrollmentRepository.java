package org.project.app.enrollment.repository;

import org.project.app.enrollment.domain.ProgramEnrollment;
import org.springframework.data.jpa.repository.JpaRepository;

public interface ProgramEnrollmentRepository extends JpaRepository<ProgramEnrollment, String> {

//    List<ProgramEnrollment> findByBeneficiaryUserId(String userId);
//
//    @Query("SELECT e FROM ProgramEnrollment e JOIN FETCH e.application WHERE e.beneficiary.userId = :userId")
//    List<ProgramEnrollment> findByBeneficiaryUserIdWithApplications(@Param("userId") String userId);
}
