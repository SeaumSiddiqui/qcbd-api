package org.project.app.enrollment.service;

import jakarta.transaction.Transactional;
import jakarta.ws.rs.NotFoundException;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.modelmapper.ModelMapper;
import org.project.app.enrollment.dto.ProgramEnrollmentUpdateDTO;
import org.project.app.enrollment.enums.ProgramStatus;
import org.project.app.enrollment.repository.ProgramEnrollmentRepository;
import org.project.app.program.orphan.domain.OrphanApplication;
import org.project.app.enrollment.domain.ProgramEnrollment;
import org.springframework.stereotype.Service;

@Slf4j
@RequiredArgsConstructor
@Transactional
@Service
public class ProgramEnrollmentService {
    private final ProgramEnrollmentRepository programEnrollmentRepository;
    private final ModelMapper modelMapper;

    public ProgramEnrollment updateEnrollment(String enrollmentId, ProgramEnrollmentUpdateDTO dto) {
        ProgramEnrollment enrollment = programEnrollmentRepository.findById(enrollmentId)
                .orElseThrow(() -> new NotFoundException("Enrollment not found"));

        enrollment.setStatus(dto.getStatus());
        enrollment.setMonthlyAmount(dto.getMonthlyAmount());
        enrollment.setPaid(dto.isPaid());
        return programEnrollmentRepository.save(enrollment);
    }

    public void createEnrollment(OrphanApplication existing) {
        ProgramEnrollment enrollment = ProgramEnrollment.builder()
                .program(existing)
                .beneficiary(existing.getBeneficiary())
                .status(ProgramStatus.ACTIVE)
                .build();

        existing.setEnrollment(programEnrollmentRepository.save(enrollment));
    }

}