package org.project.app.enrollment.controller;

import lombok.RequiredArgsConstructor;
import org.project.app.enrollment.dto.ProgramEnrollmentUpdateDTO;
import org.project.app.enrollment.service.ProgramEnrollmentService;
import org.project.app.enrollment.domain.ProgramEnrollment;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RequestMapping("/api/enrollments")
@RequiredArgsConstructor
@RestController
public class ProgramEnrollmentController {
    private final ProgramEnrollmentService enrollmentService;

    @PutMapping("/{id}")
    public ResponseEntity<ProgramEnrollment> update(@PathVariable String id, @RequestBody ProgramEnrollmentUpdateDTO dto) {
        return ResponseEntity.ok(enrollmentService.updateEnrollment(id, dto));
    }

}

