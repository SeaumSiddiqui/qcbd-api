package org.project.app.program.orphan.dto;

import lombok.Data;

@Data
public class VerificationDTO {
    private String fillerSignatureUrl;
    private String examinerSignatureUrl;
    private String investigatorSignatureUrl;
    private String qcSwdSignatureUrl;
}
