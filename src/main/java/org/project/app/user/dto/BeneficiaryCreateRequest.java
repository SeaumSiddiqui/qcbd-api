package org.project.app.user.dto;

import lombok.Data;

@Data
public class BeneficiaryCreateRequest {
    private String bcRegistration;
    private String accountTitle;
    private String accountNumber;
    private String bankTitle;
    private String branch;
    private String routingNumber;
}
