package org.project.app.user.domain;

import jakarta.persistence.Embeddable;
import lombok.Data;

@Embeddable
@Data
public class BeneficiaryExtra {
    private String bcRegistration;
    private String accountTitle;
    private String accountNumber;
    private String bankTitle;
    private String branch;
    private String routingNumber;
}