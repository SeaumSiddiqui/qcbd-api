package org.project.app.user.dto;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.Data;
import org.project.app.program.ProgramDTO;

import java.util.List;

@Data
public class UserProfileResponse {
    private String userId;
    private String username;
    private String email;
    @JsonProperty("isEnabled")
    private boolean isEnabled;
    private List<String> userRoles;
    private List<String> userPermissions;

    private String cell;
    private String address;

    // Beneficiary fields
    private String BCRegistration;
    private String accountTitle;
    private String accountNumber;
    private String bankTitle;
    private String branch;
    private String routingNumber;

    private List<ProgramDTO> programs;
}
