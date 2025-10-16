package org.project.app.user.dto;

import com.fasterxml.jackson.annotation.JsonProperty;
import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.Size;
import lombok.Data;

import java.util.List;

@Data
public class UserCreateRequest {
    private String username;
    private String email;
    @NotBlank
    @Size(min = 8, message = "Password must be at least 8 characters long")
    private String password;
    @JsonProperty("isEnabled")
    private boolean isEnabled;
    private List<String> groups;

    private String cell;
    private String address;

    private BeneficiaryCreateRequest beneficiaryCreateRequest;

}
