package org.project.app.user.dto;

import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.Size;
import lombok.Data;

@Data
public class PasswordUpdateRequest {
    @NotBlank
    @Size(min = 8, message = "Password must be at least 8 characters long")
    private String updatedPassword;
}
