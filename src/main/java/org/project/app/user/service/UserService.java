package org.project.app.user.service;

import jakarta.ws.rs.NotFoundException;
import jakarta.ws.rs.core.Response;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.keycloak.admin.client.Keycloak;
import org.keycloak.admin.client.resource.GroupsResource;
import org.keycloak.admin.client.resource.UserResource;
import org.keycloak.representations.idm.*;
import org.project.app.config.KeycloakAdminConfig;
import org.project.app.enrollment.dto.ProgramEnrollmentDTO;
import org.project.app.exception.KeycloakCreationException;
import org.project.app.exception.KeycloakDeletionException;
import org.project.app.exception.UserDeletionException;
import org.project.app.exception.UserNotFoundException;
import org.project.app.program.ProgramDTO;
import org.project.app.storage.FileStorageService;
import org.project.app.user.domain.BeneficiaryExtra;
import org.project.app.enrollment.domain.ProgramEnrollment;
import org.project.app.user.dto.*;
import org.project.app.enrollment.repository.ProgramEnrollmentRepository;
import org.project.app.user.repository.UserExtraRepository;
import org.project.app.user.domain.UserExtra;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.*;
import java.util.stream.Collectors;

@Slf4j
@RequiredArgsConstructor
@Transactional
@Service
public class UserService {
    private final UserExtraRepository userRepository;
    private final UserMediaService mediaService;
    private final Keycloak keycloak;
    private final KeycloakAdminConfig config;


    public UserProfileResponse getUserProfile(String userId) {
        // Fetch keycloak user details
        UserProfileResponse response = fetchKeycloakUser(userId);

        // Fetch additional user data form DB
        Optional<UserExtra> userExtra = getUserExtraById(userId);

        userExtra.ifPresentOrElse(extra-> {
            response.setCell(extra.getCell());
            response.setAddress(extra.getAddress());

            // Add beneficiary data if exists
            if (extra.getBeneficiaryExtra() != null) {
                BeneficiaryExtra beneficiary = extra.getBeneficiaryExtra();

                response.setBCRegistration(beneficiary.getBcRegistration());
                response.setAccountTitle(beneficiary.getAccountTitle());
                response.setAccountNumber(beneficiary.getAccountNumber());
                response.setBankTitle(beneficiary.getBankTitle());
                response.setBranch(beneficiary.getBranch());
                response.setRoutingNumber(beneficiary.getRoutingNumber());

                // Map programs
                List<ProgramDTO> programDTOs = extra.getPrograms().stream()
                        .map(program -> {
                            ProgramEnrollment enrollment = program.getEnrollment(); // will be null if not GRANTED
                            return ProgramDTO.builder()
                                    .id(program.getId())
                                    .status(program.getStatus())
                                    .rejectionMessage(program.getRejectionMessage())
                                    .enrollment(enrollment != null ? ProgramEnrollmentDTO.builder()
                                            .id(enrollment.getId())
                                            .status(enrollment.getStatus())
                                            .monthlyAmount(enrollment.getMonthlyAmount())
                                            .isPaid(enrollment.isPaid())
                                            .enrolledAt(enrollment.getEnrolledAt())
                                            .lastUpdatedAt(enrollment.getLastUpdatedAt())
                                            .build() : null)
                                    .build();
                        })
                        .collect(Collectors.toList());

                response.setPrograms(programDTOs);
            }
        }, ()-> log.warn("User {} does not exist in database!", userId));

        return response;
    }

    private UserProfileResponse fetchKeycloakUser(String userId) {
        UserProfileResponse response = new UserProfileResponse();
        response.setUserId(userId);

        // Fetch user details
        UserResource userResource = keycloak.realm(config.getRealm()).users().get(userId);
        UserRepresentation userRep = userResource.toRepresentation();

        response.setUsername(userRep.getUsername());
        response.setEmail(userRep.getEmail());
        response.setEnabled(userRep.isEnabled());

        // Fetch user roles (both frontend and backend client roles)
        ClientRepresentation apiClient = keycloak.realm(config.getRealm()).clients().findByClientId("qc-api").stream().findFirst().orElse(null);
        ClientRepresentation uiClient = keycloak.realm(config.getRealm()).clients().findByClientId("qc-client").stream().findFirst().orElse(null);

        // Initialize roles
        List<String> apiRoles = new ArrayList<>();
        List<String> uiRoles = new ArrayList<>();

        if (apiClient != null) {
            List<RoleRepresentation> roles = userResource.roles()
                    .clientLevel(apiClient.getId()).listEffective();
            apiRoles = roles.stream().map(RoleRepresentation::getName).collect(Collectors.toList());
        }

        if (uiClient != null) {
            List<RoleRepresentation> roles = userResource.roles()
                    .clientLevel(uiClient.getId()).listEffective();
            uiRoles = roles.stream().map(RoleRepresentation::getName).collect(Collectors.toList());
        }

        // Set roles if any
        if (!apiRoles.isEmpty()) {
            response.setUserRoles(apiRoles);
        }

        if (!uiRoles.isEmpty()) {
            response.setUserPermissions(uiRoles);
        }

        return response;
    }

    public Optional<UserExtra> getUserExtraById(String userId) {
        return userRepository.findById(userId);
    }

    public String createUser(UserCreateRequest request) {
        // Step 1: create keycloak user
        String userId = createKeycloakUser(request);

        // Step 2: Assign user to groups (roles)
        assignUserGroups(userId, request.getGroups());

        // Step 3: create UserExtra and return userId
        return saveUserExtra(request, userId);
    }

    private String createKeycloakUser(UserCreateRequest request) {
        UserRepresentation userRep = getUserRepresentation(request);

        // Create user
        try (Response response = keycloak.realm(config.getRealm()).users().create(userRep)) {
            if (response.getStatus() != 201) {
                String error = response.readEntity(String.class);
                throw new RuntimeException("Keycloak user creation failed: " + error);
            }

            // Extract userID
            String location = response.getHeaderString("Location");
            return location.substring(location.lastIndexOf('/') + 1);
        } catch (RuntimeException e) {
            throw new KeycloakCreationException("Failed to create Keycloak user", e);
        }
    }

    private static UserRepresentation getUserRepresentation(UserCreateRequest request) {
        // Early validation for mandatory field
        if (request.getUsername() == null || request.getUsername().isBlank()) {
            throw new IllegalArgumentException("Username cannot be null or empty");
        }

        UserRepresentation userRep = new UserRepresentation();

        userRep.setUsername(request.getUsername());
        userRep.setEmail(request.getEmail());
        userRep.setEnabled(request.isEnabled());

        // Set credentials
        CredentialRepresentation credential = new CredentialRepresentation();
        credential.setType(CredentialRepresentation.PASSWORD);
        credential.setValue(request.getPassword());
        credential.setTemporary(false); //Force password change on first login
        userRep.setCredentials(Collections.singletonList(credential));
        return userRep;
    }

    private void assignUserGroups(String userId, List<String> groupNames) {
        // Early validation
        if (groupNames == null || groupNames.isEmpty()) return;


        UserResource userResource = keycloak.realm(config.getRealm()).users().get(userId);
        GroupsResource groupsResource = keycloak.realm(config.getRealm()).groups();

        // Get all available groups
        List<GroupRepresentation> allGroups = groupsResource.groups();
        Map<String, String> groupNameToId = allGroups.stream()
                .collect(Collectors.toMap(GroupRepresentation::getName, GroupRepresentation::getId));

        // Add user to each group
        for (String groupName : groupNames) {
            String groupId = groupNameToId.get(groupName);
            if (groupId == null) {
                log.warn("Group '{}' not found in Keycloak", groupName);
                continue;
            }
            userResource.joinGroup(groupId);
        }
    }

    public String updateUser(String userId, UserCreateRequest request) {
        // Step 1: update keycloak user
        updateKeycloakUser(userId, request);

        // Step 2: update user groups (roles)
        if (request.getGroups() != null) {
            updateUserGroups(userId, request.getGroups());
        }

        // Step 3: update userExtra and return userId
        return saveUserExtra(request, userId);
    }

    private void updateKeycloakUser(String userId, UserCreateRequest request) {
        UserResource userResource = keycloak.realm(config.getRealm()).users().get(userId);
        UserRepresentation userRep = userResource.toRepresentation();

        // Update all fields
        userRep.setUsername(request.getUsername());
        userRep.setEmail(request.getEmail());
        userRep.setEnabled(request.isEnabled());

        userResource.update(userRep);
    }

    public void updateUserGroups(String userId, List<String> newGroups) {
        UserResource userResource = keycloak.realm(config.getRealm()).users().get(userId);

        // Get current groups
        List<GroupRepresentation> currentGroups = userResource.groups();
        Map<String, String> currentGroupMap = currentGroups.stream()
                .collect(Collectors.toMap(GroupRepresentation::getName, GroupRepresentation::getId));

        // Get all available groups
        GroupsResource groupsResource = keycloak.realm(config.getRealm()).groups();
        List<GroupRepresentation> allGroups = groupsResource.groups();
        Map<String, String> allGroupMap = allGroups.stream()
                .collect(Collectors.toMap(GroupRepresentation::getName, GroupRepresentation::getId));

        // Remove from old groups not in new list
        currentGroupMap.keySet().stream()
                .filter(group -> !newGroups.contains(group))
                .forEach(group -> userResource.leaveGroup(allGroupMap.get(group)));

        // Add to new groups not in current list
        newGroups.stream()
                .filter(group -> !currentGroupMap.containsKey(group))
                .filter(allGroupMap::containsKey)
                .forEach(group -> userResource.joinGroup(allGroupMap.get(group)));
    }
    // Common function for saving/updating user business data

    public String saveUserExtra(UserCreateRequest request, String userId) {
        // Fetch userExtra or create new
        UserExtra userExtra = getUserExtraById(userId).orElse(new UserExtra(userId));

        // Update UserExtra fields
        userExtra.setCell(request.getCell());
        userExtra.setAddress(request.getAddress());

        if (request.getBeneficiaryCreateRequest() != null) {
            BeneficiaryCreateRequest beneficiaryRequest = request.getBeneficiaryCreateRequest();
            // Fetch beneficiaryExtra or create new
            BeneficiaryExtra beneficiaryExtra = userExtra.getBeneficiaryExtra() != null ?
                    userExtra.getBeneficiaryExtra() : new BeneficiaryExtra();

            if (userRepository.existsByBeneficiaryExtraBcRegistration(beneficiaryRequest.getBcRegistration())) {
                throw new DuplicateFormatFlagsException("User reference already exists");
            }
            beneficiaryExtra.setBcRegistration(beneficiaryRequest.getBcRegistration());
            beneficiaryExtra.setAccountTitle(beneficiaryRequest.getAccountTitle());
            beneficiaryExtra.setAccountNumber(beneficiaryRequest.getAccountNumber());
            beneficiaryExtra.setBankTitle(beneficiaryRequest.getBankTitle());
            beneficiaryExtra.setBranch(beneficiaryRequest.getBranch());
            beneficiaryExtra.setRoutingNumber(beneficiaryRequest.getRoutingNumber());

            userExtra.setBeneficiaryExtra(beneficiaryExtra);
        }

        userRepository.save(userExtra);
        return userExtra.getUserId();
    }

    public void updatePassword(String userId, PasswordUpdateRequest request) {
        CredentialRepresentation credential = new CredentialRepresentation();
        credential.setType(CredentialRepresentation.PASSWORD);
        credential.setValue(request.getUpdatedPassword());
        credential.setTemporary(false);

        keycloak.realm(config.getRealm()).users().get(userId).resetPassword(credential);
    }

    public void deleteUser(String userId) {
        try {
            // Step 1: delete Keycloak user
            deleteKeycloakUser(userId);

            // Step 2: delete media files and user data
            deleteApplicationData(userId);

            log.info("User with ID: {} deleted successfully", userId);
        } catch (NotFoundException e) {
            log.warn("Proceeding without Keycloak user: {}", e.getMessage());
            // Partial deletion if keycloak user doesn't exist
            try {
                deleteApplicationData(userId);
            } catch (Exception ex) {
                log.error("Failed to delete application data after missing Keycloak user for userId: {}", userId, ex);
                throw new UserDeletionException("Failed to delete application data after missing Keycloak user", ex);
            }
        } catch (KeycloakDeletionException e) {
            log.error("Keycloak deletion failed for user: {}", userId, e);
            throw new UserDeletionException("Partial deletion - Keycloak user not removed", e);
        } catch (Exception e) {
            log.error("Complete deletion failed for user: {}", userId, e);
            throw new UserDeletionException("Failed to delete user: " + userId, e);
        }
    }

    private void deleteApplicationData(String userId) {
        try {
            mediaService.deleteAllMediaForUser(userId);
            userRepository.delete(getUserExtraById(userId)
                    .orElseThrow(()-> new UserNotFoundException("User not found for deletion")));
        } catch (Exception e) {
            log.error("Error deleting application data for user: {}", userId, e);
            throw new UserDeletionException("Application data deletion failed", e);
        }
    }

    private void deleteKeycloakUser(String userId) {
        try (Response response = keycloak.realm(config.getRealm()).users().delete(userId)) {
            int status = response.getStatus();

            if (status == 404) {
                throw new NotFoundException("Keycloak user not found: " + userId);
            } else if (status != 204) {
                String error = response.readEntity(String.class);
                throw new KeycloakDeletionException("Keycloak API error [" + status + "]: " + error);
            }
        } catch (RuntimeException e) {
            throw new KeycloakDeletionException("Failed to delete Keycloak user", e);
        }
    }
}

