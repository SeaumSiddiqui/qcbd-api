package org.project.app.exception;

public class KeycloakCreationException extends RuntimeException {
    public KeycloakCreationException(String message) {
        super(message);
    }
    public KeycloakCreationException(String message, Throwable cause) {
        super(message, cause);
    }
}
