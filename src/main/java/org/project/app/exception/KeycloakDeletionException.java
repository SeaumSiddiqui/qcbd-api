package org.project.app.exception;

public class KeycloakDeletionException extends RuntimeException {
    public KeycloakDeletionException(String message) {
        super(message);
    }
    public KeycloakDeletionException(String message, Throwable cause) {
        super(message, cause);
    }
}
