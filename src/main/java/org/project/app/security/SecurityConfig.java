package org.project.app.security;

import lombok.RequiredArgsConstructor;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.security.config.annotation.web.builders.HttpSecurity;
import org.springframework.security.config.http.SessionCreationPolicy;
import org.springframework.security.web.SecurityFilterChain;

import static org.springframework.http.HttpMethod.*;

@RequiredArgsConstructor
@Configuration
public class SecurityConfig {
    public static final String QC_SERVER_ADMIN = "api-admin";
    public static final String QC_SERVER_AUTHENTICATOR = "api-authenticator";
    public static final String QC_SERVER_AGENT = "api-agent";
    public static final String QC_SERVER_USER= "api-user";

    private final JwtAuthConverter jwtAuthConverter;


    @Bean
    SecurityFilterChain securityFilterChain(HttpSecurity http) throws Exception {

        return http
                .authorizeHttpRequests(request-> request
                        .requestMatchers("/api/user-extra/**").hasAnyRole(QC_SERVER_ADMIN, QC_SERVER_AUTHENTICATOR, QC_SERVER_AGENT)
//                        .requestMatchers(GET, "/api/OrphanApplication-applications/**").hasAnyRole(QC_SERVER_ADMIN, QC_SERVER_AUTHENTICATOR, QC_SERVER_AGENT)
//                        .requestMatchers(PUT, "/api/OrphanApplication-applications/**").hasAnyRole(QC_SERVER_ADMIN, QC_SERVER_AUTHENTICATOR, QC_SERVER_AGENT)
//                        .requestMatchers(POST, "/api/OrphanApplication-applications/**").hasAnyRole(QC_SERVER_ADMIN, QC_SERVER_AGENT)
//                        .requestMatchers(DELETE, "/api/OrphanApplication-applications/**").hasRole(QC_SERVER_ADMIN)
                        .anyRequest().permitAll()

                )
                .oauth2ResourceServer(oauth2ResourceServer ->
                        oauth2ResourceServer.jwt(jwt -> jwt.jwtAuthenticationConverter(jwtAuthConverter)))
                .sessionManagement(management-> management.sessionCreationPolicy(SessionCreationPolicy.STATELESS))
                .build();
    }

}
