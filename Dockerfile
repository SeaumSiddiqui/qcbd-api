# ==============================
# Stage 1 — Build the JAR
# ==============================
FROM maven:3.9.4-eclipse-temurin-21 AS build
LABEL authors="Seaum Siddiqui"

WORKDIR /app

# Cache dependencies
COPY pom.xml .
RUN mvn -B dependency:go-offline

# Copy sources and build
COPY src ./src
RUN mvn -B clean package -DskipTests

# ==============================
# Stage 2 — Run the app
# ==============================
FROM eclipse-temurin:21-jdk-alpine

# Create non-root user
RUN addgroup -S spring && adduser -S spring -G spring
USER spring:spring

WORKDIR /app

# Copy built artifact
COPY --from=build /app/target/*.jar app.jar

ENV SPRING_PROFILES_ACTIVE=prod
ENV TZ=Asia/Dhaka

EXPOSE 5050

ENTRYPOINT ["java", "-XX:+UseContainerSupport", "-Xmx512m", "-Xms256m", "-jar", "app.jar"]
