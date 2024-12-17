# Use OpenJDK as the base image
FROM openjdk:17-jdk-bullseye

# Install LibreOffice
RUN apt-get update && apt-get install -y libreoffice

# Set the working directory
WORKDIR /app

# Copy application JAR and resource files
COPY build/libs/docxEditor_java-1.0-SNAPSHOT-all.jar app.jar
COPY src/main/resources ./resources
COPY src/main/resources/input.docx /app/src/main/resources/input.docx


# Expose port (if needed, optional)
EXPOSE 8080

# Run the application
ENTRYPOINT ["java", "-jar", "app.jar"]
