FROM openjdk:8
EXPOSE 8080
ARG JAR_FILE
ADD target/${JAR_FILE} /poi.jar
ENTRYPOINT ["java", "-jar","/poi.jar"]