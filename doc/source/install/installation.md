# Installation

## 下载
- [Java 8](https://www.oracle.com/technetwork/java/javase/downloads/jdk8-downloads-2133151.html)
- [Maven 3.x](http://maven.apache.org/download.cgi)
- [Git 2.x](https://git-scm.com/downloads)
- [Tomcat 8](https://tomcat.apache.org/download-80.cgi)

## 设置
Windows中设置4个环境变量：`JAVA_HOME`、`CLASSPATH`、`M2_HOME`、`PATH`

### Java
- `JAVA_HOME`: `path/jdk1.8.0_131`
- `CLASSPATH`: `.;%JAVA_HOME%\lib;%JAVA_HOME%\lib\tools.jar`

### Maven
- `M2_HOME`: `path/apache-maven-3.3.9`
- 或者，设置IntelliJ IDEA 15 -> Settings -> Build Tools -> Maven -> Maven home directory

### Path
- `PATH`: `%JAVA_HOME%\bin;%M2_HOME%\bin;%PATH%`

### Version
- `java -version`
- `javac -version`
- `mvn -v`
- `git --version`

## 运行
```
git clone https://github.com/T5750/poi.git
cd poi
```

### Embedded Tomcat
```
mvn clean spring-boot:run
```

### Tomcat
```
mvn clean package
```
- 复制`target/poi.war`到`path/apache-tomcat-8.5.46/webapps`
- 运行`path/apache-tomcat-8.5.46/bin/startup.bat`
- [http://localhost:8080/poi](http://localhost:8080/poi)

### IDE
- `PoiApplication`