# Installation

Windows中设置4个环境变量：`JAVA_HOME`、`CLASSPATH`、`M2_HOME`、`PATH`

## Java
- `JAVA_HOME`: `path/jdk1.8.0_131`
- `CLASSPATH`: `.;%JAVA_HOME%\lib;%JAVA_HOME%\lib\tools.jar`

## Maven
- `M2_HOME`: `path/apache-maven-3.3.9`
- 或者，设置IntelliJ IDEA 15 -> Settings -> Build Tools -> Maven -> Maven home directory

## Path
- `PATH`: `%JAVA_HOME%\bin;%M2_HOME%\bin;%PATH%`

## Version
- `java -version`
- `javac -version`
- `mvn -v`