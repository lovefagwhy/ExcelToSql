@echo ***********************************开始Excel转Sql程序***********************************
::@echo on
set JAVA_HOME=Java\jdk1.7.0_79\jre
set CLASSPATH=.;%JAVA_HOME%\lib\dt.jar;%JAVA_HOME%\lib\tools.jar;  
set PATH=%JAVA_HOME%\bin;
java -jar ExcelToSql.jar
@pause