FROM java:8-alpine
ADD easypoi-demo-0.0.1-SNAPSHOT.jar  app.jar
EXPOSE 8080
#启动命令
#ENTRYPOINT ["java"," -jar ","/msa-tenant-service-0.0.1-SNAPSHOT-docker.jar"]
CMD java -jar /app.jar --spring.profiles.active=docker
