FROM mcr.microsoft.com/dotnet/aspnet:10.0

WORKDIR /app
ENV ASPNETCORE_URLS=http://0.0.0.0:5099
ENV JCBUDGETING_SERVER_DATA_DIR=/data
EXPOSE 5099

COPY app/ /app/

RUN chmod +x /app/JCBudgeting.Server

ENTRYPOINT ["/app/JCBudgeting.Server"]
