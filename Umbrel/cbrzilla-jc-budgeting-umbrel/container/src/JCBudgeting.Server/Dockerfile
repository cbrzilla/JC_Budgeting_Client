FROM mcr.microsoft.com/dotnet/aspnet:10.0 AS base
WORKDIR /app
ENV ASPNETCORE_URLS=http://0.0.0.0:5099
EXPOSE 5099

COPY . .

ENTRYPOINT ["./JCBudgeting.Server"]
