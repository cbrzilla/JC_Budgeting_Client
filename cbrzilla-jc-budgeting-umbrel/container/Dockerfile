FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build
WORKDIR /src

COPY container/src/ ./

RUN dotnet publish ./JCBudgeting.Server/JCBudgeting.Server.csproj -c Release -o /out

FROM mcr.microsoft.com/dotnet/aspnet:10.0
WORKDIR /app
ENV ASPNETCORE_URLS=http://0.0.0.0:5099
ENV JCBUDGETING_SERVER_DATA_DIR=/data
EXPOSE 5099

COPY --from=build /out/ /app/

ENTRYPOINT ["dotnet", "/app/JCBudgeting.Server.dll"]
