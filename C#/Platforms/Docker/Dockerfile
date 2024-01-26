FROM mcr.microsoft.com/dotnet/runtime:8.0 AS base
WORKDIR /app

FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src
COPY ["DocumentDocker.csproj", ""]
RUN dotnet restore "./DocumentDocker.csproj"
COPY . .
WORKDIR "/src/."
RUN dotnet build "DocumentDocker.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "DocumentDocker.csproj" -c Release -o /app/publish /p:UseAppHost=false

FROM base AS final

# Update package sources to include supplemental packages (contrib archive area).
RUN sed -i 's/main/main contrib/g' /etc/apt/sources.list.d/debian.sources

# Downloads the package lists from the repositories.
RUN apt-get update

# Install font configuration.
RUN apt-get install -y fontconfig

# Install Microsoft TrueType core fonts.
RUN apt-get install -y ttf-mscorefonts-installer

# Or install Liberation TrueType fonts.
# RUN apt-get install -y fonts-liberation

# Or some other font package...

WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "DocumentDocker.dll"]