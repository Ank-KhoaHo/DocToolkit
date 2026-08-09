# Build somewhere with an SDK, run somewhere with almost nothing.
#
# The runtime stage is the plain `runtime` image - no SDK, no LibreOffice, no browser, no fonts, and
# none of the native libraries an imaging or PDF stack would normally expect. A dependency that
# shipped a .so would fail to start here, which is the point of running this at all.
#
# NOT `runtime-deps`: that image ships no .NET runtime at all, so it can only host a
# self-contained publish. Pointing a framework-dependent app at it fails at startup, which is a
# confusing way to learn the difference.
#
# Build and run from the repository root:
#   docker build -f samples/Container/Dockerfile -t doctoolkit-container .
#   docker run --rm --network none doctoolkit-container
#
# `--network none` is worth using: the offline guarantee is not a claim you have to take on trust.

FROM mcr.microsoft.com/dotnet/sdk:10.0.302 AS build
WORKDIR /src

# Only what this sample needs. Copying the whole repository would drag in the solution and every
# other project, and this image is meant to prove a CONSUMER's experience, not ours.
COPY samples/Container/Container.csproj samples/Container/
COPY samples/Directory.Build.props samples/
COPY Directory.Build.props ./
COPY global.json ./
RUN dotnet restore samples/Container/Container.csproj

COPY samples/Container/ samples/Container/
RUN dotnet publish samples/Container/Container.csproj -c Release -o /app --no-restore

FROM mcr.microsoft.com/dotnet/runtime:8.0 AS runtime
WORKDIR /app
COPY --from=build /app .

# Invariant globalization: runs without ICU, which is one more native dependency this package does
# not need. Not required by DocToolkit - it simply makes the image prove a bit more.
ENV DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=1

ENTRYPOINT ["./Container"]
