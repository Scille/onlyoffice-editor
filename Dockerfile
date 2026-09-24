# Use this (large) base image in every build below, to reduce the overall docker cache size
FROM ubuntu:26.04 AS base

# Workaround for slow archive.ubuntu.com
RUN sed -i 's|archive.ubuntu.com|ftp.halifax.rwth-aachen.de|g' /etc/apt/sources.list.d/ubuntu.sources

RUN apt-get update && apt-get install -y openjdk-21-jdk npm wget brotli make bzip2
RUN wget -qO- https://get.pnpm.io/install.sh | ENV="$HOME/.bashrc" SHELL="$(which bash)" bash -
ENV PNPM_HOME="/root/.local/share/pnpm"
ENV PATH="$PNPM_HOME/bin:$PATH"
RUN pnpm env use --global 20
RUN pnpm install -g grunt

###################### onlyoffice-editor-build ################################
FROM base AS onlyoffice-editor-build
WORKDIR /app
COPY onlyoffice-editor/package.json /app
COPY onlyoffice-editor/pnpm-lock.yaml /app
RUN pnpm install --dangerously-allow-all-builds
COPY onlyoffice-editor/tsconfig.json /app
COPY onlyoffice-editor/webpack.config.mjs /app
COPY onlyoffice-editor/src/ /app/src
RUN pnpm build

FROM onlyoffice-editor-build AS onlyoffice-editor-test
COPY onlyoffice-editor/eslint.config.mjs /app
COPY onlyoffice-editor/.prettierignore /app
COPY .editorconfig /app
RUN pnpm lint


###################### sdkjs & web-apps ################################
FROM base AS sdkjs-build
COPY sdkjs /app/sdkjs
COPY web-apps /app/web-apps
COPY fonts/*.png /app/sdkjs/common/Images
COPY fonts/*.js /app/sdkjs/common
WORKDIR /app/sdkjs
RUN make
RUN mv deploy/web-apps/apps/api/documents/api.js deploy/web-apps/apps/api/documents/api-orig.js


FROM base AS files-build
COPY --from=sdkjs-build /app/sdkjs/deploy/web-apps /app/web-apps
COPY --from=sdkjs-build /app/sdkjs/deploy/sdkjs /app/sdkjs
COPY vendor /app/web-apps/vendor
COPY fonts/*.ttf /app/fonts/fonts/
COPY fonts/*.otf /app/fonts/fonts/
COPY dictionaries /app/dictionaries
COPY --from=onlyoffice-editor-build /app/dist/api.js /app/web-apps/apps/api/documents/api.js
COPY --from=onlyoffice-editor-build /app/dist/index.d.ts /app/index.d.ts
# Add minimal `package.json` to the release
ARG release_version
RUN test -n "$release_version" || { echo "release_version is required (for example: make build release_version=1.2.3)" >&2; exit 1; } \
 && printf '%s\n' \
    '{' \
    '  "name": "onlyoffice-editor",' \
    "  \"version\": \"$release_version\"," \
    '  "description": "OnlyOffice editor with end-to-end encryption support",' \
    '  "license": "AGPL-3.0-or-later",' \
    '  "private": true,' \
    '  "files": [' \
    '    "dictionaries",' \
    '    "fonts",' \
    '    "sdkjs",' \
    '    "web-apps",' \
    '    "index.d.ts"' \
    '  ],' \
    '  "types": "index.d.ts"' \
    '}' > /app/package.json
# Skip by default the built-in editor documentation (apps/*/main/resources/help).
# It is ~500MB (mostly per-language PNG screenshots) and is only used by the editors'
# "?" help dialog (that can itself be disabled at runtime by setting `help: false`
# in the option when opening the editor).
ARG INCLUDE_HELP=false
RUN if [ "$INCLUDE_HELP" != "true" ]; then rm -rf /app/web-apps/apps/*/main/resources/help; fi
WORKDIR /app


FROM files-build AS package-build
# Precompressing assets with Brotli (so that the hosting server doesn't have to
# run it on the fly). This adds ~100MB as the original source files must still
# be included to support web clients that don't use `Accept-Encoding: br`.
ARG PRECOMPRESSED_ASSETS=false
RUN if [ "$INCLUDE_HELP" != "true" ]; then \
find . -name "*.wasm" \
    -o -name "*.js" \
    -o -name "*.html" \
    -o -name "*.css" \
    -o -name "*.aff" \
    -o -name "*.dic" \
    | xargs -P 8 -n 16 -- brotli \
; fi
RUN package_archive="$(npm pack --ignore-scripts --silent)" \
 && sha512sum "$package_archive" > "$package_archive.sha512"



FROM package-build AS package-test
RUN tar -tzf onlyoffice-editor-*.tgz > package.content
RUN grep 'package/sdkjs/common/AllFonts.js' package.content
RUN grep 'package/sdkjs/common/Images/fonts_thumbnail@2x.png' package.content
RUN grep 'package/fonts/fonts/calibri.ttf' package.content
RUN grep 'package/package.json' package.content


FROM scratch AS build
COPY --from=package-build /app/onlyoffice-editor-*.tgz /
COPY --from=package-build /app/onlyoffice-editor-*.tgz.sha512 /


FROM scratch AS files
COPY --from=files-build /app/ /
