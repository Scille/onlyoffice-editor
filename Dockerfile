# Use this (large) base image in every build below, to reduce the overall docker cache size
FROM ubuntu:26.04 AS base

# Workaround for slow archive.ubuntu.com
RUN sed -i 's|archive.ubuntu.com|ftp.halifax.rwth-aachen.de|g' /etc/apt/sources.list.d/ubuntu.sources

RUN apt-get update && apt-get install -y openjdk-21-jdk npm wget zip brotli make bzip2
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
# Skip by default the built-in editor documentation (apps/*/main/resources/help).
# It is ~500MB (mostly per-language PNG screenshots) and is only used by the editors'
# "?" help dialog (that can itself be disabled at runtime by setting `help: false`
# in the option when opening the editor).
ARG INCLUDE_HELP=false
COPY --from=sdkjs-build /app/sdkjs/deploy/web-apps /app/web-apps
COPY --from=sdkjs-build /app/sdkjs/deploy/sdkjs /app/sdkjs
COPY vendor /app/web-apps/vendor
COPY fonts/*.ttf /app/fonts/fonts/
COPY fonts/*.otf /app/fonts/fonts/
COPY dictionaries /app/dictionaries
COPY --from=onlyoffice-editor-build /app/dist/api.js /app/web-apps/apps/api/documents/api.js
RUN if [ "$INCLUDE_HELP" != "true" ]; then rm -rf /app/web-apps/apps/*/main/resources/help; fi
WORKDIR /app


FROM files-build AS zip-build
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
RUN zip -r onlyoffice-editor.zip .
RUN sha512sum onlyoffice-editor.zip > onlyoffice-editor.zip.sha512



FROM zip-build AS zip-test
RUN unzip -l onlyoffice-editor.zip > zip.content
RUN grep ' sdkjs/common/AllFonts.js' zip.content
RUN grep ' sdkjs/common/Images/fonts_thumbnail@2x.png' zip.content
RUN grep ' fonts/fonts/calibri.ttf' zip.content


FROM scratch AS build
COPY --from=zip-build /app/onlyoffice-editor.zip /
COPY --from=zip-build /app/onlyoffice-editor.zip.sha512 /


FROM scratch AS files
COPY --from=files-build /app/ /
