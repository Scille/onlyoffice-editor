tl;dr:
```shell
# Generates `output/onlyoffice-editor.zip`
podman build --target build -o output .
# Generates `output/onlyoffice-editor.zip`, but also includes:
# - The .br files, i.e. brotli-compressed assets (~100MB)
# - The OnlyOffice help documentation (~500MB)
# (Those add up a lot of space and are most likely not needed).
podman build --target build -o output . \
  --build-arg INCLUDE_HELP=true \
  --build-arg PRECOMPRESSED_ASSETS=true \
```

---
Original README
---

# TODO

# Build onlyoffice-editor.zip

Run

```
make build
```

and find the result here: `output/onlyoffice-editor.zip`

# Running all tests

```
make test
```

# Diff changes we made to sdkjs and web-apps

## sdkjs

```sh
git fetch --depth=1 https://github.com/ONLYOFFICE/sdkjs.git v9.3.0.140
git diff FETCH_HEAD HEAD:sdkjs
```

## web-apps

```sh
git fetch --depth=1 https://github.com/ONLYOFFICE/web-apps.git v9.3.0.140
git diff FETCH_HEAD HEAD:web-apps
```

## diff UI

Instead of calling `git diff ...`, you can use `git difftool --dir-diff ...` with the same parameters, to see the diff in a UI.

# Pull OnlyOffice upstream changes
```shell
git subtree pull --prefix sdkjs https://github.com/ONLYOFFICE/sdkjs.git <TAG> --squash
git subtree pull --prefix web-apps https://github.com/ONLYOFFICE/web-apps.git <TAG> --squash
```
