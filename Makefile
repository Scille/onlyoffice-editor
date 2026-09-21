.PHONY: build
build:
	docker build --build-arg release_version="$(release_version)" --target build -o output .

.PHONY: files
files:
	docker build --build-arg release_version="$(release_version)" --target files -o output-files .

.PHONY: onlyoffice-editor-test
onlyoffice-editor-test:
	docker build --target onlyoffice-editor-test .

.PHONY: package-test
package-test:
	docker build --build-arg release_version=0.0.0-test --target package-test .

.PHONY: test
test: onlyoffice-editor-test package-test
