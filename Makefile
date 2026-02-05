SHELL := /bin/bash
.SHELLFLAGS := -euo pipefail -c

APP_NAME := mpv-yomitan.app
ZIP ?= $(shell ls -t release/mpv-yomitan-*-mac.zip 2>/dev/null | head -n 1)
INSTALL_DIR ?= /Applications

.PHONY: install-mac install-mac-user

build-mac:
	pnpm build:mac
install-mac:
	@ZIP_PATH="$(ZIP)"; \
	if [[ -z "$$ZIP_PATH" ]]; then \
		echo "ZIP not found. Set ZIP=path/to/mpv-yomitan-*-mac.zip"; \
		exit 1; \
	fi; \
	if [[ ! -f "$$ZIP_PATH" ]]; then \
		echo "ZIP does not exist: $$ZIP_PATH"; \
		exit 1; \
	fi; \
	TMP_DIR="$$(mktemp -d /tmp/mpv-yomitan-install.XXXXXX)"; \
	unzip -q "$$ZIP_PATH" -d "$$TMP_DIR"; \
	if [[ ! -d "$$TMP_DIR/$(APP_NAME)" ]]; then \
		echo "Expected $$TMP_DIR/$(APP_NAME) after unzip"; \
		ls -la "$$TMP_DIR"; \
		rm -rf "$$TMP_DIR"; \
		exit 1; \
	fi; \
	if [[ -d "$(INSTALL_DIR)/$(APP_NAME)" ]]; then \
		rm -rf "$(INSTALL_DIR)/$(APP_NAME)"; \
	fi; \
	mv "$$TMP_DIR/$(APP_NAME)" "$(INSTALL_DIR)/$(APP_NAME)"; \
	rmdir "$$TMP_DIR" 2>/dev/null || true; \
	echo "Installed $(APP_NAME) to $(INSTALL_DIR)"

install-mac-user:
	@$(MAKE) install-mac INSTALL_DIR="$(HOME)/Applications"
