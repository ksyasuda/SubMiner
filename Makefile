.PHONY: help deps build build-linux build-macos build-appimage install install-linux install-macos uninstall uninstall-linux uninstall-macos print-dirs pretty

APP_NAME := subminer
THEME_FILE := subminer.rasi

# Default install prefix for the wrapper script.
PREFIX ?= $(HOME)/.local
BINDIR ?= $(PREFIX)/bin

# Linux data dir defaults to XDG_DATA_HOME/SubMiner.
XDG_DATA_HOME ?= $(HOME)/.local/share
LINUX_DATA_DIR ?= $(XDG_DATA_HOME)/SubMiner

# macOS data dir uses the standard Application Support location.
# Note: contains spaces; recipes must quote it.
MACOS_DATA_DIR ?= $(HOME)/Library/Application Support/SubMiner

# If building from source, the AppImage will typically land in dist/.
APPIMAGE_SRC := $(firstword $(wildcard release/SubMiner-*.AppImage))

help:
	@printf '%s\n' \
		"Targets:" \
		"  deps             Install JS dependencies (root + texthooker-ui)" \
		"  build            Build app JS + texthooker-ui" \
		"  build-linux       Build Linux AppImage" \
		"  build-macos       Build app JS + texthooker-ui (no packaging)" \
		"  install-linux     Install wrapper + theme (and AppImage if present)" \
		"  install-macos     Install wrapper + theme" \
		"  uninstall-linux   Remove installed wrapper/theme" \
		"  uninstall-macos   Remove installed wrapper/theme" \
		"  print-dirs        Show resolved install locations" \
		"" \
		"Variables:" \
		"  PREFIX=...        Override wrapper install prefix (default: $$HOME/.local)" \
		"  BINDIR=...        Override wrapper install bin dir" \
		"  XDG_DATA_HOME=... Override Linux data dir base (default: $$HOME/.local/share)" \
		"  LINUX_DATA_DIR=... Override Linux app data dir" \
		"  MACOS_DATA_DIR=... Override macOS app data dir"

print-dirs:
	@printf '%s\n' \
		"BINDIR=$(BINDIR)" \
		"LINUX_DATA_DIR=$(LINUX_DATA_DIR)" \
		"MACOS_DATA_DIR=$(MACOS_DATA_DIR)" \
		"APPIMAGE_SRC=$(APPIMAGE_SRC)"

deps:
	@command -v pnpm >/dev/null 2>&1 || { printf '%s\n' "[ERROR] pnpm not found"; exit 1; }
	@pnpm install
	@pnpm -C vendor/texthooker-ui install

pretty:
	@pnpm exec prettier --write 'src/**/*.ts'

build:
	@command -v pnpm >/dev/null 2>&1 || { printf '%s\n' "[ERROR] pnpm not found"; exit 1; }
	@pnpm install
	@pnpm -C vendor/texthooker-ui install
	@pnpm -C vendor/texthooker-ui build
	@pnpm run build

build-macos: build

build-appimage:
	@command -v pnpm >/dev/null 2>&1 || { printf '%s\n' "[ERROR] pnpm not found"; exit 1; }
	@pnpm install
	@pnpm -C vendor/texthooker-ui install
	@pnpm -C vendor/texthooker-ui build
	@pnpm run build:appimage

build-linux: build-appimage

install: install-linux

install-linux:
	@install -d "$(BINDIR)"
	@install -m 0755 "./$(APP_NAME)" "$(BINDIR)/$(APP_NAME)"
	@install -d "$(LINUX_DATA_DIR)/themes"
	@install -m 0644 "./$(THEME_FILE)" "$(LINUX_DATA_DIR)/themes/$(THEME_FILE)"
	@if [ -n "$(APPIMAGE_SRC)" ]; then \
		install -m 0755 "$(APPIMAGE_SRC)" "$(BINDIR)/SubMiner.AppImage"; \
	else \
		printf '%s\n' "[WARN] No dist/SubMiner-*.AppImage found; skipping AppImage install"; \
		printf '%s\n' "       Build one with: pnpm run build:appimage"; \
	fi
	@printf '%s\n' "Installed to:" "  $(BINDIR)/subminer" "  $(LINUX_DATA_DIR)/themes/$(THEME_FILE)"

install-macos:
	@install -d "$(BINDIR)"
	@install -m 0755 "./$(APP_NAME)" "$(BINDIR)/$(APP_NAME)"
	@install -d "$(MACOS_DATA_DIR)/themes"
	@install -m 0644 "./$(THEME_FILE)" "$(MACOS_DATA_DIR)/themes/$(THEME_FILE)"
	@printf '%s\n' "Installed to:" "  $(BINDIR)/subminer" "  $(MACOS_DATA_DIR)/themes/$(THEME_FILE)"

uninstall: uninstall-linux

uninstall-linux:
	@rm -f "$(BINDIR)/subminer" "$(BINDIR)/SubMiner.AppImage"
	@rm -f "$(LINUX_DATA_DIR)/themes/$(THEME_FILE)"
	@printf '%s\n' "Removed:" "  $(BINDIR)/subminer" "  $(BINDIR)/SubMiner.AppImage" "  $(LINUX_DATA_DIR)/themes/$(THEME_FILE)"

uninstall-macos:
	@rm -f "$(BINDIR)/subminer"
	@rm -f "$(MACOS_DATA_DIR)/themes/$(THEME_FILE)"
	@printf '%s\n' "Removed:" "  $(BINDIR)/subminer" "  $(MACOS_DATA_DIR)/themes/$(THEME_FILE)"
