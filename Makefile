.PHONY: help deps build build-launcher install build-linux build-macos build-macos-unsigned clean install-linux install-macos install-plugin uninstall uninstall-linux uninstall-macos print-dirs pretty ensure-bun generate-config generate-example-config dev-start dev-start-macos dev-watch dev-watch-macos dev-toggle dev-stop

APP_NAME := subminer
THEME_SOURCE := assets/themes/subminer.rasi
LAUNCHER_OUT := dist/launcher/$(APP_NAME)
THEME_FILE := subminer.rasi
PLUGIN_CONF := plugin/subminer.conf

# Default install prefix for the wrapper script.
PREFIX ?= $(HOME)/.local
BINDIR ?= $(PREFIX)/bin

# Linux data dir defaults to XDG_DATA_HOME/SubMiner.
XDG_DATA_HOME ?= $(HOME)/.local/share
LINUX_DATA_DIR ?= $(XDG_DATA_HOME)/SubMiner

# macOS data dir uses the standard Application Support location.
# Note: contains spaces; recipes must quote it.
MACOS_DATA_DIR ?= $(HOME)/Library/Application Support/SubMiner
MACOS_APP_DIR ?= $(HOME)/Applications
MACOS_APP_DEST ?= $(MACOS_APP_DIR)/SubMiner.app

# mpv plugin install directories.
MPV_CONFIG_DIR ?= $(HOME)/.config/mpv
MPV_SCRIPTS_DIR ?= $(MPV_CONFIG_DIR)/scripts
MPV_SCRIPT_OPTS_DIR ?= $(MPV_CONFIG_DIR)/script-opts

# If building from source, the AppImage will typically land in release/.
APPIMAGE_SRC := $(firstword $(wildcard release/SubMiner-*.AppImage))
MACOS_APP_SRC := $(firstword $(wildcard release/*.app release/*/*.app))
MACOS_ZIP_SRC := $(firstword $(wildcard release/SubMiner-*.zip))

UNAME_S := $(shell uname -s 2>/dev/null || echo Unknown)
ifeq ($(OS),Windows_NT)
PLATFORM := windows
else ifeq ($(UNAME_S),Linux)
PLATFORM := linux
else ifeq ($(UNAME_S),Darwin)
PLATFORM := macos
else
PLATFORM := unknown
endif

help:
	@printf '%s\n' \
		"Targets:" \
		"  build            Build platform package for detected OS ($(PLATFORM))" \
		"  install          Install platform artifacts for detected OS ($(PLATFORM))" \
		"  build-linux      Build Linux AppImage" \
		"  build-macos      Build macOS DMG/ZIP (signed if configured)" \
		"  build-macos-unsigned Build macOS DMG/ZIP without signing/notarization" \
		"  clean            Remove build artifacts (dist/, release/, AppImage, binary)" \
		"  dev-start        Build and launch local Electron app" \
		"  dev-start-macos  Build and launch local Electron app with macOS tracker backend" \
		"  dev-watch        Start fast watch loop (tsc + renderer + Electron dev app)" \
		"  dev-watch-macos  Start watch loop with forced macOS tracker backend" \
		"  dev-toggle       Toggle overlay in a running local Electron app" \
		"  dev-stop         Stop a running local Electron app" \
		"  install-linux    Install Linux wrapper/theme/app artifacts" \
		"  install-macos    Install macOS wrapper/theme/app artifacts" \
		"  install-plugin   Install mpv Lua plugin and plugin config" \
		"  generate-config  Generate ~/.config/SubMiner/config.jsonc from centralized defaults" \
		"" \
		"Other targets:" \
		"  deps             Install JS dependencies (root + texthooker-ui)" \
		"  uninstall-linux  Remove Linux install artifacts" \
		"  uninstall-macos  Remove macOS install artifacts" \
		"  print-dirs       Show resolved install locations" \
		"" \
		"Variables:" \
		"  PREFIX=...        Override wrapper install prefix (default: $$HOME/.local)" \
		"  BINDIR=...        Override wrapper install bin dir" \
		"  XDG_DATA_HOME=... Override Linux data dir base (default: $$HOME/.local/share)" \
		"  LINUX_DATA_DIR=... Override Linux app data dir" \
		"  MACOS_DATA_DIR=... Override macOS app data dir" \
		"  MACOS_APP_DIR=...  Override macOS app install dir (default: $$HOME/Applications)" \
		"  MPV_CONFIG_DIR=... Override mpv config dir (default: $$HOME/.config/mpv)"

print-dirs:
	@printf '%s\n' \
		"PLATFORM=$(PLATFORM)" \
		"UNAME_S=$(UNAME_S)" \
		"BINDIR=$(BINDIR)" \
		"LINUX_DATA_DIR=$(LINUX_DATA_DIR)" \
		"MACOS_DATA_DIR=$(MACOS_DATA_DIR)" \
		"MACOS_APP_DIR=$(MACOS_APP_DIR)" \
		"MACOS_APP_DEST=$(MACOS_APP_DEST)" \
		"APPIMAGE_SRC=$(APPIMAGE_SRC)" \
		"MACOS_APP_SRC=$(MACOS_APP_SRC)" \
		"MACOS_ZIP_SRC=$(MACOS_ZIP_SRC)"

deps:
	@$(MAKE) --no-print-directory ensure-bun
	@bun install
	@cd vendor/texthooker-ui && bun install --frozen-lockfile

ensure-bun:
	@command -v bun >/dev/null 2>&1 || { printf '%s\n' "[ERROR] bun not found"; exit 1; }

pretty: ensure-bun
	@bun run format:src

build:
	@printf '%s\n' "[INFO] Detected platform: $(PLATFORM)"
	@case "$(PLATFORM)" in \
		linux) $(MAKE) --no-print-directory build-linux ;; \
		macos) $(MAKE) --no-print-directory build-macos ;; \
		*) printf '%s\n' "[ERROR] Unsupported OS for this Makefile target: $(PLATFORM)"; exit 1 ;; \
	esac

install:
	@printf '%s\n' "[INFO] Detected platform: $(PLATFORM)"
	@case "$(PLATFORM)" in \
		linux) $(MAKE) --no-print-directory install-linux ;; \
		macos) $(MAKE) --no-print-directory install-macos ;; \
		*) printf '%s\n' "[ERROR] Unsupported OS for this Makefile target: $(PLATFORM)"; exit 1 ;; \
	esac

build-linux: deps
	@printf '%s\n' "[INFO] Building Linux package (AppImage)"
	@cd vendor/texthooker-ui && bun run build
	@bun run build:appimage

build-macos: deps
	@printf '%s\n' "[INFO] Building macOS package (DMG + ZIP)"
	@cd vendor/texthooker-ui && bun run build
	@bun run build:mac

build-macos-unsigned: deps
	@printf '%s\n' "[INFO] Building macOS package (DMG + ZIP, unsigned)"
	@cd vendor/texthooker-ui && bun run build
	@bun run build:mac:unsigned

build-launcher:
	@printf '%s\n' "[INFO] Bundling launcher script"
	@install -d "$(dir $(LAUNCHER_OUT))"
	@bun build ./launcher/main.ts --target=bun --packages=bundle --outfile="$(LAUNCHER_OUT)"
	@if ! head -1 "$(LAUNCHER_OUT)" | grep -q '^#!/usr/bin/env bun'; then \
		{ printf '#!/usr/bin/env bun\n'; cat "$(LAUNCHER_OUT)"; } > "$(LAUNCHER_OUT).tmp" && mv "$(LAUNCHER_OUT).tmp" "$(LAUNCHER_OUT)"; \
	fi
	@chmod +x "$(LAUNCHER_OUT)"
	@printf '%s\n' "[INFO] Launcher artifact: $(LAUNCHER_OUT)"

clean:
	@printf '%s\n' "[INFO] Removing build artifacts"
	@rm -rf dist release
	@rm -f "$(BINDIR)/subminer" "$(BINDIR)/SubMiner.AppImage"

generate-config: ensure-bun
	@bun run build
	@bun run electron . --generate-config

generate-example-config: ensure-bun
	@bun run build
	@bun run generate:config-example

dev-start: ensure-bun
	@bun run build
	@bun run electron . --start

dev-start-macos: ensure-bun
	@bun run build
	@bun run electron . --start --backend macos

dev-watch: ensure-bun
	@bash scripts/dev-watch.sh

dev-watch-macos: ensure-bun
	@bash scripts/dev-watch.sh --start --dev --backend macos

dev-toggle: ensure-bun
	@bun run electron . --toggle

dev-stop: ensure-bun
	@bun run electron . --stop


install-linux: build-launcher
	@printf '%s\n' "[INFO] Installing Linux wrapper/theme artifacts"
	@install -d "$(BINDIR)"
	@install -m 0755 "$(LAUNCHER_OUT)" "$(BINDIR)/$(APP_NAME)"
	@install -d "$(LINUX_DATA_DIR)/themes"
	@install -m 0644 "./$(THEME_SOURCE)" "$(LINUX_DATA_DIR)/themes/$(THEME_FILE)"
	@if [ -n "$(APPIMAGE_SRC)" ]; then \
		install -m 0755 "$(APPIMAGE_SRC)" "$(BINDIR)/SubMiner.AppImage"; \
	else \
		printf '%s\n' "[WARN] No release/SubMiner-*.AppImage found; skipping AppImage install"; \
		printf '%s\n' "       Build one with: make build"; \
	fi
	@printf '%s\n' "Installed to:" "  $(BINDIR)/subminer" "  $(LINUX_DATA_DIR)/themes/$(THEME_FILE)"

install-macos: build-launcher
	@printf '%s\n' "[INFO] Installing macOS wrapper/theme/app artifacts"
	@install -d "$(BINDIR)"
	@install -m 0755 "$(LAUNCHER_OUT)" "$(BINDIR)/$(APP_NAME)"
	@install -d "$(MACOS_DATA_DIR)/themes"
	@install -m 0644 "./$(THEME_SOURCE)" "$(MACOS_DATA_DIR)/themes/$(THEME_FILE)"
	@install -d "$(MACOS_APP_DIR)"
	@if [ -n "$(MACOS_APP_SRC)" ]; then \
		rm -rf "$(MACOS_APP_DEST)"; \
		cp -R "$(MACOS_APP_SRC)" "$(MACOS_APP_DEST)"; \
		printf '%s\n' "[INFO] Installed app bundle from $(MACOS_APP_SRC)"; \
	elif [ -n "$(MACOS_ZIP_SRC)" ]; then \
		rm -rf "$(MACOS_APP_DEST)"; \
		ditto -x -k "$(MACOS_ZIP_SRC)" "$(MACOS_APP_DIR)"; \
		printf '%s\n' "[INFO] Installed app bundle from $(MACOS_ZIP_SRC)"; \
	else \
		printf '%s\n' "[WARN] No macOS app bundle or zip found in release/; skipping app install"; \
		printf '%s\n' "       Build one with: make build"; \
	fi
	@printf '%s\n' "Installed to:" "  $(BINDIR)/subminer" "  $(MACOS_DATA_DIR)/themes/$(THEME_FILE)" "  $(MACOS_APP_DEST)"

install-plugin:
	@printf '%s\n' "[INFO] Installing mpv plugin artifacts"
	@install -d "$(MPV_SCRIPTS_DIR)"
	@rm -f "$(MPV_SCRIPTS_DIR)/subminer.lua"
	@install -d "$(MPV_SCRIPTS_DIR)/subminer"
	@install -d "$(MPV_SCRIPT_OPTS_DIR)"
	@cp -R ./plugin/subminer/. "$(MPV_SCRIPTS_DIR)/subminer/"
	@install -m 0644 "./$(PLUGIN_CONF)" "$(MPV_SCRIPT_OPTS_DIR)/subminer.conf"
	@printf '%s\n' "Installed to:" "  $(MPV_SCRIPTS_DIR)/subminer/main.lua" "  $(MPV_SCRIPTS_DIR)/subminer/" "  $(MPV_SCRIPT_OPTS_DIR)/subminer.conf"

# Uninstall behavior kept unchanged by default.
uninstall: uninstall-linux

uninstall-linux:
	@rm -f "$(BINDIR)/subminer" "$(BINDIR)/SubMiner.AppImage"
	@rm -f "$(LINUX_DATA_DIR)/themes/$(THEME_FILE)"
	@printf '%s\n' "Removed:" "  $(BINDIR)/subminer" "  $(BINDIR)/SubMiner.AppImage" "  $(LINUX_DATA_DIR)/themes/$(THEME_FILE)"

uninstall-macos:
	@rm -f "$(BINDIR)/subminer"
	@rm -f "$(MACOS_DATA_DIR)/themes/$(THEME_FILE)"
	@rm -rf "$(MACOS_APP_DEST)"
	@printf '%s\n' "Removed:" "  $(BINDIR)/subminer" "  $(MACOS_DATA_DIR)/themes/$(THEME_FILE)" "  $(MACOS_APP_DEST)"
