.PHONY: help submodules deps build build-launcher install build-linux build-macos build-macos-unsigned clean install-linux install-macos install-windows uninstall uninstall-linux uninstall-macos uninstall-windows print-dirs pretty lint ensure-bun generate-config generate-example-config dev-start dev-start-macos dev-watch dev-watch-macos dev-toggle dev-stop docs-test docs-build docs-build-versioned docs-dev

APP_NAME := subminer
THEME_SOURCE := assets/themes/subminer.rasi
THUMBNAILER_SOURCE := assets/thumbnailers/subminer-ffmpegthumbnailer.thumbnailer
LAUNCHER_OUT := dist/launcher/$(APP_NAME)
THEME_FILE := subminer.rasi
THUMBNAILER_FILE := subminer-ffmpegthumbnailer.thumbnailer

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

# If building from source, the AppImage will typically land in release/.
APPIMAGE_SRC = $(firstword $(wildcard release/SubMiner-*.AppImage))
MACOS_APP_SRC = $(firstword $(wildcard release/*.app release/*/*.app))
MACOS_ZIP_SRC = $(firstword $(wildcard release/SubMiner-*.zip))
PRERELEASE_NOTES := release/prerelease-notes.md

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

WINDOWS_APPDATA ?= $(if $(APPDATA),$(subst \,/,$(APPDATA)),$(HOME)/AppData/Roaming)

# mpv plugin install directories.
ifeq ($(PLATFORM),windows)
MPV_CONFIG_DIR ?= $(WINDOWS_APPDATA)/mpv
else
MPV_CONFIG_DIR ?= $(HOME)/.config/mpv
endif
MPV_SCRIPTS_DIR ?= $(MPV_CONFIG_DIR)/scripts
MPV_SCRIPT_OPTS_DIR ?= $(MPV_CONFIG_DIR)/script-opts

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
		"  docs-test        Run docs tests" \
		"  docs-build       Build the docs site" \
		"  docs-build-versioned Build production versioned docs site" \
		"  docs-dev         Start the docs dev server" \
		"  install-linux    Install Linux wrapper/theme/app artifacts" \
		"  install-macos    Install macOS wrapper/theme/app artifacts" \
		"  install-windows  Print Windows packaging/install guidance" \
		"  generate-config  Generate ~/.config/SubMiner/config.jsonc from centralized defaults" \
		"" \
		"Other targets:" \
		"  submodules       Initialize/update git submodules" \
		"  deps             Initialize submodules and install JS dependencies (root + stats + texthooker-ui)" \
		"  uninstall-linux  Remove Linux install artifacts" \
		"  uninstall-macos  Remove macOS install artifacts" \
		"  uninstall-windows Remove Windows mpv plugin artifacts" \
		"  print-dirs       Show resolved install locations" \
		"  lint             Lint stats (format check)" \
		"" \
		"Variables:" \
		"  PREFIX=...        Override wrapper install prefix (default: $$HOME/.local)" \
		"  BINDIR=...        Override wrapper install bin dir" \
		"  XDG_DATA_HOME=... Override Linux data dir base (default: $$HOME/.local/share)" \
		"  LINUX_DATA_DIR=... Override Linux app data dir" \
		"  MACOS_DATA_DIR=... Override macOS app data dir" \
		"  MACOS_APP_DIR=...  Override macOS app install dir (default: $$HOME/Applications)" \
		"  MPV_CONFIG_DIR=... Override mpv config dir (default: $$HOME/.config/mpv or %APPDATA%/mpv on Windows)"

print-dirs:
	@printf '%s\n' \
		"PLATFORM=$(PLATFORM)" \
		"UNAME_S=$(UNAME_S)" \
		"BINDIR=$(BINDIR)" \
		"LINUX_DATA_DIR=$(LINUX_DATA_DIR)" \
		"MACOS_DATA_DIR=$(MACOS_DATA_DIR)" \
		"MACOS_APP_DIR=$(MACOS_APP_DIR)" \
		"MACOS_APP_DEST=$(MACOS_APP_DEST)" \
		"WINDOWS_APPDATA=$(WINDOWS_APPDATA)" \
		"MPV_CONFIG_DIR=$(MPV_CONFIG_DIR)" \
		"MPV_SCRIPTS_DIR=$(MPV_SCRIPTS_DIR)" \
		"MPV_SCRIPT_OPTS_DIR=$(MPV_SCRIPT_OPTS_DIR)" \
		"APPIMAGE_SRC=$(APPIMAGE_SRC)" \
		"MACOS_APP_SRC=$(MACOS_APP_SRC)" \
		"MACOS_ZIP_SRC=$(MACOS_ZIP_SRC)"

submodules:
	@git submodule update --init --recursive

deps: submodules ensure-bun
	@bun install
	@cd stats && bun install --frozen-lockfile
	@cd vendor/texthooker-ui && bun install --frozen-lockfile

ensure-bun:
	@command -v bun >/dev/null 2>&1 || { printf '%s\n' "[ERROR] bun not found"; exit 1; }

pretty: ensure-bun
	@bun run format:src
	@bun run format:stats

lint: ensure-bun
	@bun run lint:stats

build:
	@printf '%s\n' "[INFO] Detected platform: $(PLATFORM)"
	@case "$(PLATFORM)" in \
		linux) $(MAKE) --no-print-directory build-linux ;; \
		macos) $(MAKE) --no-print-directory build-macos ;; \
		windows) printf '%s\n' "[INFO] Windows builds run via: bun run build:win" ;; \
		*) printf '%s\n' "[ERROR] Unsupported OS for this Makefile target: $(PLATFORM)"; exit 1 ;; \
	esac

install:
	@printf '%s\n' "[INFO] Detected platform: $(PLATFORM)"
	@case "$(PLATFORM)" in \
		linux) $(MAKE) --no-print-directory install-linux ;; \
		macos) $(MAKE) --no-print-directory install-macos ;; \
		windows) $(MAKE) --no-print-directory install-windows ;; \
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
	@if [ -f "$(PRERELEASE_NOTES)" ]; then \
		PRERELEASE_NOTES_BACKUP="$$(mktemp -t subminer-prerelease-notes.XXXXXX)" && \
		cp "$(PRERELEASE_NOTES)" "$$PRERELEASE_NOTES_BACKUP" && \
		rm -rf dist release && \
		install -d release && \
		mv "$$PRERELEASE_NOTES_BACKUP" "$(PRERELEASE_NOTES)"; \
	else \
		rm -rf dist release; \
	fi
	@rm -f "$(BINDIR)/subminer" "$(BINDIR)/SubMiner.AppImage"

generate-config: ensure-bun
	@bun run build
	@bun run electron . --generate-config

generate-example-config: ensure-bun
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

docs-test: ensure-bun
	@bun run docs:test

docs-build: ensure-bun
	@bun run docs:build

docs-build-versioned: ensure-bun
	@bun run docs:build:versioned

docs-dev: ensure-bun
	@bun run docs:dev


install-linux: build-launcher
	@printf '%s\n' "[INFO] Installing Linux wrapper/support artifacts"
	@install -d "$(BINDIR)"
	@install -m 0755 "$(LAUNCHER_OUT)" "$(BINDIR)/$(APP_NAME)"
	@install -d "$(LINUX_DATA_DIR)/themes"
	@install -m 0644 "./$(THEME_SOURCE)" "$(LINUX_DATA_DIR)/themes/$(THEME_FILE)"
	@install -d "$(LINUX_DATA_DIR)/thumbnailers"
	@install -m 0644 "./$(THUMBNAILER_SOURCE)" "$(LINUX_DATA_DIR)/thumbnailers/$(THUMBNAILER_FILE)"
	@install -d "$(LINUX_DATA_DIR)/plugin/subminer"
	@cp -R ./plugin/subminer/. "$(LINUX_DATA_DIR)/plugin/subminer/"
	@if [ -n "$(APPIMAGE_SRC)" ]; then \
		install -m 0755 "$(APPIMAGE_SRC)" "$(BINDIR)/SubMiner.AppImage"; \
	else \
		printf '%s\n' "[WARN] No release/SubMiner-*.AppImage found; skipping AppImage install"; \
		printf '%s\n' "       Build one with: make build"; \
	fi
	@printf '%s\n' "Installed to:" "  $(BINDIR)/subminer" "  $(LINUX_DATA_DIR)/themes/$(THEME_FILE)" "  $(LINUX_DATA_DIR)/thumbnailers/$(THUMBNAILER_FILE)"

install-macos: build-launcher
	@printf '%s\n' "[INFO] Installing macOS wrapper/theme/app artifacts"
	@install -d "$(BINDIR)"
	@install -m 0755 "$(LAUNCHER_OUT)" "$(BINDIR)/$(APP_NAME)"
	@install -d "$(MACOS_DATA_DIR)/themes"
	@install -m 0644 "./$(THEME_SOURCE)" "$(MACOS_DATA_DIR)/themes/$(THEME_FILE)"
	@install -d "$(MACOS_DATA_DIR)/plugin/subminer"
	@cp -R ./plugin/subminer/. "$(MACOS_DATA_DIR)/plugin/subminer/"
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

install-windows:
	@printf '%s\n' "[INFO] Windows builds run via: bun run build:win"
	@printf '%s\n' "[INFO] SubMiner-managed mpv launches inject the bundled runtime plugin; no global mpv plugin install is needed."

uninstall:
	@printf '%s\n' "[INFO] Detected platform: $(PLATFORM)"
	@case "$(PLATFORM)" in \
		linux) $(MAKE) --no-print-directory uninstall-linux ;; \
		macos) $(MAKE) --no-print-directory uninstall-macos ;; \
		windows) $(MAKE) --no-print-directory uninstall-windows ;; \
		*) printf '%s\n' "[ERROR] Unsupported OS for this Makefile target: $(PLATFORM)"; exit 1 ;; \
	esac

uninstall-linux:
	@rm -f "$(BINDIR)/subminer" "$(BINDIR)/SubMiner.AppImage"
	@rm -f "$(LINUX_DATA_DIR)/themes/$(THEME_FILE)"
	@rm -f "$(LINUX_DATA_DIR)/thumbnailers/$(THUMBNAILER_FILE)"
	@rm -rf "$(LINUX_DATA_DIR)/plugin/subminer"
	@printf '%s\n' "Removed:" "  $(BINDIR)/subminer" "  $(BINDIR)/SubMiner.AppImage" "  $(LINUX_DATA_DIR)/themes/$(THEME_FILE)" "  $(LINUX_DATA_DIR)/thumbnailers/$(THUMBNAILER_FILE)" "  $(LINUX_DATA_DIR)/plugin/subminer"

uninstall-macos:
	@rm -f "$(BINDIR)/subminer"
	@rm -f "$(MACOS_DATA_DIR)/themes/$(THEME_FILE)"
	@rm -rf "$(MACOS_DATA_DIR)/plugin/subminer"
	@rm -rf "$(MACOS_APP_DEST)"
	@printf '%s\n' "Removed:" "  $(BINDIR)/subminer" "  $(MACOS_DATA_DIR)/themes/$(THEME_FILE)" "  $(MACOS_DATA_DIR)/plugin/subminer" "  $(MACOS_APP_DEST)"

uninstall-windows:
	@rm -rf "$(MPV_SCRIPTS_DIR)/subminer"
	@rm -f "$(MPV_SCRIPTS_DIR)/subminer.lua" "$(MPV_SCRIPTS_DIR)/subminer-loader.lua" "$(MPV_SCRIPT_OPTS_DIR)/subminer.conf"
	@printf '%s\n' "Removed:" "  $(MPV_SCRIPTS_DIR)/subminer" "  $(MPV_SCRIPT_OPTS_DIR)/subminer.conf"
