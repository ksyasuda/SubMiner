.PHONY: help deps build install build-linux build-macos build-macos-unsigned install-linux install-macos install-plugin uninstall uninstall-linux uninstall-macos print-dirs pretty ensure-pnpm generate-config generate-example-config docs-dev docs docs-preview dev-start dev-start-macos dev-toggle dev-stop

APP_NAME := subminer
THEME_FILE := subminer.rasi
PLUGIN_LUA := plugin/subminer.lua
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
		"  dev-start        Build and launch local Electron app" \
		"  dev-start-macos  Build and launch local Electron app with macOS tracker backend" \
		"  dev-toggle       Toggle overlay in a running local Electron app" \
		"  dev-stop         Stop a running local Electron app" \
		"  docs-dev         Run VitePress docs dev server" \
		"  docs       Build VitePress static docs" \
		"  docs-preview     Preview built VitePress docs" \
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
	@$(MAKE) --no-print-directory ensure-pnpm
	@pnpm install
	@pnpm -C vendor/texthooker-ui install

ensure-pnpm:
	@command -v pnpm >/dev/null 2>&1 || { printf '%s\n' "[ERROR] pnpm not found"; exit 1; }

pretty:
	@pnpm exec prettier --write 'src/**/*.ts'

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
	@pnpm -C vendor/texthooker-ui build
	@pnpm run build:appimage

build-macos: deps
	@printf '%s\n' "[INFO] Building macOS package (DMG + ZIP)"
	@pnpm -C vendor/texthooker-ui build
	@pnpm run build:mac

build-macos-unsigned: deps
	@printf '%s\n' "[INFO] Building macOS package (DMG + ZIP, unsigned)"
	@pnpm -C vendor/texthooker-ui build
	@pnpm run build:mac:unsigned

generate-config: ensure-pnpm
	@pnpm run build
	@pnpm exec electron . --generate-config

generate-example-config: ensure-pnpm
	@pnpm run build
	@pnpm run generate:config-example

docs-dev: ensure-pnpm
	@pnpm run docs:dev

docs: ensure-pnpm
	@pnpm run docs:build

docs-preview: ensure-pnpm
	@pnpm run docs:preview

dev-start: ensure-pnpm
	@pnpm run build
	@pnpm exec electron . --start

dev-start-macos: ensure-pnpm
	@pnpm run build
	@pnpm exec electron . --start --backend macos

dev-toggle: ensure-pnpm
	@pnpm exec electron . --toggle

dev-stop: ensure-pnpm
	@pnpm exec electron . --stop


install-linux:
	@printf '%s\n' "[INFO] Installing Linux wrapper/theme artifacts"
	@install -d "$(BINDIR)"
	@install -m 0755 "./$(APP_NAME)" "$(BINDIR)/$(APP_NAME)"
	@install -d "$(LINUX_DATA_DIR)/themes"
	@install -m 0644 "./$(THEME_FILE)" "$(LINUX_DATA_DIR)/themes/$(THEME_FILE)"
	@if [ -n "$(APPIMAGE_SRC)" ]; then \
		install -m 0755 "$(APPIMAGE_SRC)" "$(BINDIR)/SubMiner.AppImage"; \
	else \
		printf '%s\n' "[WARN] No release/SubMiner-*.AppImage found; skipping AppImage install"; \
		printf '%s\n' "       Build one with: make build"; \
	fi
	@printf '%s\n' "Installed to:" "  $(BINDIR)/subminer" "  $(LINUX_DATA_DIR)/themes/$(THEME_FILE)"

install-macos:
	@printf '%s\n' "[INFO] Installing macOS wrapper/theme/app artifacts"
	@install -d "$(BINDIR)"
	@install -m 0755 "./$(APP_NAME)" "$(BINDIR)/$(APP_NAME)"
	@install -d "$(MACOS_DATA_DIR)/themes"
	@install -m 0644 "./$(THEME_FILE)" "$(MACOS_DATA_DIR)/themes/$(THEME_FILE)"
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
	@install -d "$(MPV_SCRIPT_OPTS_DIR)"
	@install -m 0644 "./$(PLUGIN_LUA)" "$(MPV_SCRIPTS_DIR)/subminer.lua"
	@install -m 0644 "./$(PLUGIN_CONF)" "$(MPV_SCRIPT_OPTS_DIR)/subminer.conf"
	@printf '%s\n' "Installed to:" "  $(MPV_SCRIPTS_DIR)/subminer.lua" "  $(MPV_SCRIPT_OPTS_DIR)/subminer.conf"

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
