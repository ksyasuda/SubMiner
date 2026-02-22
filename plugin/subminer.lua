local input = require("mp.input")
local mp = require("mp")
local msg = require("mp.msg")
local options = require("mp.options")
local utils = require("mp.utils")

local function is_windows()
	return package.config:sub(1, 1) == "\\"
end

local function is_macos()
	local platform = mp.get_property("platform") or ""
	if platform == "macos" or platform == "darwin" then
		return true
	end
	local ostype = os.getenv("OSTYPE") or ""
	return ostype:find("darwin") ~= nil
end

local function default_socket_path()
	if is_windows() then
		return "\\\\.\\pipe\\subminer-socket"
	end
	return "/tmp/subminer-socket"
end

local function is_linux()
	return not is_windows() and not is_macos()
end

local function normalize_binary_path_candidate(candidate)
	if type(candidate) ~= "string" then
		return nil
	end
	local trimmed = candidate:match("^%s*(.-)%s*$") or ""
	if trimmed == "" then
		return nil
	end
	if #trimmed >= 2 then
		local first = trimmed:sub(1, 1)
		local last = trimmed:sub(-1)
		if (first == '"' and last == '"') or (first == "'" and last == "'") then
			trimmed = trimmed:sub(2, -2)
		end
	end
	return trimmed ~= "" and trimmed or nil
end

local function binary_candidates_from_app_path(app_path)
	return {
		utils.join_path(app_path, "Contents", "MacOS", "SubMiner"),
		utils.join_path(app_path, "Contents", "MacOS", "subminer"),
	}
end

local opts = {
	binary_path = "",
	socket_path = default_socket_path(),
	texthooker_enabled = true,
	texthooker_port = 5174,
	backend = "auto",
	auto_start = true,
	auto_start_overlay = false, -- legacy alias, maps to auto_start_visible_overlay
	auto_start_visible_overlay = false,
	auto_start_invisible_overlay = "platform-default", -- platform-default | visible | hidden
	osd_messages = true,
	log_level = "info",
}

options.read_options(opts, "subminer")

local state = {
	overlay_running = false,
	texthooker_running = false,
	overlay_process = nil,
	binary_available = false,
	binary_path = nil,
	detected_backend = nil,
	invisible_overlay_visible = false,
	hover_highlight = {
		revision = -1,
		payload = nil,
		saved_sub_visibility = nil,
		saved_secondary_sub_visibility = nil,
		overlay_active = false,
		cached_ass = nil,
		clear_timer = nil,
		last_hover_update_ts = 0,
	},
}

local HOVER_MESSAGE_NAME = "subminer-hover-token"
local HOVER_MESSAGE_NAME_LEGACY = "yomipv-hover-token"
local DEFAULT_HOVER_BASE_COLOR = "FFFFFF"
local DEFAULT_HOVER_COLOR = "C6A0F6"

local LOG_LEVEL_PRIORITY = {
	debug = 10,
	info = 20,
	warn = 30,
	error = 40,
}

local function normalize_log_level(level)
	local normalized = (level or "info"):lower()
	if LOG_LEVEL_PRIORITY[normalized] then
		return normalized
	end
	return "info"
end

local function should_log(level)
	local current = normalize_log_level(opts.log_level)
	local target = normalize_log_level(level)
	return LOG_LEVEL_PRIORITY[target] >= LOG_LEVEL_PRIORITY[current]
end

local function subminer_log(level, scope, message)
	if not should_log(level) then
		return
	end
	local timestamp = os.date("%Y-%m-%d %H:%M:%S")
	local line = string.format("[subminer] - %s - %s - [%s] %s", timestamp, string.upper(level), scope, message)
	if level == "error" then
		msg.error(line)
	elseif level == "warn" then
		msg.warn(line)
	elseif level == "debug" then
		msg.debug(line)
	else
		msg.info(line)
	end
end

local function show_osd(message)
	if opts.osd_messages then
		mp.osd_message("SubMiner: " .. message, 3)
	end
end

local function to_hex_color(input)
	if type(input) ~= "string" then
		return nil
	end

	local hex = input:gsub("[%#%']", ""):gsub("^0x", "")
	if #hex ~= 6 and #hex ~= 3 then
		return nil
	end
	if #hex == 3 then
		return hex:sub(1, 1) .. hex:sub(1, 1) .. hex:sub(2, 2) .. hex:sub(2, 2) .. hex:sub(3, 3) .. hex:sub(3, 3)
	end
	return hex
end

local function fix_ass_color(input, fallback)
	local hex = to_hex_color(input)
	if not hex then
		return fallback or DEFAULT_HOVER_BASE_COLOR
	end
	local r, g, b = hex:sub(1, 2), hex:sub(3, 4), hex:sub(5, 6)
	return b .. g .. r
end

local function escape_ass_text(text)
	return (text or "")
		:gsub("\\", "\\\\")
		:gsub("{", "\\{")
		:gsub("}", "\\}")
		:gsub("\n", "\\N")
end

local function resolve_osd_dimensions()
	local width = mp.get_property_number("osd-width", 0) or 0
	local height = mp.get_property_number("osd-height", 0) or 0

	if width <= 0 or height <= 0 then
		local osd_dims = mp.get_property_native("osd-dimensions")
		if type(osd_dims) == "table" and type(osd_dims.w) == "number" and osd_dims.w > 0 then
			width = osd_dims.w
		end
		if type(osd_dims) == "table" and type(osd_dims.h) == "number" and osd_dims.h > 0 then
			height = osd_dims.h
		end
	end

	if width <= 0 then
		width = 1280
	end
	if height <= 0 then
		height = 720
	end

	return width, height
end

local function resolve_metrics()
	local sub_font_size = mp.get_property_number("sub-font-size", 36) or 36
	local sub_scale = mp.get_property_number("sub-scale", 1) or 1
	local sub_scale_by_window = mp.get_property_bool("sub-scale-by-window", true) == true
	local sub_pos = mp.get_property_number("sub-pos", 100) or 100
	local sub_margin_y = mp.get_property_number("sub-margin-y", 0) or 0
	local sub_font = mp.get_property("sub-font", "sans-serif") or "sans-serif"
	local sub_spacing = mp.get_property_number("sub-spacing", 0) or 0
	local sub_bold = mp.get_property_bool("sub-bold", false) == true
	local sub_italic = mp.get_property_bool("sub-italic", false) == true
	local sub_border_size = mp.get_property_number("sub-border-size", 2) or 2
	local sub_shadow_offset = mp.get_property_number("sub-shadow-offset", 0) or 0
	local osd_w, osd_h = resolve_osd_dimensions()
	local window_scale = 1
	if sub_scale_by_window and osd_h > 0 then
		window_scale = osd_h / 720
	end
	local effective_margin_y = sub_margin_y * window_scale

	return {
		font_size = sub_font_size * (sub_scale > 0 and sub_scale or 1) * window_scale,
		pos = sub_pos,
		margin_y = effective_margin_y,
		font = sub_font,
		spacing = sub_spacing,
		bold = sub_bold,
		italic = sub_italic,
		border = sub_border_size * window_scale,
		shadow = sub_shadow_offset * window_scale,
		base_color = fix_ass_color(mp.get_property("sub-color"), DEFAULT_HOVER_BASE_COLOR),
		hover_color = fix_ass_color(mp.get_property("sub-color"), DEFAULT_HOVER_COLOR),
	}
end

local function get_subtitle_ass_property()
	local ass_text = mp.get_property("sub-text/ass")
	if type(ass_text) == "string" and ass_text ~= "" then
		return ass_text
	end

	ass_text = mp.get_property("sub-text-ass")
	if type(ass_text) == "string" and ass_text ~= "" then
		return ass_text
	end

	return nil
end

local function plain_text_and_ass_map(text)
	local plain = {}
	local map = {}
	local plain_len = 0
	local i = 1
	local text_len = #text

	while i <= text_len do
		local ch = text:sub(i, i)
		if ch == "{" then
			local close = text:find("}", i + 1, true)
			if not close then
				break
			end
			i = close + 1
		elseif ch == "\\" then
			local esc = text:sub(i + 1, i + 1)
			if esc == "N" or esc == "n" then
				plain_len = plain_len + 1
				plain[plain_len] = "\n"
				map[plain_len] = i
				i = i + 2
			elseif esc == "h" then
				plain_len = plain_len + 1
				plain[plain_len] = " "
				map[plain_len] = i
				i = i + 2
			elseif esc == "{" then
				plain_len = plain_len + 1
				plain[plain_len] = "{"
				map[plain_len] = i
				i = i + 2
			elseif esc == "}" then
				plain_len = plain_len + 1
				plain[plain_len] = "}"
				map[plain_len] = i
				i = i + 2
			elseif esc == "\\" then
				plain_len = plain_len + 1
				plain[plain_len] = "\\"
				map[plain_len] = i
				i = i + 2
			else
				local seq_end = i + 1
				while seq_end <= text_len and text:sub(seq_end, seq_end):match("[%a]") do
					seq_end = seq_end + 1
				end
				if text:sub(seq_end, seq_end) == "(" then
					local close = text:find(")", seq_end, true)
					if close then
						i = close + 1
					else
						i = seq_end + 1
					end
				else
					i = seq_end + 1
				end
			end
		else
			plain_len = plain_len + 1
			plain[plain_len] = ch
			map[plain_len] = i
			i = i + 1
		end
	end

	return table.concat(plain), map
end

local function find_hover_span(payload, plain)
	local source_len = #plain
	local cursor = 1
	for _, token in ipairs(payload.tokens or {}) do
		if type(token) ~= "table" or type(token.text) ~= "string" or token.text == "" then
			goto continue
		end

		local token_text = token.text
		local start_pos = nil
		local end_pos = nil

		if type(token.startPos) == "number" and type(token.endPos) == "number" then
			if token.startPos >= 0 and token.endPos >= token.startPos then
				local candidate_start = token.startPos + 1
				local candidate_stop = token.endPos
				if
					candidate_start >= 1
					and candidate_stop <= source_len
					and candidate_stop >= candidate_start
					and plain:sub(candidate_start, candidate_stop) == token_text
				then
					start_pos = candidate_start
					end_pos = candidate_stop
				end
			end
		end

		if not start_pos or not end_pos then
			local fallback_start, fallback_stop = plain:find(token_text, cursor, true)
			if not fallback_start then
				fallback_start, fallback_stop = plain:find(token_text, 1, true)
			end
			start_pos, end_pos = fallback_start, fallback_stop
		end

		if start_pos and end_pos then
			if token.index == payload.hoveredTokenIndex then
				return start_pos, end_pos
			end
			cursor = end_pos + 1
		end

		::continue::
	end

	return nil
end

local function inject_hover_color_to_ass(raw_ass, plain_map, hover_start, hover_end, hover_color, base_color)
	if hover_start == nil or hover_end == nil then
		return raw_ass
	end

	local raw_open_idx = plain_map[hover_start] or 1
	local raw_close_idx = plain_map[hover_end + 1] or (#raw_ass + 1)
	if raw_open_idx < 1 then
		raw_open_idx = 1
	end
	if raw_close_idx < 1 then
		raw_close_idx = 1
	end
	if raw_open_idx > #raw_ass + 1 then
		raw_open_idx = #raw_ass + 1
	end
	if raw_close_idx > #raw_ass + 1 then
		raw_close_idx = #raw_ass + 1
	end

	local open_tag = string.format("{\\1c&H%s&}", hover_color)
	local close_tag = string.format("{\\1c&H%s&}", base_color)
	local changes = {
		{ idx = raw_open_idx, tag = open_tag },
		{ idx = raw_close_idx, tag = close_tag },
	}
	table.sort(changes, function(a, b)
		return a.idx < b.idx
	end)

	local output = {}
	local cursor = 1
	for _, change in ipairs(changes) do
		if change.idx > #raw_ass + 1 then
			change.idx = #raw_ass + 1
		end
		if change.idx < 1 then
			change.idx = 1
		end
		if change.idx > cursor then
			output[#output + 1] = raw_ass:sub(cursor, change.idx - 1)
		end
		output[#output + 1] = change.tag
		cursor = change.idx
	end
	if cursor <= #raw_ass then
		output[#output + 1] = raw_ass:sub(cursor)
	end

	return table.concat(output)
end

local function build_hover_subtitle_content(payload)
	local source_ass = get_subtitle_ass_property()
	if type(source_ass) == "string" and source_ass ~= "" then
		state.hover_highlight.cached_ass = source_ass
	else
		source_ass = state.hover_highlight.cached_ass
	end
	if type(source_ass) ~= "string" or source_ass == "" then
		return nil
	end

	local plain_source, plain_map = plain_text_and_ass_map(source_ass)
	if type(plain_source) ~= "string" or plain_source == "" then
		return nil
	end

	local hover_start, hover_end = find_hover_span(payload, plain_source)
	if not hover_start or not hover_end then
		return nil
	end

	local metrics = resolve_metrics()
	local hover_color = fix_ass_color(payload.colors and payload.colors.hover or nil, metrics.hover_color)
	local base_color = fix_ass_color(payload.colors and payload.colors.base or nil, metrics.base_color)
	return inject_hover_color_to_ass(source_ass, plain_map, hover_start, hover_end, hover_color, base_color)
end

local function clear_hover_overlay()
	if state.hover_highlight.clear_timer then
		state.hover_highlight.clear_timer:kill()
		state.hover_highlight.clear_timer = nil
	end
	if state.hover_highlight.overlay_active then
		if type(state.hover_highlight.saved_sub_visibility) == "string" then
			mp.set_property("sub-visibility", state.hover_highlight.saved_sub_visibility)
		else
			mp.set_property("sub-visibility", "yes")
		end
		if type(state.hover_highlight.saved_secondary_sub_visibility) == "string" then
			mp.set_property("secondary-sub-visibility", state.hover_highlight.saved_secondary_sub_visibility)
		end
		state.hover_highlight.saved_sub_visibility = nil
		state.hover_highlight.saved_secondary_sub_visibility = nil
		state.hover_highlight.overlay_active = false
	end
	mp.set_osd_ass(0, 0, "")
	state.hover_highlight.payload = nil
	state.hover_highlight.revision = -1
	state.hover_highlight.cached_ass = nil
	state.hover_highlight.last_hover_update_ts = 0
end

local function schedule_hover_clear(delay_seconds)
	if state.hover_highlight.clear_timer then
		state.hover_highlight.clear_timer:kill()
		state.hover_highlight.clear_timer = nil
	end
	state.hover_highlight.clear_timer = mp.add_timeout(delay_seconds or 0.08, function()
		state.hover_highlight.clear_timer = nil
		clear_hover_overlay()
	end)
end

local function render_hover_overlay(payload)
	if not payload or payload.hoveredTokenIndex == nil or payload.subtitle == nil then
		clear_hover_overlay()
		return
	end

	local ass = build_hover_subtitle_content(payload)
	if not ass then
		-- Transient parse/mapping miss; keep previous frame to avoid flicker.
		return
	end

	local osd_w, osd_h = resolve_osd_dimensions()
	local metrics = resolve_metrics()
	local osd_dims = mp.get_property_native("osd-dimensions")
	local ml = (type(osd_dims) == "table" and type(osd_dims.ml) == "number") and osd_dims.ml or 0
	local mr = (type(osd_dims) == "table" and type(osd_dims.mr) == "number") and osd_dims.mr or 0
	local mt = (type(osd_dims) == "table" and type(osd_dims.mt) == "number") and osd_dims.mt or 0
	local mb = (type(osd_dims) == "table" and type(osd_dims.mb) == "number") and osd_dims.mb or 0
	local usable_w = math.max(1, osd_w - ml - mr)
	local usable_h = math.max(1, osd_h - mt - mb)
	local anchor_x = math.floor(ml + usable_w / 2)
	local baseline_adjust = (metrics.border + metrics.shadow) * 5
	local anchor_y = math.floor(mt + (usable_h * metrics.pos / 100) - metrics.margin_y + baseline_adjust)
	local font_size = math.max(8, metrics.font_size)
	local anchor_tag = string.format(
		"{\\an2\\q2\\pos(%d,%d)\\fn%s\\fs%g\\b%d\\i%d\\fsp%g\\bord%g\\shad%g\\1c&H%s&}",
		anchor_x,
		anchor_y,
		escape_ass_text(metrics.font),
		font_size,
		metrics.bold and 1 or 0,
		metrics.italic and 1 or 0,
		metrics.spacing,
		metrics.border,
		metrics.shadow,
		metrics.base_color
	)
	if not state.hover_highlight.overlay_active then
		state.hover_highlight.saved_sub_visibility = mp.get_property("sub-visibility")
		state.hover_highlight.saved_secondary_sub_visibility = mp.get_property("secondary-sub-visibility")
		mp.set_property("sub-visibility", "no")
		mp.set_property("secondary-sub-visibility", "no")
		state.hover_highlight.overlay_active = true
	end
	mp.set_osd_ass(osd_w, osd_h, anchor_tag .. ass)
end

local function handle_hover_message(payload_json)
	local parsed, parse_error = utils.parse_json(payload_json)
	if not parsed then
		msg.warn("Invalid hover-highlight payload: " .. tostring(parse_error))
		clear_hover_overlay()
		return
	end

	if type(parsed.revision) ~= "number" then
		clear_hover_overlay()
		return
	end

	if parsed.revision < state.hover_highlight.revision then
		return
	end

	if type(parsed.hoveredTokenIndex) == "number" and type(parsed.tokens) == "table" then
		if state.hover_highlight.clear_timer then
			state.hover_highlight.clear_timer:kill()
			state.hover_highlight.clear_timer = nil
		end
		state.hover_highlight.revision = parsed.revision
		state.hover_highlight.payload = parsed
		state.hover_highlight.last_hover_update_ts = mp.get_time() or 0
		render_hover_overlay(state.hover_highlight.payload)
		return
	end

	local now = mp.get_time() or 0
	local elapsed_since_hover = now - (state.hover_highlight.last_hover_update_ts or 0)
	state.hover_highlight.revision = parsed.revision
	state.hover_highlight.payload = nil
	if state.hover_highlight.overlay_active then
		if elapsed_since_hover > 0.35 then
			-- Ignore stale null-hover updates while pointer is stationary.
			return
		end
		schedule_hover_clear(0.08)
	else
		clear_hover_overlay()
	end
end

local function detect_backend()
	if state.detected_backend then
		return state.detected_backend
	end

	local backend = nil

	if is_macos() then
		backend = "macos"
	elseif is_windows() then
		backend = nil
	elseif os.getenv("HYPRLAND_INSTANCE_SIGNATURE") then
		backend = "hyprland"
	elseif os.getenv("SWAYSOCK") then
		backend = "sway"
	elseif os.getenv("XDG_SESSION_TYPE") == "x11" or os.getenv("DISPLAY") then
		backend = "x11"
	else
		subminer_log("warn", "backend", "Could not detect window manager, falling back to x11")
		backend = "x11"
	end

	state.detected_backend = backend
	if backend then
		subminer_log("info", "backend", "Detected backend: " .. backend)
	else
		subminer_log("info", "backend", "No backend detected")
	end
	return backend
end

local function file_exists(path)
	local info = utils.file_info(path)
	if not info then return false end
	if info.is_dir ~= nil then
		return not info.is_dir
	end
	return true
end

local function resolve_binary_candidate(candidate)
	local normalized = normalize_binary_path_candidate(candidate)
	if not normalized then
		return nil
	end

	if file_exists(normalized) then
		return normalized
	end

	if not normalized:lower():find("%.app") then
		return nil
	end

	local app_root = normalized
	if not app_root:lower():match("%.app$") then
		app_root = normalized:match("(.+%.app)")
	end
	if not app_root then
		return nil
	end

	for _, path in ipairs(binary_candidates_from_app_path(app_root)) do
		if file_exists(path) then
			return path
		end
	end

	return nil
end

local function find_binary_override()
	local candidates = {
		resolve_binary_candidate(os.getenv("SUBMINER_APPIMAGE_PATH")),
		resolve_binary_candidate(os.getenv("SUBMINER_BINARY_PATH")),
	}

	for _, path in ipairs(candidates) do
		if path and path ~= "" then
			return path
		end
	end

	return nil
end

local function find_binary()
	local override = find_binary_override()
	if override then
		return override
	end

	local configured = resolve_binary_candidate(opts.binary_path)
	if configured then
		return configured
	end

	local search_paths = {
		"/Applications/SubMiner.app/Contents/MacOS/SubMiner",
		utils.join_path(os.getenv("HOME") or "", "Applications/SubMiner.app/Contents/MacOS/SubMiner"),
		"C:\\Program Files\\SubMiner\\SubMiner.exe",
		"C:\\Program Files (x86)\\SubMiner\\SubMiner.exe",
		"C:\\SubMiner\\SubMiner.exe",
		utils.join_path(os.getenv("HOME") or "", ".local/bin/SubMiner.AppImage"),
		"/opt/SubMiner/SubMiner.AppImage",
		"/usr/local/bin/SubMiner",
		"/usr/bin/SubMiner",
	}

	for _, path in ipairs(search_paths) do
		if file_exists(path) then
			subminer_log("info", "binary", "Found binary at: " .. path)
			return path
		end
	end

	return nil
end

local function ensure_binary_available()
	if state.binary_available and state.binary_path and file_exists(state.binary_path) then
		return true
	end

	local discovered = find_binary()
	if discovered then
		state.binary_path = discovered
		state.binary_available = true
		return true
	end

	state.binary_path = nil
	state.binary_available = false
	return false
end

local function resolve_backend(override_backend)
	local selected = override_backend
	if selected == nil or selected == "" then
		selected = opts.backend
	end
	if selected == "auto" then
		return detect_backend()
	end
	return selected
end

local function build_command_args(action, overrides)
	overrides = overrides or {}
	local args = { state.binary_path }

	table.insert(args, "--" .. action)
	local log_level = normalize_log_level(overrides.log_level or opts.log_level)
	if log_level ~= "info" then
		table.insert(args, "--log-level")
		table.insert(args, log_level)
	end

	local needs_start_context = action == "start"

	if needs_start_context then
		local backend = resolve_backend(overrides.backend)
		if backend and backend ~= "" then
			table.insert(args, "--backend")
			table.insert(args, backend)
		end

		local socket_path = overrides.socket_path or opts.socket_path
		table.insert(args, "--socket")
		table.insert(args, socket_path)
	end

	return args
end

local function run_control_command(action)
	local args = build_command_args(action)
	subminer_log("debug", "process", "Control command: " .. table.concat(args, " "))
	local result = mp.command_native({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})
	return result and result.status == 0
end

local function coerce_bool(value, fallback)
	if type(value) == "boolean" then
		return value
	end
	if type(value) == "string" then
		local normalized = value:lower()
		if normalized == "yes" or normalized == "true" or normalized == "1" or normalized == "on" then
			return true
		end
		if normalized == "no" or normalized == "false" or normalized == "0" or normalized == "off" then
			return false
		end
	end
	return fallback
end

local function parse_start_script_message_overrides(...)
	local overrides = {}
	for i = 1, select("#", ...) do
		local token = select(i, ...)
		if type(token) == "string" and token ~= "" then
			local key, value = token:match("^([%w_%-]+)=(.+)$")
			if key and value then
				local normalized_key = key:lower()
				if normalized_key == "backend" then
					local backend = value:lower()
					if backend == "auto" or backend == "hyprland" or backend == "sway" or backend == "x11" or backend == "macos" then
						overrides.backend = backend
					end
				elseif normalized_key == "socket" or normalized_key == "socket_path" then
					overrides.socket_path = value
				elseif normalized_key == "texthooker" or normalized_key == "texthooker_enabled" then
					local parsed = coerce_bool(value, nil)
					if parsed ~= nil then
						overrides.texthooker_enabled = parsed
					end
				elseif normalized_key == "log-level" or normalized_key == "log_level" then
					overrides.log_level = normalize_log_level(value)
				end
			end
		end
	end
	return overrides
end

local function resolve_visible_overlay_startup()
	local visible = coerce_bool(opts.auto_start_visible_overlay, false)
	-- Backward compatibility for old config key.
	if coerce_bool(opts.auto_start_overlay, false) then
		visible = true
	end
	return visible
end

local function resolve_invisible_overlay_startup()
	local raw = opts.auto_start_invisible_overlay
	if type(raw) == "boolean" then
		return raw
	end

	local mode = type(raw) == "string" and raw:lower() or "platform-default"
	if mode == "visible" or mode == "show" or mode == "yes" or mode == "true" or mode == "on" then
		return true
	end
	if mode == "hidden" or mode == "hide" or mode == "no" or mode == "false" or mode == "off" then
		return false
	end

	-- platform-default
	return not is_linux()
end

local function apply_startup_overlay_preferences()
	local should_show_visible = resolve_visible_overlay_startup()
	local should_show_invisible = resolve_invisible_overlay_startup()

	local visible_action = should_show_visible and "show-visible-overlay" or "hide-visible-overlay"
	if not run_control_command(visible_action) then
		subminer_log("warn", "process", "Failed to apply visible startup action: " .. visible_action)
	end

	local invisible_action = should_show_invisible and "show-invisible-overlay" or "hide-invisible-overlay"
	if not run_control_command(invisible_action) then
		subminer_log("warn", "process", "Failed to apply invisible startup action: " .. invisible_action)
	end

	state.invisible_overlay_visible = should_show_invisible
end

local function build_texthooker_args()
	local args = { state.binary_path, "--texthooker", "--port", tostring(opts.texthooker_port) }
	local log_level = normalize_log_level(opts.log_level)
	if log_level ~= "info" then
		table.insert(args, "--log-level")
		table.insert(args, log_level)
	end
	return args
end

local function ensure_texthooker_running(callback)
	if not opts.texthooker_enabled then
		callback()
		return
	end

	if state.texthooker_running then
		callback()
		return
	end

	local args = build_texthooker_args()
	subminer_log("info", "texthooker", "Starting texthooker process: " .. table.concat(args, " "))
	state.texthooker_running = true

	mp.command_native_async({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	}, function(success, result, error)
		if not success or (result and result.status ~= 0) then
			state.texthooker_running = false
			subminer_log(
				"warn",
				"texthooker",
				"Texthooker process exited unexpectedly: " .. (error or (result and result.stderr) or "unknown error")
			)
		end
	end)

	-- Give the process a moment to acquire the app lock before sending --start.
	mp.add_timeout(0.35, callback)
end

local function start_overlay(overrides)
	if not ensure_binary_available() then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	if state.overlay_running then
		subminer_log("info", "process", "Overlay already running")
		show_osd("Already running")
		return
	end

	overrides = overrides or {}
	local texthooker_enabled = overrides.texthooker_enabled
	if texthooker_enabled == nil then
		texthooker_enabled = opts.texthooker_enabled
	end

	local function launch_overlay()
		local args = build_command_args("start", overrides)
		subminer_log("info", "process", "Starting overlay: " .. table.concat(args, " "))

		show_osd("Starting...")
		state.overlay_running = true

		mp.command_native_async({
			name = "subprocess",
			args = args,
			playback_only = false,
			capture_stdout = true,
			capture_stderr = true,
		}, function(success, result, error)
			if not success or (result and result.status ~= 0) then
				state.overlay_running = false
				subminer_log(
					"error",
					"process",
					"Overlay start failed: " .. (error or (result and result.stderr) or "unknown error")
				)
				show_osd("Overlay start failed")
			end
		end)

		-- Apply explicit startup visibility for each overlay layer.
		mp.add_timeout(0.6, function()
			apply_startup_overlay_preferences()
		end)
	end

	if texthooker_enabled then
		ensure_texthooker_running(launch_overlay)
	else
		launch_overlay()
	end
end

local function start_overlay_from_script_message(...)
	local overrides = parse_start_script_message_overrides(...)
	start_overlay(overrides)
end

local function stop_overlay()
	if not state.binary_available then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	local args = build_command_args("stop")
	subminer_log("info", "process", "Stopping overlay: " .. table.concat(args, " "))

	local result = mp.command_native({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})

	state.overlay_running = false
	state.texthooker_running = false
	if result.status == 0 then
		subminer_log("info", "process", "Overlay stopped")
	else
		subminer_log("warn", "process", "Stop command returned non-zero status: " .. tostring(result.status))
	end
	show_osd("Stopped")
end

local function toggle_overlay()
	if not state.binary_available then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	local args = build_command_args("toggle")
	subminer_log("info", "process", "Toggling overlay: " .. table.concat(args, " "))

	local result = mp.command_native({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})

	if result and result.status ~= 0 then
		subminer_log("warn", "process", "Toggle command failed")
		show_osd("Toggle failed")
	end
end

local function toggle_invisible_overlay()
	if not state.binary_available then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	local args = build_command_args("toggle-invisible-overlay")
	subminer_log("info", "process", "Toggling invisible overlay: " .. table.concat(args, " "))

	local result = mp.command_native({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})

	if result and result.status ~= 0 then
		subminer_log("warn", "process", "Invisible toggle command failed")
		show_osd("Invisible toggle failed")
		return
	end

	state.invisible_overlay_visible = not state.invisible_overlay_visible
	show_osd("Invisible overlay: " .. (state.invisible_overlay_visible and "visible" or "hidden"))
end

local function show_invisible_overlay()
	if not state.binary_available then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	local args = build_command_args("show-invisible-overlay")
	subminer_log("info", "process", "Showing invisible overlay: " .. table.concat(args, " "))

	local result = mp.command_native({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})

	if result and result.status ~= 0 then
		subminer_log("warn", "process", "Show invisible command failed")
		show_osd("Show invisible failed")
		return
	end

	state.invisible_overlay_visible = true
	show_osd("Invisible overlay: visible")
end

local function hide_invisible_overlay()
	if not state.binary_available then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	local args = build_command_args("hide-invisible-overlay")
	subminer_log("info", "process", "Hiding invisible overlay: " .. table.concat(args, " "))

	local result = mp.command_native({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})

	if result and result.status ~= 0 then
		subminer_log("warn", "process", "Hide invisible command failed")
		show_osd("Hide invisible failed")
		return
	end

	state.invisible_overlay_visible = false
	show_osd("Invisible overlay: hidden")
end

local function open_options()
	if not state.binary_available then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end
	local args = build_command_args("settings")
	subminer_log("info", "process", "Opening options: " .. table.concat(args, " "))
	local result = mp.command_native({
		name = "subprocess",
		args = args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})
	if result.status == 0 then
		subminer_log("info", "process", "Options window opened")
		show_osd("Options opened")
	else
		subminer_log("warn", "process", "Failed to open options")
		show_osd("Failed to open options")
	end
end

local restart_overlay
local check_status

local function show_menu()
	if not state.binary_available then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	local items = {
		"Start overlay",
		"Stop overlay",
		"Toggle overlay",
		"Toggle invisible overlay",
		"Open options",
		"Restart overlay",
		"Check status",
	}

	local actions = {
		start_overlay,
		stop_overlay,
		toggle_overlay,
		toggle_invisible_overlay,
		open_options,
		restart_overlay,
		check_status,
	}

	input.select({
		prompt = "SubMiner: ",
		items = items,
		submit = function(index)
			if index and actions[index] then
				actions[index]()
			end
		end,
	})
end

restart_overlay = function()
	if not state.binary_available then
		subminer_log("error", "binary", "SubMiner binary not found")
		show_osd("Error: binary not found")
		return
	end

	subminer_log("info", "process", "Restarting overlay...")
	show_osd("Restarting...")

	local stop_args = build_command_args("stop")
	mp.command_native({
		name = "subprocess",
		args = stop_args,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})

	state.overlay_running = false
	state.texthooker_running = false

	ensure_texthooker_running(function()
		local start_args = build_command_args("start")
		subminer_log("info", "process", "Starting overlay: " .. table.concat(start_args, " "))

		state.overlay_running = true
		mp.command_native_async({
			name = "subprocess",
			args = start_args,
			playback_only = false,
			capture_stdout = true,
			capture_stderr = true,
		}, function(success, result, error)
			if not success or (result and result.status ~= 0) then
				state.overlay_running = false
				subminer_log(
					"error",
					"process",
					"Overlay start failed: " .. (error or (result and result.stderr) or "unknown error")
				)
				show_osd("Restart failed")
			else
				show_osd("Restarted successfully")
			end
		end)
	end)
end

check_status = function()
	if not state.binary_available then
		show_osd("Status: binary not found")
		return
	end

	local status = state.overlay_running and "running" or "stopped"
	show_osd("Status: overlay is " .. status)
	subminer_log("info", "process", "Status check: overlay is " .. status)
end

local function on_file_loaded()
	state.binary_path = find_binary()
	if state.binary_path then
		state.binary_available = true
		subminer_log("info", "lifecycle", "SubMiner ready (binary: " .. state.binary_path .. ")")
		local should_auto_start = coerce_bool(opts.auto_start, false)
		if should_auto_start then
			start_overlay()
		end
	else
		state.binary_available = false
		subminer_log("warn", "binary", "SubMiner binary not found - overlay features disabled")
		if opts.binary_path ~= "" then
			subminer_log("warn", "binary", "Configured path '" .. opts.binary_path .. "' does not exist")
		end
	end
end

local function on_shutdown()
	clear_hover_overlay()
	if (state.overlay_running or state.texthooker_running) and state.binary_available then
		subminer_log("info", "lifecycle", "mpv shutting down, stopping SubMiner process")
		show_osd("Shutting down...")
		stop_overlay()
	end
end

local function register_keybindings()
	mp.add_key_binding("y-s", "subminer-start", start_overlay)
	mp.add_key_binding("y-S", "subminer-stop", stop_overlay)
	mp.add_key_binding("y-t", "subminer-toggle", toggle_overlay)
	mp.add_key_binding("y-i", "subminer-toggle-invisible", toggle_invisible_overlay)
	mp.add_key_binding("y-I", "subminer-show-invisible", show_invisible_overlay)
	mp.add_key_binding("y-u", "subminer-hide-invisible", hide_invisible_overlay)
	mp.add_key_binding("y-y", "subminer-menu", show_menu)
	mp.add_key_binding("y-o", "subminer-options", open_options)
	mp.add_key_binding("y-r", "subminer-restart", restart_overlay)
	mp.add_key_binding("y-c", "subminer-status", check_status)
end

local function register_script_messages()
	mp.register_script_message("subminer-start", start_overlay_from_script_message)
	mp.register_script_message("subminer-stop", stop_overlay)
	mp.register_script_message("subminer-toggle", toggle_overlay)
	mp.register_script_message("subminer-toggle-invisible", toggle_invisible_overlay)
	mp.register_script_message("subminer-show-invisible", show_invisible_overlay)
	mp.register_script_message("subminer-hide-invisible", hide_invisible_overlay)
	mp.register_script_message("subminer-menu", show_menu)
	mp.register_script_message("subminer-options", open_options)
	mp.register_script_message("subminer-restart", restart_overlay)
	mp.register_script_message("subminer-status", check_status)
	mp.register_script_message(HOVER_MESSAGE_NAME, function(payload_json)
		handle_hover_message(payload_json)
	end)
	mp.register_script_message(HOVER_MESSAGE_NAME_LEGACY, function(payload_json)
		handle_hover_message(payload_json)
	end)
end

local function init()
	register_keybindings()
	register_script_messages()

	mp.register_event("file-loaded", on_file_loaded)
	mp.register_event("shutdown", on_shutdown)
	mp.register_event("file-loaded", clear_hover_overlay)
	mp.register_event("end-file", clear_hover_overlay)
	mp.register_event("shutdown", clear_hover_overlay)
	mp.add_hook("on_unload", 10, function()
		clear_hover_overlay()
	end)
	mp.observe_property("sub-start", "native", function()
		clear_hover_overlay()
	end)

	subminer_log("info", "lifecycle", "SubMiner plugin loaded")
end

init()
