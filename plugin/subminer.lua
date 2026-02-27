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

local function is_subminer_process_running()
	local command = is_windows() and { "tasklist", "/FO", "CSV", "/NH" } or { "ps", "-A", "-o", "args=" }
	local result = mp.command_native({
		name = "subprocess",
		args = command,
		playback_only = false,
		capture_stdout = true,
		capture_stderr = false,
	})
	if not result or type(result.stdout) ~= "string" or result.status ~= 0 then
		return false
	end

	local process_list = result.stdout:lower()
	for line in process_list:gmatch("[^\\n]+") do
		if is_windows() then
			local image = line:match('^"([^"]+)","')
			if not image then
				image = line:match("^\"([^\"]+)\"")
			end
			if not image then
				goto continue
			end
			if image == "subminer" or image == "subminer.exe" or image == "subminer.appimage" or image == "subminer.app" then
				return true
			end
			if image:find("subminer", 1, true) and not image:find(".lua", 1, true) then
				return true
			end
		else
			local argv0 = line:match('^"([^"]+)"') or line:match("^%s*([^%s]+)")
			if not argv0 then
				goto continue
			end
			if argv0:find("subminer.lua", 1, true) or argv0:find("subminer.conf", 1, true) then
				goto continue
			end
			local exe = argv0:match("([^/\\]+)$") or argv0
			if exe == "SubMiner" or exe == "SubMiner.AppImage" or exe == "SubMiner.exe" or exe == "subminer" or exe == "subminer.appimage" or exe == "subminer.exe" then
				return true
			end
			if exe:find("subminer", 1, true) and exe:find("%.lua", 1, true) == nil and exe:find("%.app", 1, true) == nil then
				return true
			end
		end

		::continue::
	end
	return false
end

local function is_subminer_app_running()
	if is_subminer_process_running() then
		return true
	end
	return false
end

local function is_subminer_ipc_ready()
	if not is_subminer_process_running() then
		return false, "SubMiner process not running"
	end

	if is_windows() then
		return true, nil
	end

	if opts.socket_path ~= default_socket_path() then
		return false, "SubMiner socket path mismatch"
	end

	if not file_exists(default_socket_path()) then
		return false, "SubMiner IPC socket missing at /tmp/subminer-socket"
	end

	return true, nil
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
	osd_messages = true,
	log_level = "info",
	aniskip_enabled = true,
	aniskip_title = "",
	aniskip_season = "",
	aniskip_mal_id = "",
	aniskip_episode = "",
	aniskip_show_button = true,
	aniskip_button_text = "You can skip by pressing %s",
	aniskip_button_key = "y-k",
	aniskip_button_duration = 3,
}

options.read_options(opts, "subminer")

local state = {
	overlay_running = false,
	texthooker_running = false,
	overlay_process = nil,
	binary_available = false,
	binary_path = nil,
	detected_backend = nil,
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
	aniskip = {
		mal_id = nil,
		title = nil,
		episode = nil,
		intro_start = nil,
		intro_end = nil,
		found = false,
		prompt_shown = false,
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

local function url_encode(text)
	if type(text) ~= "string" then
		return ""
	end
	local encoded = text:gsub("\n", " ")
	encoded = encoded:gsub("([^%w%-_%.~ ])", function(char)
		return string.format("%%%02X", string.byte(char))
	end)
	return encoded:gsub(" ", "%%20")
end

local function run_json_curl(url)
	local result = mp.command_native({
		name = "subprocess",
		args = { "curl", "-sL", "--connect-timeout", "5", "-A", "SubMiner-mpv/ani-skip", url },
		playback_only = false,
		capture_stdout = true,
		capture_stderr = true,
	})
	if not result or result.status ~= 0 or type(result.stdout) ~= "string" or result.stdout == "" then
		return nil, result and result.stderr or "curl failed"
	end
	local parsed, parse_error = utils.parse_json(result.stdout)
	if type(parsed) ~= "table" then
		return nil, parse_error or "invalid json"
	end
	return parsed, nil
end

local function parse_episode_hint(text)
	if type(text) ~= "string" or text == "" then
		return nil
	end
	local patterns = {
		"[Ss]%d+[Ee](%d+)",
		"[Ee][Pp]?[%s%._%-]*(%d+)",
		"[%s%._%-]+(%d+)[%s%._%-]+",
	}
	for _, pattern in ipairs(patterns) do
		local token = text:match(pattern)
		if token then
			local episode = tonumber(token)
			if episode and episode > 0 and episode < 10000 then
				return episode
			end
		end
	end
	return nil
end

local function cleanup_title(raw)
	if type(raw) ~= "string" then
		return nil
	end
	local cleaned = raw
	cleaned = cleaned:gsub("%b[]", " ")
	cleaned = cleaned:gsub("%b()", " ")
	cleaned = cleaned:gsub("[Ss]%d+[Ee]%d+", " ")
	cleaned = cleaned:gsub("[Ee][Pp]?[%s%._%-]*%d+", " ")
	cleaned = cleaned:gsub("[%._%-]+", " ")
	cleaned = cleaned:gsub("%s+", " ")
	cleaned = cleaned:match("^%s*(.-)%s*$") or ""
	if cleaned == "" then
		return nil
	end
	return cleaned
end

local function extract_show_title_from_path(media_path)
	if type(media_path) ~= "string" or media_path == "" then
		return nil
	end
	local normalized = media_path:gsub("\\", "/")
	local segments = {}
	for segment in normalized:gmatch("[^/]+") do
		segments[#segments + 1] = segment
	end
	for index = 1, #segments do
		local segment = segments[index] or ""
		if segment:match("^[Ss]eason[%s%._%-]*%d+$") or segment:match("^[Ss][%s%._%-]*%d+$") then
			local prior = segments[index - 1]
			local cleaned = cleanup_title(prior or "")
			if cleaned and cleaned ~= "" then
				return cleaned
			end
		end
	end
	return nil
end

local function normalize_for_match(value)
	if type(value) ~= "string" then
		return ""
	end
	return value:lower():gsub("[^%w]+", " "):gsub("%s+", " "):match("^%s*(.-)%s*$") or ""
end

local MATCH_STOPWORDS = {
	the = true,
	this = true,
	that = true,
	world = true,
	animated = true,
	series = true,
	season = true,
	no = true,
	on = true,
	["and"] = true,
}

local function tokenize_match_words(value)
	local normalized = normalize_for_match(value)
	local tokens = {}
	for token in normalized:gmatch("%S+") do
		if #token >= 3 and not MATCH_STOPWORDS[token] then
			tokens[#tokens + 1] = token
		end
	end
	return tokens
end

local function token_set(tokens)
	local set = {}
	for _, token in ipairs(tokens) do
		set[token] = true
	end
	return set
end

local function title_overlap_score(expected_title, candidate_title)
	local expected = normalize_for_match(expected_title)
	local candidate = normalize_for_match(candidate_title)
	if expected == "" or candidate == "" then
		return 0
	end
	if candidate:find(expected, 1, true) then
		return 120
	end
	local expected_tokens = tokenize_match_words(expected_title)
	local candidate_tokens = token_set(tokenize_match_words(candidate_title))
	if #expected_tokens == 0 then
		return 0
	end
	local score = 0
	local matched = 0
	for _, token in ipairs(expected_tokens) do
		if candidate_tokens[token] then
			score = score + 30
			matched = matched + 1
		else
			score = score - 20
		end
	end
	if matched == 0 then
		score = score - 80
	end
	local coverage = matched / #expected_tokens
	if #expected_tokens >= 2 then
		-- Require strong multi-token agreement to avoid false positives like "Shadow Skill".
		if coverage >= 0.8 then
			score = score + 30
		elseif coverage >= 0.6 then
			score = score + 10
		else
			score = score - 50
		end
	else
		if coverage >= 1 then
			score = score + 10
		end
	end
	return score
end

local function has_any_sequel_marker(candidate_title)
	local normalized = normalize_for_match(candidate_title)
	if normalized == "" then
		return false
	end
	local markers = {
		"season 2",
		"season 3",
		"season 4",
		"2nd season",
		"3rd season",
		"4th season",
		"second season",
		"third season",
		"fourth season",
		" ii ",
		" iii ",
		" iv ",
	}
	local padded = " " .. normalized .. " "
	for _, marker in ipairs(markers) do
		if padded:find(marker, 1, true) then
			return true
		end
	end
	return false
end

local function season_signal_score(requested_season, candidate_title)
	local season = tonumber(requested_season)
	if not season or season < 1 then
		return 0
	end
	local normalized = " " .. normalize_for_match(candidate_title) .. " "
	if normalized == "  " then
		return 0
	end

	if season == 1 then
		return has_any_sequel_marker(candidate_title) and -60 or 20
	end

	local numeric_marker = string.format(" season %d ", season)
	local ordinal_marker = string.format(" %dth season ", season)
	local roman_markers = {
		[2] = { " ii ", " second season ", " 2nd season " },
		[3] = { " iii ", " third season ", " 3rd season " },
		[4] = { " iv ", " fourth season ", " 4th season " },
		[5] = { " v ", " fifth season ", " 5th season " },
	}

	if normalized:find(numeric_marker, 1, true) or normalized:find(ordinal_marker, 1, true) then
		return 40
	end
	local aliases = roman_markers[season] or {}
	for _, marker in ipairs(aliases) do
		if normalized:find(marker, 1, true) then
			return 40
		end
	end
	if has_any_sequel_marker(candidate_title) then
		return -20
	end
	return 5
end

local function resolve_title_and_episode()
	local forced_title = type(opts.aniskip_title) == "string" and (opts.aniskip_title:match("^%s*(.-)%s*$") or "") or ""
	local forced_season = tonumber(opts.aniskip_season)
	local forced_episode = tonumber(opts.aniskip_episode)
	local media_title = mp.get_property("media-title")
	local filename = mp.get_property("filename/no-ext") or mp.get_property("filename") or ""
	local path = mp.get_property("path") or ""
	local path_show_title = extract_show_title_from_path(path)
	local candidate_title = nil
	if path_show_title and path_show_title ~= "" then
		candidate_title = path_show_title
	elseif forced_title ~= "" then
		candidate_title = forced_title
	else
		candidate_title = cleanup_title(media_title) or cleanup_title(filename) or cleanup_title(path)
	end
	local episode = forced_episode
		or parse_episode_hint(media_title)
		or parse_episode_hint(filename)
		or parse_episode_hint(path)
		or 1
	return candidate_title, episode, forced_season
end

local function resolve_mal_id(title, season)
	local forced_mal_id = tonumber(opts.aniskip_mal_id)
	if forced_mal_id and forced_mal_id > 0 then
		return forced_mal_id, "(forced-mal-id)"
	end
	if type(title) == "string" and title:match("^%d+$") then
		local numeric = tonumber(title)
		if numeric and numeric > 0 then
			return numeric, title
		end
	end
	if type(title) ~= "string" or title == "" then
		return nil, nil
	end

	local lookup = title
	if season and season > 1 then
		lookup = string.format("%s Season %d", lookup, season)
	end
	local mal_url = "https://myanimelist.net/search/prefix.json?type=anime&keyword=" .. url_encode(lookup)
	local mal_json, mal_error = run_json_curl(mal_url)
	if not mal_json then
		subminer_log("warn", "aniskip", "MAL lookup failed: " .. tostring(mal_error))
		return nil, lookup
	end
	local categories = mal_json.categories
	if type(categories) ~= "table" then
		return nil, lookup
	end
	for _, category in ipairs(categories) do
		if type(category) == "table" and type(category.items) == "table" then
			for _, item in ipairs(category.items) do
				if type(item) == "table" and tonumber(item.id) then
					subminer_log(
						"info",
						"aniskip",
						string.format(
							'MAL candidate selected (first result): id=%s name="%s" season_hint=%s',
							tostring(item.id),
							tostring(item.name or ""),
							tostring(season or "-")
						)
					)
					return tonumber(item.id), lookup
				end
			end
		end
	end
	return nil, lookup
end

local function set_intro_chapters(intro_start, intro_end)
	if type(intro_start) ~= "number" or type(intro_end) ~= "number" then
		return
	end
	local current = mp.get_property_native("chapter-list")
	local chapters = {}
	if type(current) == "table" then
		for _, chapter in ipairs(current) do
			local title = type(chapter) == "table" and chapter.title or nil
			if type(title) ~= "string" or not title:match("^AniSkip ") then
				chapters[#chapters + 1] = chapter
			end
		end
	end
	chapters[#chapters + 1] = { time = intro_start, title = "AniSkip Intro Start" }
	chapters[#chapters + 1] = { time = intro_end, title = "AniSkip Intro End" }
	table.sort(chapters, function(a, b)
		local a_time = type(a) == "table" and tonumber(a.time) or 0
		local b_time = type(b) == "table" and tonumber(b.time) or 0
		return a_time < b_time
	end)
	mp.set_property_native("chapter-list", chapters)
end

local function remove_aniskip_chapters()
	local current = mp.get_property_native("chapter-list")
	if type(current) ~= "table" then
		return
	end
	local chapters = {}
	local changed = false
	for _, chapter in ipairs(current) do
		local title = type(chapter) == "table" and chapter.title or nil
		if type(title) == "string" and title:match("^AniSkip ") then
			changed = true
		else
			chapters[#chapters + 1] = chapter
		end
	end
	if changed then
		mp.set_property_native("chapter-list", chapters)
	end
end

local function clear_aniskip_state()
	state.aniskip.prompt_shown = false
	state.aniskip.found = false
	state.aniskip.mal_id = nil
	state.aniskip.title = nil
	state.aniskip.episode = nil
	state.aniskip.intro_start = nil
	state.aniskip.intro_end = nil
	remove_aniskip_chapters()
end

local function skip_intro_now()
	if not state.aniskip.found then
		show_osd("Intro skip unavailable")
		return
	end
	local intro_start = state.aniskip.intro_start
	local intro_end = state.aniskip.intro_end
	if type(intro_start) ~= "number" or type(intro_end) ~= "number" then
		show_osd("Intro markers missing")
		return
	end
	local now = mp.get_property_number("time-pos")
	if type(now) ~= "number" then
		show_osd("Skip unavailable")
		return
	end
	local epsilon = 0.35
	if now < (intro_start - epsilon) or now > (intro_end + epsilon) then
		show_osd("Skip intro only during intro")
		return
	end
	mp.set_property_number("time-pos", intro_end)
	show_osd("Skipped intro")
end

local function update_intro_button_visibility()
	if not opts.aniskip_enabled or not opts.aniskip_show_button or not state.aniskip.found then
		return
	end
	local now = mp.get_property_number("time-pos")
	if type(now) ~= "number" then
		return
	end
	local in_intro = now >= (state.aniskip.intro_start or -1) and now < (state.aniskip.intro_end or -1)
	local intro_start = state.aniskip.intro_start or -1
	local hint_window_end = intro_start + 3
	if in_intro and not state.aniskip.prompt_shown and now >= intro_start and now < hint_window_end then
		local key = opts.aniskip_button_key ~= "" and opts.aniskip_button_key or "y-k"
		local message = string.format(opts.aniskip_button_text, key)
		mp.osd_message(message, tonumber(opts.aniskip_button_duration) or 3)
		state.aniskip.prompt_shown = true
	end
end

local function apply_aniskip_payload(mal_id, title, episode, payload)
	local results = payload and payload.results
	if type(results) ~= "table" then
		return false
	end
	for _, item in ipairs(results) do
		if type(item) == "table" and item.skip_type == "op" and type(item.interval) == "table" then
			local intro_start = tonumber(item.interval.start_time)
			local intro_end = tonumber(item.interval.end_time)
			if intro_start and intro_end and intro_end > intro_start then
				state.aniskip.found = true
				state.aniskip.mal_id = mal_id
				state.aniskip.title = title
				state.aniskip.episode = episode
				state.aniskip.intro_start = intro_start
				state.aniskip.intro_end = intro_end
				state.aniskip.prompt_shown = false
				set_intro_chapters(intro_start, intro_end)
				subminer_log(
					"info",
					"aniskip",
					string.format("Intro window %.3f -> %.3f (MAL %d, ep %d)", intro_start, intro_end, mal_id, episode)
				)
				return true
			end
		end
	end
	return false
end

local function fetch_aniskip_for_current_media()
	if not is_subminer_app_running() then
		subminer_log("debug", "lifecycle", "Skipping aniskip lookup: SubMiner app not running")
		return
	end

	clear_aniskip_state()
	if not opts.aniskip_enabled then
		return
	end
	local title, episode, season = resolve_title_and_episode()
	local media_title_fallback = cleanup_title(mp.get_property("media-title"))
	local filename_fallback = cleanup_title(mp.get_property("filename/no-ext") or mp.get_property("filename") or "")
	local path_fallback = cleanup_title(mp.get_property("path") or "")
	local lookup_titles = {}
	local seen_titles = {}
	local function push_lookup_title(candidate)
		if type(candidate) ~= "string" then
			return
		end
		local trimmed = candidate:match("^%s*(.-)%s*$") or ""
		if trimmed == "" then
			return
		end
		local key = trimmed:lower()
		if seen_titles[key] then
			return
		end
		seen_titles[key] = true
		lookup_titles[#lookup_titles + 1] = trimmed
	end
	push_lookup_title(title)
	push_lookup_title(media_title_fallback)
	push_lookup_title(filename_fallback)
	push_lookup_title(path_fallback)

	subminer_log(
		"info",
		"aniskip",
		string.format(
			'Query context: title="%s" season=%s episode=%s (opts: title="%s" season=%s episode=%s mal_id=%s; fallback_titles=%d)',
			tostring(title or ""),
			tostring(season or "-"),
			tostring(episode or "-"),
			tostring(opts.aniskip_title or ""),
			tostring(opts.aniskip_season or "-"),
			tostring(opts.aniskip_episode or "-"),
			tostring(opts.aniskip_mal_id or "-"),
			#lookup_titles
		)
	)
	local mal_id, mal_lookup = nil, nil
	for index, lookup_title in ipairs(lookup_titles) do
		subminer_log(
			"info",
			"aniskip",
			string.format('MAL lookup attempt %d/%d using title="%s"', index, #lookup_titles, lookup_title)
		)
		local attempt_mal_id, attempt_lookup = resolve_mal_id(lookup_title, season)
		if attempt_mal_id then
			mal_id = attempt_mal_id
			mal_lookup = attempt_lookup
			break
		end
		mal_lookup = attempt_lookup or mal_lookup
	end
	if not mal_id then
		subminer_log(
			"info",
			"aniskip",
			string.format('Skipped: MAL id unavailable for query="%s"', tostring(mal_lookup or ""))
		)
		return
	end
	local url = string.format("https://api.aniskip.com/v1/skip-times/%d/%d?types=op&types=ed", mal_id, episode)
	subminer_log(
		"info",
		"aniskip",
		string.format('Resolved MAL id=%d using query="%s"; AniSkip URL=%s', mal_id, tostring(mal_lookup or ""), url)
	)
	local payload, fetch_error = run_json_curl(url)
	if not payload then
		subminer_log("warn", "aniskip", "AniSkip fetch failed: " .. tostring(fetch_error))
		return
	end
	if payload.found ~= true then
		subminer_log("info", "aniskip", "AniSkip: no skip windows found")
		return
	end
	if not apply_aniskip_payload(mal_id, title, episode, payload) then
		subminer_log("info", "aniskip", "AniSkip payload did not include OP interval")
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

local function sanitize_hover_ass_color(input, fallback_rgb)
	local fallback = fix_ass_color(fallback_rgb or DEFAULT_HOVER_COLOR, DEFAULT_HOVER_COLOR)
	local converted = fix_ass_color(input, fallback)
	if converted == "000000" then
		return fallback
	end
	return converted
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
		hover_color = sanitize_hover_ass_color(nil, DEFAULT_HOVER_COLOR),
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
	local hover_color = sanitize_hover_ass_color(payload.colors and payload.colors.hover or nil, DEFAULT_HOVER_COLOR)
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

local function apply_startup_overlay_preferences()
	local should_show_visible = resolve_visible_overlay_startup()

	local visible_action = should_show_visible and "show-visible-overlay" or "hide-visible-overlay"
	if not run_control_command(visible_action) then
		subminer_log("warn", "process", "Failed to apply visible startup action: " .. visible_action)
	end
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
	local socket_ready, reason = is_subminer_ipc_ready()
	local process_not_running = reason == "SubMiner process not running"
	if not socket_ready and not process_not_running then
		subminer_log("warn", "process", "Refusing to start overlay: " .. tostring(reason))
		show_osd("SubMiner IPC not set up. Launch mpv with --input-ipc-server=/tmp/subminer-socket")
		return
	end

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
	if not ensure_binary_available() then
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
	if not ensure_binary_available() then
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
		"Open options",
		"Restart overlay",
		"Check status",
	}

	local actions = {
		start_overlay,
		stop_overlay,
		toggle_overlay,
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
	if not ensure_binary_available() then
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
	if not is_subminer_app_running() then
		clear_aniskip_state()
		subminer_log("debug", "lifecycle", "Skipping file load hooks: SubMiner app not running")
		return true
	end

	clear_aniskip_state()
	fetch_aniskip_for_current_media()
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
	clear_aniskip_state()
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
	mp.add_key_binding("y-y", "subminer-menu", show_menu)
	mp.add_key_binding("y-o", "subminer-options", open_options)
	mp.add_key_binding("y-r", "subminer-restart", restart_overlay)
	mp.add_key_binding("y-c", "subminer-status", check_status)
	if type(opts.aniskip_button_key) == "string" and opts.aniskip_button_key ~= "" then
		mp.add_key_binding(opts.aniskip_button_key, "subminer-skip-intro", skip_intro_now)
	end
	if opts.aniskip_button_key ~= "y-k" then
		mp.add_key_binding("y-k", "subminer-skip-intro-fallback", skip_intro_now)
	end
end

local function register_script_messages()
	mp.register_script_message("subminer-start", start_overlay_from_script_message)
	mp.register_script_message("subminer-stop", stop_overlay)
	mp.register_script_message("subminer-toggle", toggle_overlay)
	mp.register_script_message("subminer-menu", show_menu)
	mp.register_script_message("subminer-options", open_options)
	mp.register_script_message("subminer-restart", restart_overlay)
	mp.register_script_message("subminer-status", check_status)
	mp.register_script_message("subminer-aniskip-refresh", fetch_aniskip_for_current_media)
	mp.register_script_message("subminer-skip-intro", skip_intro_now)
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
	mp.register_event("end-file", clear_aniskip_state)
	mp.register_event("shutdown", clear_aniskip_state)
	mp.add_hook("on_unload", 10, function()
		clear_hover_overlay()
		clear_aniskip_state()
	end)
	mp.observe_property("sub-start", "native", function()
		clear_hover_overlay()
	end)
	mp.observe_property("time-pos", "number", function()
		update_intro_button_visibility()
	end)

	subminer_log("info", "lifecycle", "SubMiner plugin loaded")
end

init()
