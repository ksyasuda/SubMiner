local M = {}
local matcher = require("aniskip_match")

function M.create(ctx)
	local mp = ctx.mp
	local utils = ctx.utils
	local opts = ctx.opts
	local state = ctx.state
	local environment = ctx.environment
	local subminer_log = ctx.log.subminer_log
	local show_osd = ctx.log.show_osd

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
		local episode = forced_episode or parse_episode_hint(media_title) or parse_episode_hint(filename) or parse_episode_hint(path) or 1
		return candidate_title, episode, forced_season
	end

	local function select_best_mal_item(items, title, season)
		if type(items) ~= "table" then
			return nil
		end
		local best_item = nil
		local best_score = -math.huge
		for _, item in ipairs(items) do
			if type(item) == "table" and tonumber(item.id) then
				local candidate_name = tostring(item.name or "")
				local score = matcher.title_overlap_score(title, candidate_name) + matcher.season_signal_score(season, candidate_name)
				if score > best_score then
					best_score = score
					best_item = item
				end
			end
		end
		return best_item
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

		local all_items = {}
		for _, category in ipairs(categories) do
			if type(category) == "table" and type(category.items) == "table" then
				for _, item in ipairs(category.items) do
					all_items[#all_items + 1] = item
				end
			end
		end
		local best_item = select_best_mal_item(all_items, title, season)
		if best_item and tonumber(best_item.id) then
			subminer_log(
				"info",
				"aniskip",
				string.format(
					'MAL candidate selected (score-based): id=%s name="%s" season_hint=%s',
					tostring(best_item.id),
					tostring(best_item.name or ""),
					tostring(season or "-")
				)
			)
			return tonumber(best_item.id), lookup
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
					subminer_log("info", "aniskip", string.format("Intro window %.3f -> %.3f (MAL %d, ep %d)", intro_start, intro_end, mal_id, episode))
					return true
				end
			end
		end
		return false
	end

	local function fetch_aniskip_for_current_media()
		if not environment.is_subminer_app_running() then
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
			subminer_log("info", "aniskip", string.format('MAL lookup attempt %d/%d using title="%s"', index, #lookup_titles, lookup_title))
			local attempt_mal_id, attempt_lookup = resolve_mal_id(lookup_title, season)
			if attempt_mal_id then
				mal_id = attempt_mal_id
				mal_lookup = attempt_lookup
				break
			end
			mal_lookup = attempt_lookup or mal_lookup
		end
		if not mal_id then
			subminer_log("info", "aniskip", string.format('Skipped: MAL id unavailable for query="%s"', tostring(mal_lookup or "")))
			return
		end
		local url = string.format("https://api.aniskip.com/v1/skip-times/%d/%d?types=op&types=ed", mal_id, episode)
		subminer_log("info", "aniskip", string.format('Resolved MAL id=%d using query="%s"; AniSkip URL=%s', mal_id, tostring(mal_lookup or ""), url))
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

	return {
		clear_aniskip_state = clear_aniskip_state,
		skip_intro_now = skip_intro_now,
		update_intro_button_visibility = update_intro_button_visibility,
		fetch_aniskip_for_current_media = fetch_aniskip_for_current_media,
	}
end

return M
