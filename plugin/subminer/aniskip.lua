local M = {}
local matcher = require("aniskip_match")
local DEFAULT_ANISKIP_BUTTON_KEY = "TAB"

function M.create(ctx)
	local mp = ctx.mp
	local utils = ctx.utils
	local opts = ctx.opts
	local state = ctx.state
	local environment = ctx.environment
	local subminer_log = ctx.log.subminer_log
	local show_osd = ctx.log.show_osd
	local request_generation = 0
	local mal_lookup_cache = {}
	local payload_cache = {}
	local title_context_cache = {}
	local base64_reverse = {}
	local base64_chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

	for i = 1, #base64_chars do
		base64_reverse[base64_chars:sub(i, i)] = i - 1
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

	local function is_remote_media_path()
		local media_path = mp.get_property("path")
		if type(media_path) ~= "string" then
			return false
		end
		local trimmed = media_path:match("^%s*(.-)%s*$") or ""
		if trimmed == "" then
			return false
		end
		return trimmed:match("^%a[%w+.-]*://") ~= nil
	end

	local function parse_json_payload(text)
		if type(text) ~= "string" then
			return nil
		end
		local parsed, parse_error = utils.parse_json(text)
		if type(parsed) == "table" then
			return parsed
		end
		return nil, parse_error
	end

	local function decode_base64(input)
		if type(input) ~= "string" then
			return nil
		end
		local cleaned = input:gsub("%s", ""):gsub("-", "+"):gsub("_", "/")
		cleaned = cleaned:match("^%s*(.-)%s*$") or ""
		if cleaned == "" then
			return nil
		end
		if #cleaned % 4 == 1 then
			return nil
		end
		if #cleaned % 4 ~= 0 then
			cleaned = cleaned .. string.rep("=", 4 - (#cleaned % 4))
		end
		if not cleaned:match("^[A-Za-z0-9+/%=]+$") then
			return nil
		end
		local out = {}
		local out_len = 0
		for index = 1, #cleaned, 4 do
			local c1 = cleaned:sub(index, index)
			local c2 = cleaned:sub(index + 1, index + 1)
			local c3 = cleaned:sub(index + 2, index + 2)
			local c4 = cleaned:sub(index + 3, index + 3)
			local v1 = base64_reverse[c1]
			local v2 = base64_reverse[c2]
			if not v1 or not v2 then
				return nil
			end
			local v3 = c3 == "=" and 0 or base64_reverse[c3]
			local v4 = c4 == "=" and 0 or base64_reverse[c4]
			if (c3 ~= "=" and not v3) or (c4 ~= "=" and not v4) then
				return nil
			end
			local n = (((v1 * 64 + v2) * 64 + v3) * 64 + v4)
			local b1 = math.floor(n / 65536)
			local remaining = n % 65536
			local b2 = math.floor(remaining / 256)
			local b3 = remaining % 256
			out_len = out_len + 1
			out[out_len] = string.char(b1)
			if c3 ~= "=" then
				out_len = out_len + 1
				out[out_len] = string.char(b2)
			end
			if c4 ~= "=" then
				out_len = out_len + 1
				out[out_len] = string.char(b3)
			end
		end
		return table.concat(out)
	end

	local function resolve_launcher_payload()
		local raw_payload = type(opts.aniskip_payload) == "string" and opts.aniskip_payload or ""
		local trimmed = raw_payload:match("^%s*(.-)%s*$") or ""
		if trimmed == "" then
			return nil
		end

		local parsed, parse_error = parse_json_payload(trimmed)
		if type(parsed) == "table" then
			return parsed
		end

		local url_decoded = trimmed:gsub("%%(%x%x)", function(hex)
			local value = tonumber(hex, 16)
			if value then
				return string.char(value)
			end
			return "%"
		end)
		if url_decoded ~= trimmed then
			parsed, parse_error = parse_json_payload(url_decoded)
			if type(parsed) == "table" then
				return parsed
			end
		end

		local b64_decoded = decode_base64(trimmed)
		if type(b64_decoded) == "string" and b64_decoded ~= "" then
			parsed, parse_error = parse_json_payload(b64_decoded)
			if type(parsed) == "table" then
				return parsed
			end
		end

		subminer_log("warn", "aniskip", "Invalid launcher AniSkip payload: " .. tostring(parse_error or "unparseable"))
		return nil
	end

	local function run_json_curl_async(url, callback)
		mp.command_native_async({
			name = "subprocess",
			args = { "curl", "-sL", "--connect-timeout", "5", "-A", "SubMiner-mpv/ani-skip", url },
			playback_only = false,
			capture_stdout = true,
			capture_stderr = true,
		}, function(success, result, error)
			if not success or not result or result.status ~= 0 or type(result.stdout) ~= "string" or result.stdout == "" then
				local detail = error or (result and result.stderr) or "curl failed"
				callback(nil, detail)
				return
			end
			local parsed, parse_error = utils.parse_json(result.stdout)
			if type(parsed) ~= "table" then
				callback(nil, parse_error or "invalid json")
				return
			end
			callback(parsed, nil)
		end)
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
		local cache_key = table.concat({
			tostring(forced_title or ""),
			tostring(forced_season or ""),
			tostring(forced_episode or ""),
			tostring(media_title or ""),
			tostring(filename or ""),
			tostring(path or ""),
		}, "\31")
		local cached = title_context_cache[cache_key]
		if type(cached) == "table" then
			return cached.title, cached.episode, cached.season
		end
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
		title_context_cache[cache_key] = {
			title = candidate_title,
			episode = episode,
			season = forced_season,
		}
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

	local function resolve_mal_id_async(title, season, request_id, callback)
		local forced_mal_id = tonumber(opts.aniskip_mal_id)
		if forced_mal_id and forced_mal_id > 0 then
			callback(forced_mal_id, "(forced-mal-id)")
			return
		end
		if type(title) == "string" and title:match("^%d+$") then
			local numeric = tonumber(title)
			if numeric and numeric > 0 then
				callback(numeric, title)
				return
			end
		end
		if type(title) ~= "string" or title == "" then
			callback(nil, nil)
			return
		end

		local lookup = title
		if season and season > 1 then
			lookup = string.format("%s Season %d", lookup, season)
		end
		local cache_key = string.format("%s|%s", lookup:lower(), tostring(season or "-"))
		local cached = mal_lookup_cache[cache_key]
		if cached ~= nil then
			if cached == false then
				callback(nil, lookup)
			else
				callback(cached, lookup)
			end
			return
		end

		local mal_url = "https://myanimelist.net/search/prefix.json?type=anime&keyword=" .. url_encode(lookup)
		run_json_curl_async(mal_url, function(mal_json, mal_error)
			if request_id ~= request_generation then
				return
			end
			if not mal_json then
				subminer_log("warn", "aniskip", "MAL lookup failed: " .. tostring(mal_error))
				callback(nil, lookup)
				return
			end
			local categories = mal_json.categories
			if type(categories) ~= "table" then
				mal_lookup_cache[cache_key] = false
				callback(nil, lookup)
				return
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
				local matched_id = tonumber(best_item.id)
				mal_lookup_cache[cache_key] = matched_id
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
				callback(matched_id, lookup)
				return
			end
			mal_lookup_cache[cache_key] = false
			callback(nil, lookup)
		end)
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

	local function reset_aniskip_fields()
		state.aniskip.prompt_shown = false
		state.aniskip.found = false
		state.aniskip.mal_id = nil
		state.aniskip.title = nil
		state.aniskip.episode = nil
		state.aniskip.intro_start = nil
		state.aniskip.intro_end = nil
		state.aniskip.payload = nil
		state.aniskip.payload_source = nil
		remove_aniskip_chapters()
	end

	local function clear_aniskip_state()
		request_generation = request_generation + 1
		reset_aniskip_fields()
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
			local key = opts.aniskip_button_key ~= "" and opts.aniskip_button_key or DEFAULT_ANISKIP_BUTTON_KEY
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
						string.format(
							"Intro window %.3f -> %.3f (MAL %s, ep %s)",
							intro_start,
							intro_end,
							tostring(mal_id or "-"),
							tostring(episode or "-")
						)
					)
					return true
				end
			end
		end
		return false
	end

	local function has_launcher_payload()
		return type(opts.aniskip_payload) == "string" and opts.aniskip_payload:match("%S") ~= nil
	end

	local function is_launcher_context()
		local forced_title = type(opts.aniskip_title) == "string" and (opts.aniskip_title:match("^%s*(.-)%s*$") or "") or ""
		if forced_title ~= "" then
			return true
		end
		local forced_mal_id = tonumber(opts.aniskip_mal_id)
		if forced_mal_id and forced_mal_id > 0 then
			return true
		end
		local forced_episode = tonumber(opts.aniskip_episode)
		if forced_episode and forced_episode > 0 then
			return true
		end
		local forced_season = tonumber(opts.aniskip_season)
		if forced_season and forced_season > 0 then
			return true
		end
		if has_launcher_payload() then
			return true
		end
		return false
	end

	local function should_fetch_aniskip_async(trigger_source, callback)
		if is_remote_media_path() then
			callback(false, "remote-url")
			return
		end
		if trigger_source == "script-message" or trigger_source == "overlay-start" then
			callback(true, trigger_source)
			return
		end
		if is_launcher_context() then
			callback(true, "launcher-context")
			return
		end
		if type(environment.is_subminer_app_running_async) == "function" then
			environment.is_subminer_app_running_async(function(running)
				if running then
					callback(true, "subminer-app-running")
				else
					callback(false, "subminer-context-missing")
				end
			end)
			return
		end
		if environment.is_subminer_app_running() then
			callback(true, "subminer-app-running")
			return
		end
		callback(false, "subminer-context-missing")
	end

	local function resolve_lookup_titles(primary_title)
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
		push_lookup_title(primary_title)
		push_lookup_title(media_title_fallback)
		push_lookup_title(filename_fallback)
		push_lookup_title(path_fallback)
		return lookup_titles
	end

	local function resolve_mal_from_candidates_async(lookup_titles, season, request_id, callback, index, last_lookup)
		local current_index = index or 1
		local current_lookup = last_lookup
		if current_index > #lookup_titles then
			callback(nil, current_lookup)
			return
		end
		local lookup_title = lookup_titles[current_index]
		subminer_log("info", "aniskip", string.format('MAL lookup attempt %d/%d using title="%s"', current_index, #lookup_titles, lookup_title))
		resolve_mal_id_async(lookup_title, season, request_id, function(mal_id, lookup)
			if request_id ~= request_generation then
				return
			end
			if mal_id then
				callback(mal_id, lookup)
				return
			end
			resolve_mal_from_candidates_async(lookup_titles, season, request_id, callback, current_index + 1, lookup or current_lookup)
		end)
	end

	local function fetch_payload_for_episode_async(mal_id, episode, request_id, callback)
		local payload_cache_key = string.format("%d:%d", mal_id, episode)
		local cached_payload = payload_cache[payload_cache_key]
		if cached_payload ~= nil then
			if cached_payload == false then
				callback(nil, nil, true)
			else
				callback(cached_payload, nil, true)
			end
			return
		end
		local url = string.format("https://api.aniskip.com/v1/skip-times/%d/%d?types=op&types=ed", mal_id, episode)
		subminer_log("info", "aniskip", string.format("AniSkip URL=%s", url))
		run_json_curl_async(url, function(payload, fetch_error)
			if request_id ~= request_generation then
				return
			end
			if not payload then
				callback(nil, fetch_error, false)
				return
			end
			if payload.found ~= true then
				payload_cache[payload_cache_key] = false
				callback(nil, nil, false)
				return
			end
			payload_cache[payload_cache_key] = payload
			callback(payload, nil, false)
		end)
	end

	local function fetch_payload_from_launcher(payload, mal_id, title, episode)
		if not payload then
			return false
		end
		state.aniskip.payload = payload
		state.aniskip.payload_source = "launcher"
		state.aniskip.mal_id = mal_id
		state.aniskip.title = title
		state.aniskip.episode = episode
		return apply_aniskip_payload(mal_id, title, episode, payload)
	end

	local function fetch_aniskip_for_current_media(trigger_source)
		local trigger = type(trigger_source) == "string" and trigger_source or "manual"
		if not opts.aniskip_enabled then
			clear_aniskip_state()
			return
		end

		should_fetch_aniskip_async(trigger, function(allowed, reason)
			if not allowed then
				subminer_log("debug", "aniskip", "Skipping lookup: " .. tostring(reason))
				return
			end

			request_generation = request_generation + 1
			local request_id = request_generation
			reset_aniskip_fields()
			local title, episode, season = resolve_title_and_episode()
			local lookup_titles = resolve_lookup_titles(title)
			local launcher_payload = resolve_launcher_payload()
			if launcher_payload then
				local launcher_mal_id = tonumber(opts.aniskip_mal_id)
				if not launcher_mal_id then
					launcher_mal_id = nil
				end
				if fetch_payload_from_launcher(launcher_payload, launcher_mal_id, title, episode) then
					subminer_log(
						"info",
						"aniskip",
						string.format(
							"Using launcher-provided AniSkip payload (title=%s, season=%s, episode=%s)",
							tostring(title or ""),
							tostring(season or "-"),
							tostring(episode or "-")
						)
					)
					return
				end
				subminer_log("info", "aniskip", "Launcher payload present but no OP interval was available")
				return
			end

			subminer_log(
				"info",
				"aniskip",
				string.format(
					'Query context: trigger=%s reason=%s title="%s" season=%s episode=%s (opts: title="%s" season=%s episode=%s mal_id=%s; fallback_titles=%d)',
					tostring(trigger),
					tostring(reason or "-"),
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

			resolve_mal_from_candidates_async(lookup_titles, season, request_id, function(mal_id, mal_lookup)
				if request_id ~= request_generation then
					return
				end
				if not mal_id then
					subminer_log("info", "aniskip", string.format('Skipped: MAL id unavailable for query="%s"', tostring(mal_lookup or "")))
					return
				end
				subminer_log("info", "aniskip", string.format('Resolved MAL id=%d using query="%s"', mal_id, tostring(mal_lookup or "")))
				fetch_payload_for_episode_async(mal_id, episode, request_id, function(payload, fetch_error)
					if request_id ~= request_generation then
						return
					end
					if not payload then
						if fetch_error then
							subminer_log("warn", "aniskip", "AniSkip fetch failed: " .. tostring(fetch_error))
						else
							subminer_log("info", "aniskip", "AniSkip: no skip windows found")
						end
						return
					end
					state.aniskip.payload = payload
					state.aniskip.payload_source = "remote"
					if not apply_aniskip_payload(mal_id, title, episode, payload) then
						subminer_log("info", "aniskip", "AniSkip payload did not include OP interval")
					end
				end)
			end)
		end)
	end

	return {
		clear_aniskip_state = clear_aniskip_state,
		skip_intro_now = skip_intro_now,
		update_intro_button_visibility = update_intro_button_visibility,
		fetch_aniskip_for_current_media = fetch_aniskip_for_current_media,
	}
end

return M
