local M = {}

local DEFAULT_HOVER_BASE_COLOR = "FFFFFF"
local DEFAULT_HOVER_COLOR = "C6A0F6"

function M.create(ctx)
	local mp = ctx.mp
	local msg = ctx.msg
	local utils = ctx.utils
	local state = ctx.state

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
		return (text or ""):gsub("\\", "\\\\"):gsub("{", "\\{"):gsub("}", "\\}"):gsub("\n", "\\N")
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
					if candidate_start >= 1 and candidate_stop <= source_len and candidate_stop >= candidate_start and plain:sub(candidate_start, candidate_stop) == token_text then
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

		local before = raw_ass:sub(1, raw_open_idx - 1)
		local hovered = raw_ass:sub(raw_open_idx, raw_close_idx - 1)
		local after = raw_ass:sub(raw_close_idx)
		local hover_suffix = string.format("\\1c&H%s&", hover_color)

		-- Keep hover foreground stable even when inline ASS override tags (\1c/\c/\r) appear inside token.
		hovered = hovered:gsub("{([^}]*)}", function(inner)
			if inner:find("\\1c&H", 1, true) or inner:find("\\c&H", 1, true) or inner:find("\\r", 1, true) then
				return "{" .. inner .. hover_suffix .. "}"
			end
			return "{" .. inner .. "}"
		end)

		local open_tag = string.format("{\\1c&H%s&}", hover_color)
		local close_tag = string.format("{\\1c&H%s&}", base_color)
		return before .. open_tag .. hovered .. close_tag .. after
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
				return
			end
			schedule_hover_clear(0.08)
		else
			clear_hover_overlay()
		end
	end

	return {
		HOVER_MESSAGE_NAME = "subminer-hover-token",
		HOVER_MESSAGE_NAME_LEGACY = "yomipv-hover-token",
		handle_hover_message = handle_hover_message,
		clear_hover_overlay = clear_hover_overlay,
	}
end

return M
