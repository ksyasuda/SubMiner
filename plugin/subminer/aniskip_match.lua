local M = {}

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

function M.title_overlap_score(expected_title, candidate_title)
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
		if coverage >= 0.8 then
			score = score + 30
		elseif coverage >= 0.6 then
			score = score + 10
		else
			score = score - 50
		end
	elseif coverage >= 1 then
		score = score + 10
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

function M.season_signal_score(requested_season, candidate_title)
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

return M
