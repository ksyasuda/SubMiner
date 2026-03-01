local M = {}

function M.create(ctx)
	local utils = ctx.utils
	local opts = ctx.opts
	local state = ctx.state
	local environment = ctx.environment
	local subminer_log = ctx.log.subminer_log

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

	local function file_exists(path)
		local info = utils.file_info(path)
		if not info then
			return false
		end
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

	return {
		normalize_binary_path_candidate = normalize_binary_path_candidate,
		file_exists = file_exists,
		find_binary = find_binary,
		ensure_binary_available = ensure_binary_available,
		is_windows = environment.is_windows,
	}
end

return M
