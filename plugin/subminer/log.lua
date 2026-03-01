local M = {}

local LOG_LEVEL_PRIORITY = {
	debug = 10,
	info = 20,
	warn = 30,
	error = 40,
}

function M.create(ctx)
	local mp = ctx.mp
	local msg = ctx.msg
	local opts = ctx.opts

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
			local payload = "SubMiner: " .. message
			local sent = false
			if type(mp.osd_message) == "function" then
				sent = pcall(mp.osd_message, payload, 3)
			end
			if not sent and type(mp.commandv) == "function" then
				pcall(mp.commandv, "show-text", payload, "3000")
			end
		end
	end

	return {
		normalize_log_level = normalize_log_level,
		should_log = should_log,
		subminer_log = subminer_log,
		show_osd = show_osd,
	}
end

return M
