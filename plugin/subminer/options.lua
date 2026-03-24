local M = {}
local DEFAULT_ANISKIP_BUTTON_KEY = "TAB"

local function normalize_socket_path_option(socket_path, default_socket_path)
	if type(default_socket_path) ~= "string" then
		return socket_path
	end

	local trimmed_default = default_socket_path:match("^%s*(.-)%s*$")
	local trimmed_socket = type(socket_path) == "string" and socket_path:match("^%s*(.-)%s*$") or socket_path
	if trimmed_default ~= "\\\\.\\pipe\\subminer-socket" then
		return trimmed_socket
	end
	if type(trimmed_socket) ~= "string" or trimmed_socket == "" then
		return trimmed_default
	end
	if trimmed_socket == "/tmp/subminer-socket" or trimmed_socket == "\\tmp\\subminer-socket" then
		return trimmed_default
	end
	if trimmed_socket == "\\\\.\\pipe\\tmp\\subminer-socket" then
		return trimmed_default
	end
	return trimmed_socket
end

function M.load(options_lib, default_socket_path)
	local opts = {
		binary_path = "",
		socket_path = default_socket_path,
		texthooker_enabled = true,
		texthooker_port = 5174,
		backend = "auto",
		auto_start = true,
		auto_start_visible_overlay = true,
		auto_start_pause_until_ready = true,
		auto_start_pause_until_ready_timeout_seconds = 15,
		osd_messages = true,
		log_level = "info",
		aniskip_enabled = true,
		aniskip_title = "",
		aniskip_season = "",
		aniskip_mal_id = "",
		aniskip_episode = "",
		aniskip_payload = "",
		aniskip_show_button = true,
		aniskip_button_text = "You can skip by pressing %s",
		aniskip_button_key = DEFAULT_ANISKIP_BUTTON_KEY,
		aniskip_button_duration = 3,
	}

	options_lib.read_options(opts, "subminer")
	opts.socket_path = normalize_socket_path_option(opts.socket_path, default_socket_path)
	return opts
end

function M.coerce_bool(value, fallback)
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

return M
