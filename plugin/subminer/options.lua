local M = {}

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
		aniskip_button_key = "y-k",
		aniskip_button_duration = 3,
	}

	options_lib.read_options(opts, "subminer")
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
