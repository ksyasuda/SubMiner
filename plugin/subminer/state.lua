local M = {}

function M.new()
	return {
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
			payload = nil,
			payload_source = nil,
			found = false,
			prompt_shown = false,
		},
		auto_play_ready_gate_armed = false,
		auto_play_ready_timeout = nil,
		auto_play_ready_osd_timer = nil,
		suppress_ready_overlay_restore = false,
	}
end

return M
