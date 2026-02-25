local M = {}

function M.new()
	return {
		overlay_running = false,
		texthooker_running = false,
		overlay_process = nil,
		binary_available = false,
		binary_path = nil,
		invisible_overlay_visible = false,
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
end

return M
