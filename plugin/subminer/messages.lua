local M = {}

function M.create(ctx)
	local mp = ctx.mp
	local process = ctx.process
	local aniskip = ctx.aniskip
	local hover = ctx.hover
	local ui = ctx.ui

	local function register_script_messages()
		mp.register_script_message("subminer-start", process.start_overlay_from_script_message)
		mp.register_script_message("subminer-stop", process.stop_overlay)
		mp.register_script_message("subminer-toggle", process.toggle_overlay)
		mp.register_script_message("subminer-toggle-invisible", process.toggle_invisible_overlay)
		mp.register_script_message("subminer-show-invisible", process.show_invisible_overlay)
		mp.register_script_message("subminer-hide-invisible", process.hide_invisible_overlay)
		mp.register_script_message("subminer-menu", ui.show_menu)
		mp.register_script_message("subminer-options", process.open_options)
		mp.register_script_message("subminer-restart", process.restart_overlay)
		mp.register_script_message("subminer-status", process.check_status)
		mp.register_script_message("subminer-aniskip-refresh", aniskip.fetch_aniskip_for_current_media)
		mp.register_script_message("subminer-skip-intro", aniskip.skip_intro_now)
		mp.register_script_message(hover.HOVER_MESSAGE_NAME, function(payload_json)
			hover.handle_hover_message(payload_json)
		end)
		mp.register_script_message(hover.HOVER_MESSAGE_NAME_LEGACY, function(payload_json)
			hover.handle_hover_message(payload_json)
		end)
	end

	return {
		register_script_messages = register_script_messages,
	}
end

return M
