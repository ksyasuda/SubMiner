local function assert_equals(expected, actual, message)
	if expected == actual then
		return
	end
	error(message .. " (expected " .. tostring(expected) .. ", got " .. tostring(actual) .. ")")
end

local function detect_backend(env)
	local original_getenv = os.getenv
	os.getenv = function(name)
		return env[name]
	end

	local recorded_logs = {}
	local environment_module = dofile("plugin/subminer/environment.lua")
	local environment = environment_module.create({
		mp = {
			get_property = function(name)
				if name == "platform" then
					return env.platform or "linux"
				end
				return ""
			end,
			get_script_directory = function()
				return "plugin/subminer"
			end,
		},
		opts = {},
		utils = {
			file_info = function(_)
				return nil
			end,
			join_path = function(...)
				return table.concat({ ... }, "/")
			end,
		},
		log = {
			subminer_log = function(level, scope, line)
				recorded_logs[#recorded_logs + 1] = {
					level = level,
					scope = scope,
					line = line,
				}
			end,
		},
	})

	local ok, backend = pcall(environment.detect_backend)
	os.getenv = original_getenv

	if not ok then
		error(backend)
	end

	return backend, recorded_logs
end

local backend = detect_backend({
	platform = "linux",
	WAYLAND_DISPLAY = "wayland-0",
	XDG_SESSION_TYPE = "wayland",
	XDG_CURRENT_DESKTOP = "KDE",
	XDG_SESSION_DESKTOP = "KDE",
})

assert_equals("kwin", backend, "expected KDE Plasma Wayland to resolve to the kwin backend")

print("plugin environment backend detection tests: OK")
