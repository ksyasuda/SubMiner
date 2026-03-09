local function assert_equal(actual, expected, message)
	if actual == expected then
		return
	end
	error((message or "assert_equal failed") .. "\nexpected: " .. tostring(expected) .. "\nactual: " .. tostring(actual))
end

local function assert_true(condition, message)
	if condition then
		return
	end
	error(message or "assert_true failed")
end

local function with_env(env, callback)
	local original_getenv = os.getenv
	os.getenv = function(name)
		local value = env[name]
		if value ~= nil then
			return value
		end
		return original_getenv(name)
	end

	local ok, result = pcall(callback)
	os.getenv = original_getenv
	if not ok then
		error(result)
	end
	return result
end

local function create_binary_module(config)
	local binary_module = dofile("plugin/subminer/binary.lua")
	local entries = config.entries or {}

	local binary = binary_module.create({
		mp = config.mp,
		utils = {
			file_info = function(path)
				local entry = entries[path]
				if entry == "file" then
					return { is_dir = false }
				end
				if entry == "dir" then
					return { is_dir = true }
				end
				return nil
			end,
			join_path = function(...)
				return table.concat({ ... }, "\\")
			end,
		},
		opts = {
			binary_path = config.binary_path or "",
		},
		state = {},
		environment = {
			is_windows = function()
				return config.is_windows == true
			end,
		},
		log = {
			subminer_log = function() end,
		},
	})

	return binary
end

do
	local binary = create_binary_module({
		is_windows = true,
		binary_path = "C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner\\SubMiner",
		entries = {
			["C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner\\SubMiner.exe"] = "file",
		},
	})

	assert_equal(
		binary.find_binary(),
		"C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner\\SubMiner.exe",
		"windows resolver should append .exe for configured binary_path"
	)
end

do
	local binary = create_binary_module({
		is_windows = true,
		mp = {
			command_native = function(command)
				local args = command.args or {}
				if args[1] == "powershell.exe" then
					return {
						status = 0,
						stdout = "C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner\\SubMiner.exe\n",
						stderr = "",
					}
				end
				return {
					status = 1,
					stdout = "",
					stderr = "unexpected command",
				}
			end,
		},
		entries = {
			["C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner\\SubMiner.exe"] = "file",
		},
	})

	assert_equal(
		binary.find_binary(),
		"C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner\\SubMiner.exe",
		"windows resolver should recover binary from running SubMiner process"
	)
end

do
	local binary = create_binary_module({
		is_windows = true,
		binary_path = "C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner",
		entries = {
			["C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner"] = "dir",
			["C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner\\SubMiner.exe"] = "file",
		},
	})

	assert_equal(
		binary.find_binary(),
		"C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner\\SubMiner.exe",
		"windows resolver should accept install directory binary_path"
	)
end

do
	local resolved = with_env({
		LOCALAPPDATA = "C:\\Users\\tester\\AppData\\Local",
		HOME = "",
		USERPROFILE = "C:\\Users\\tester",
		ProgramFiles = "C:\\Program Files",
		["ProgramFiles(x86)"] = "C:\\Program Files (x86)",
	}, function()
		local binary = create_binary_module({
			is_windows = true,
			entries = {
				["C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner\\SubMiner.exe"] = "file",
			},
		})
		return binary.find_binary()
	end)

	assert_equal(
		resolved,
		"C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner\\SubMiner.exe",
		"windows auto-detection should probe LOCALAPPDATA install path"
	)
end

do
	local resolved = with_env({
		APPDATA = "C:\\Users\\tester\\AppData\\Roaming",
		LOCALAPPDATA = "",
		HOME = "",
		USERPROFILE = "C:\\Users\\tester",
		ProgramFiles = "C:\\Program Files",
		["ProgramFiles(x86)"] = "C:\\Program Files (x86)",
	}, function()
		local binary = create_binary_module({
			is_windows = true,
			entries = {
				["C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner\\SubMiner.exe"] = "file",
			},
		})
		return binary.find_binary()
	end)

	assert_equal(
		resolved,
		"C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner\\SubMiner.exe",
		"windows auto-detection should derive Local install path from APPDATA"
	)
end

do
	local resolved = with_env({
		SUBMINER_BINARY_PATH = "C:\\Portable\\SubMiner\\SubMiner",
	}, function()
		local binary = create_binary_module({
			is_windows = true,
			entries = {
				["C:\\Portable\\SubMiner\\SubMiner.exe"] = "file",
			},
		})
		return binary.find_binary()
	end)

	assert_equal(
		resolved,
		"C:\\Portable\\SubMiner\\SubMiner.exe",
		"windows env override should resolve .exe suffix"
	)
end

do
	local binary = create_binary_module({
		is_windows = true,
		binary_path = "C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner",
		entries = {
			["C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner"] = "dir",
			["C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner\\SubMiner.exe"] = "file",
		},
	})

	assert_true(binary.ensure_binary_available() == true, "ensure_binary_available should cache discovered windows binary")
	assert_equal(
		binary.find_binary(),
		"C:\\Users\\tester\\AppData\\Local\\Programs\\SubMiner\\SubMiner.exe",
		"ensure_binary_available should not break follow-up lookup"
	)
end

print("plugin windows binary resolver tests: OK")
