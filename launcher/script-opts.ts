function sanitizeScriptOptValue(value: string): string {
  return value
    .replace(/,/g, ' ')
    .replace(/[\r\n]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function buildSubminerScriptOpts(
  appPath: string,
  socketPath: string,
  extraParts: string[] = [],
): string {
  const hasBinaryPath = extraParts.some((part) => part.startsWith('subminer-binary_path='));
  const hasSocketPath = extraParts.some((part) => part.startsWith('subminer-socket_path='));
  const parts = [
    ...(hasBinaryPath ? [] : [`subminer-binary_path=${sanitizeScriptOptValue(appPath)}`]),
    ...(hasSocketPath ? [] : [`subminer-socket_path=${sanitizeScriptOptValue(socketPath)}`]),
    ...extraParts.map(sanitizeScriptOptValue),
  ];
  return parts.join(',');
}
