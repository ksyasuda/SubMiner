function sanitizeScriptOptValue(value: string): string {
  return value
    .replace(/,/g, ' ')
    .replace(/[\r\n]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function assertScriptOptPathValue(name: string, value: string): void {
  if (/[,\r\n]/.test(value)) {
    throw new Error(`${name} contains unsupported script option delimiter`);
  }
}

export function buildSubminerScriptOpts(
  appPath: string,
  socketPath: string,
  extraParts: string[] = [],
): string {
  const hasBinaryPath = extraParts.some((part) => part.startsWith('subminer-binary_path='));
  const hasSocketPath = extraParts.some((part) => part.startsWith('subminer-socket_path='));
  if (!hasBinaryPath) {
    assertScriptOptPathValue('subminer-binary_path', appPath);
  }
  if (!hasSocketPath) {
    assertScriptOptPathValue('subminer-socket_path', socketPath);
  }
  const parts = [
    ...(hasBinaryPath ? [] : [`subminer-binary_path=${appPath}`]),
    ...(hasSocketPath ? [] : [`subminer-socket_path=${socketPath}`]),
    ...extraParts.map(sanitizeScriptOptValue),
  ];
  return parts.join(',');
}
