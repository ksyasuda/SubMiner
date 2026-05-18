export function getDefaultMpvSocketPath(platform: NodeJS.Platform = process.platform): string {
  return platform === 'win32' ? '\\\\.\\pipe\\subminer-socket' : '/tmp/subminer-socket';
}
