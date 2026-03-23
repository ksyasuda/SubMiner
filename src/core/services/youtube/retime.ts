export async function retimeYoutubeSubtitle(input: {
  primaryPath: string;
  secondaryPath: string | null;
}): Promise<{ ok: boolean; path: string; strategy: 'none'; message: string }> {
  return {
    ok: true,
    path: input.primaryPath,
    strategy: 'none',
    message: `Using downloaded subtitle as-is${input.secondaryPath ? ' (no automatic retime enabled)' : ''}`,
  };
}
