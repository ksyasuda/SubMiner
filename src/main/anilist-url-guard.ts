const ANILIST_ALLOWED_HOSTS = new Set(['anilist.co', 'www.anilist.co']);

export function isAllowedAnilistExternalUrl(rawUrl: string): boolean {
  try {
    const parsedUrl = new URL(rawUrl);
    return (
      parsedUrl.protocol === 'https:' && ANILIST_ALLOWED_HOSTS.has(parsedUrl.hostname.toLowerCase())
    );
  } catch {
    return false;
  }
}

export function isAllowedAnilistSetupNavigationUrl(rawUrl: string): boolean {
  if (isAllowedAnilistExternalUrl(rawUrl)) {
    return true;
  }
  try {
    const parsedUrl = new URL(rawUrl);
    return parsedUrl.protocol === 'data:';
  } catch {
    return false;
  }
}
