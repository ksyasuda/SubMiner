import type { ResolveContext } from '../context';
import { asBoolean, asNumber, asString, isObject } from '../shared';

export function applyProxyResolution(
  context: ResolveContext,
  ankiConnect: Record<string, unknown>,
): void {
  if (isObject(ankiConnect.proxy)) {
    const proxy = ankiConnect.proxy;
    const proxyEnabled = asBoolean(proxy.enabled);
    if (proxyEnabled !== undefined) {
      context.resolved.ankiConnect.proxy.enabled = proxyEnabled;
    } else if (proxy.enabled !== undefined) {
      context.warn(
        'ankiConnect.proxy.enabled',
        proxy.enabled,
        context.resolved.ankiConnect.proxy.enabled,
        'Expected boolean.',
      );
    }

    const proxyHost = asString(proxy.host);
    if (proxyHost !== undefined && proxyHost.trim().length > 0) {
      context.resolved.ankiConnect.proxy.host = proxyHost.trim();
    } else if (proxy.host !== undefined) {
      context.warn(
        'ankiConnect.proxy.host',
        proxy.host,
        context.resolved.ankiConnect.proxy.host,
        'Expected non-empty string.',
      );
    }

    const proxyUpstreamUrl = asString(proxy.upstreamUrl);
    if (proxyUpstreamUrl !== undefined && proxyUpstreamUrl.trim().length > 0) {
      context.resolved.ankiConnect.proxy.upstreamUrl = proxyUpstreamUrl.trim();
    } else if (proxy.upstreamUrl !== undefined) {
      context.warn(
        'ankiConnect.proxy.upstreamUrl',
        proxy.upstreamUrl,
        context.resolved.ankiConnect.proxy.upstreamUrl,
        'Expected non-empty string.',
      );
    }

    const proxyPort = asNumber(proxy.port);
    if (
      proxyPort !== undefined &&
      Number.isInteger(proxyPort) &&
      proxyPort >= 1 &&
      proxyPort <= 65535
    ) {
      context.resolved.ankiConnect.proxy.port = proxyPort;
    } else if (proxy.port !== undefined) {
      context.warn(
        'ankiConnect.proxy.port',
        proxy.port,
        context.resolved.ankiConnect.proxy.port,
        'Expected integer between 1 and 65535.',
      );
    }
  } else if (ankiConnect.proxy !== undefined) {
    context.warn(
      'ankiConnect.proxy',
      ankiConnect.proxy,
      context.resolved.ankiConnect.proxy,
      'Expected object.',
    );
  }
}
