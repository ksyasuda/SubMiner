import test from 'node:test';
import assert from 'node:assert/strict';
import {
  buildJellyfinTimelinePayload,
  JellyfinPlaybackReporter,
} from './jellyfin-playback-reporter';

type FetchCall = { input: string; init: RequestInit };

function createReporter(respond: (input: string) => Response, logWarn?: (message: string) => void) {
  const calls: FetchCall[] = [];
  const reporter = new JellyfinPlaybackReporter({
    serverUrl: 'http://jellyfin.local/',
    accessToken: 'token-1',
    deviceId: 'device-1',
    clientName: 'SubMiner',
    clientVersion: '1.2.3',
    deviceName: 'host',
    fetchImpl: (async (input, init) => {
      calls.push({ input: String(input), init: init ?? {} });
      return respond(String(input));
    }) as typeof fetch,
    logWarn,
  });
  return { reporter, calls };
}

test('reportProgress posts the timeline payload as this device and treats failure as non-fatal', async () => {
  let fail = false;
  const { reporter, calls } = createReporter(
    () => new Response(fail ? 'boom' : null, { status: fail ? 500 : 200 }),
  );
  const state = {
    itemId: 'movie-2',
    positionTicks: 123456,
    isPaused: true,
    volumeLevel: 33,
    audioStreamIndex: 1,
    subtitleStreamIndex: 2,
  };

  assert.equal(await reporter.reportProgress(state), true);
  fail = true;
  assert.equal(await reporter.reportProgress({ itemId: 'movie-2', positionTicks: 999 }), false);

  const call = calls[0]!;
  assert.equal(call.input, 'http://jellyfin.local/Sessions/Playing/Progress');
  assert.deepEqual(
    JSON.parse(String(call.init.body)),
    JSON.parse(JSON.stringify(buildJellyfinTimelinePayload(state))),
  );
  assert.equal(
    (call.init.headers as Record<string, string>).Authorization,
    'MediaBrowser Client="SubMiner", Device="host", DeviceId="device-1", Version="1.2.3", Token="token-1"',
  );
});

test('timeline payload omits websocket-only event names', () => {
  const payload = buildJellyfinTimelinePayload({
    itemId: 'movie-2',
    positionTicks: 123456,
    eventName: 'TimeUpdate',
  });

  assert.equal('EventName' in payload, false);
});

test('reportStopped posts final position and explicit non-failed state', async () => {
  const { reporter, calls } = createReporter(() => new Response(null, { status: 200 }));

  assert.equal(
    await reporter.reportStopped({ itemId: 'movie-stop', positionTicks: 7654321 }),
    true,
  );

  assert.equal(calls[0]!.input, 'http://jellyfin.local/Sessions/Playing/Stopped');
  const posted = JSON.parse(String(calls[0]!.init.body));
  assert.equal(posted.PositionTicks, 7654321);
  assert.equal(posted.Failed, false);
});

test('warns once per failing endpoint until it recovers', async () => {
  const warnings: string[] = [];
  let status = 400;
  const { reporter } = createReporter(
    () => new Response(null, { status }),
    (message) => warnings.push(message),
  );
  const state = { itemId: 'item-1', positionTicks: 10, playMethod: 'DirectPlay' };

  assert.equal(await reporter.reportStopped(state), false);
  assert.equal(await reporter.reportStopped(state), false);
  assert.equal(warnings.length, 1);
  assert.match(warnings[0]!, /Sessions\/Playing\/Stopped/);

  status = 200;
  assert.equal(await reporter.reportStopped(state), true);
  status = 500;
  assert.equal(await reporter.reportStopped(state), false);
  assert.equal(warnings.length, 2);
});
