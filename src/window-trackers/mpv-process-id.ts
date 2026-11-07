import net from 'node:net';

const PID_REQUEST_ID = 1;
const PID_QUERY_TIMEOUT_MS = 5_000;

// This runs inside a socket callback, where a parse error would be an uncaught exception.
function parseIpcLine(line: string): unknown {
  try {
    return JSON.parse(line);
  } catch {
    return null;
  }
}

// Asks the mpv listening on `socketPath` for its process ID over a short-lived IPC connection.
export function queryMpvProcessId(
  socketPath: string,
  timeoutMs = PID_QUERY_TIMEOUT_MS,
): Promise<number> {
  return new Promise((resolve, reject) => {
    const socket = net.createConnection(socketPath);
    let buffered = '';

    const fail = (message: string): void => {
      socket.destroy();
      reject(new Error(message));
    };

    socket.setEncoding('utf8');
    socket.setTimeout(timeoutMs, () => fail(`no answer within ${timeoutMs}ms`));
    socket.once('error', (error) => fail(error.message));
    socket.once('close', () => fail('connection closed before mpv answered'));
    socket.once('connect', () => {
      socket.write(
        `${JSON.stringify({ command: ['get_property', 'pid'], request_id: PID_REQUEST_ID })}\n`,
      );
    });

    socket.on('data', (chunk: string) => {
      buffered += chunk;
      const lines = buffered.split('\n');
      buffered = lines.pop() ?? '';

      // mpv also streams events on this connection, so skip to the line answering the request.
      for (const line of lines) {
        const message = parseIpcLine(line);
        if (
          typeof message !== 'object' ||
          message === null ||
          !('request_id' in message) ||
          message.request_id !== PID_REQUEST_ID
        ) {
          continue;
        }

        const pid = 'data' in message ? message.data : undefined;
        if (typeof pid === 'number') {
          socket.destroy();
          resolve(pid);
        } else {
          fail(`mpv returned no process ID: ${line}`);
        }
        return;
      }
    });
  });
}
