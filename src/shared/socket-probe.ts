import net from 'node:net';

/**
 * True when something is accepting connections on the unix socket / Windows
 * named pipe (400 ms probe). Electron-free; shared by the sync CLI's
 * running-app guard and the Windows mpv launch attach wait.
 */
export async function canConnectSocket(socketPath: string): Promise<boolean> {
  return await new Promise<boolean>((resolve) => {
    const socket = net.createConnection(socketPath);
    let settled = false;

    const finish = (value: boolean): void => {
      if (settled) return;
      settled = true;
      try {
        socket.destroy();
      } catch {
        // ignore
      }
      resolve(value);
    };

    socket.once('connect', () => finish(true));
    socket.once('error', () => finish(false));
    socket.setTimeout(400, () => finish(false));
  });
}
