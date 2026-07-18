import type { Hono } from 'hono';
import { createStatsStaticResponse } from './route-support.js';

export function registerStatsStaticRoutes(app: Hono, staticDir?: string): void {
  if (!staticDir) return;

  app.get('/assets/*', (c) => {
    const response = createStatsStaticResponse(staticDir, c.req.path);
    if (!response) return c.text('Not found', 404);
    return response;
  });

  app.get('/index.html', (c) => {
    const response = createStatsStaticResponse(staticDir, '/index.html');
    if (!response) return c.text('Stats UI not built', 404);
    return response;
  });

  app.get('*', (c) => {
    const staticResponse = createStatsStaticResponse(staticDir, c.req.path);
    if (staticResponse) return staticResponse;
    const fallback = createStatsStaticResponse(staticDir, '/index.html');
    if (!fallback) return c.text('Stats UI not built', 404);
    return fallback;
  });
}
