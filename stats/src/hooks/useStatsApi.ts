import { apiClient } from '../lib/api-client';

export type StatsClient = typeof apiClient;

export function getStatsClient(): StatsClient {
  return apiClient;
}
