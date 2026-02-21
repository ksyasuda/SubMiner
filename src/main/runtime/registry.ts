import * as domains from './domains';

export type MainRuntimeRegistry = typeof domains;

export function createMainRuntimeRegistry(): MainRuntimeRegistry {
  return domains;
}
