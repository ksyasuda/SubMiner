export interface SubminerModule<TContext = unknown> {
  id: string;
  init?: (context: TContext) => void | Promise<void>;
  start?: () => void | Promise<void>;
  stop?: () => void | Promise<void>;
}
