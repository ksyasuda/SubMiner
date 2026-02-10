export type ActionWithType = { type: string };

export type ActionHandler<TAction extends ActionWithType> = (
  action: TAction,
) => void | Promise<void>;

export class ActionBus<TAction extends ActionWithType> {
  private handlers = new Map<string, ActionHandler<TAction>>();

  register(type: TAction["type"], handler: ActionHandler<TAction>): void {
    this.handlers.set(type, handler);
  }

  async dispatch(action: TAction): Promise<void> {
    const handler = this.handlers.get(action.type);
    if (!handler) {
      throw new Error(`No handler registered for action: ${action.type}`);
    }
    await handler(action);
  }
}
