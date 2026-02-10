import { SubminerModule } from "./module";

export class ModuleRegistry<TContext = unknown> {
  private readonly modules: SubminerModule<TContext>[] = [];

  register(module: SubminerModule<TContext>): void {
    if (this.modules.some((existing) => existing.id === module.id)) {
      throw new Error(`Module already registered: ${module.id}`);
    }
    this.modules.push(module);
  }

  async initAll(context: TContext): Promise<void> {
    for (const module of this.modules) {
      if (module.init) {
        await module.init(context);
      }
    }
  }

  async startAll(): Promise<void> {
    for (const module of this.modules) {
      if (module.start) {
        await module.start();
      }
    }
  }

  async stopAll(): Promise<void> {
    for (const module of [...this.modules].reverse()) {
      if (module.stop) {
        await module.stop();
      }
    }
  }
}
