import { readFileSync } from 'node:fs';

export type WorkflowStep = {
  name?: string;
  run?: string;
  env?: Record<string, unknown>;
};

export type ParsedWorkflow = {
  jobs?: Record<string, { steps?: WorkflowStep[] } | undefined>;
};

// Workflow tests only ever run under `bun test`, which parses YAML natively.
function parseWorkflowYaml(source: string): ParsedWorkflow {
  const bunRuntime = globalThis as typeof globalThis & {
    Bun?: { YAML?: { parse?: (input: string) => unknown } };
  };
  const parse = bunRuntime.Bun?.YAML?.parse;
  if (!parse) {
    throw new Error('Bun.YAML.parse is unavailable; workflow tests must run under bun.');
  }
  return parse(source) as ParsedWorkflow;
}

export function readWorkflow(workflowPath: string): ParsedWorkflow {
  return parseWorkflowYaml(readFileSync(workflowPath, 'utf8'));
}

// Steps of one job, in declaration order. Throws on an unknown job so a renamed
// job fails loudly instead of silently emptying an ordering assertion.
export function jobSteps(workflow: ParsedWorkflow, jobName: string): WorkflowStep[] {
  const job = workflow.jobs?.[jobName];
  if (!job) {
    throw new Error(`Workflow has no job named ${jobName}.`);
  }
  return job.steps ?? [];
}

function allSteps(workflow: ParsedWorkflow): Array<{ job: string; step: WorkflowStep }> {
  return Object.entries(workflow.jobs ?? {}).flatMap(([job, definition]) =>
    (definition?.steps ?? []).map((step) => ({ job, step })),
  );
}

// GitHub substitutes ${{ }} into a run script before the shell parses it, so any
// value used that way is executed as script rather than read as data. Reporting
// every expression (rather than allow-listing known-safe ones) also covers
// alternate spellings such as ${{ steps.version.outputs['VERSION'] }}.
export function templateExpressionsInRunBodies(workflow: ParsedWorkflow): string[] {
  return allSteps(workflow).flatMap(({ job, step }) =>
    (typeof step.run === 'string' ? (step.run.match(/\$\{\{[\s\S]*?\}\}/g) ?? []) : []).map(
      (expression) => `${job}/${step.name ?? '<unnamed>'}: ${expression}`,
    ),
  );
}

// Steps whose shell body reads $NAME without the step declaring it in env, which
// would silently expand to an empty string at run time.
export function stepsMissingEnvDeclaration(workflow: ParsedWorkflow, name: string): string[] {
  const reference = new RegExp(`\\$${name}\\b|\\$\\{${name}\\b`);
  return allSteps(workflow)
    .filter(({ step }) => typeof step.run === 'string' && reference.test(step.run))
    .filter(({ step }) => !Object.prototype.hasOwnProperty.call(step.env ?? {}, name))
    .map(({ job, step }) => `${job}/${step.name ?? '<unnamed>'}`);
}
