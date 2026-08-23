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

// Lines of a step's shell body that actually execute. Comments are dropped so a
// commented-out command cannot satisfy a "this step runs X" assertion.
export function executableRunLines(step: WorkflowStep): string[] {
  return (typeof step.run === 'string' ? step.run.split('\n') : [])
    .map((line) => line.trim())
    .filter((line) => line.length > 0 && !line.startsWith('#'));
}

// Leading shell keywords and operators that can precede a real command.
const COMMAND_PREFIX = /^(?:if|elif|while|until|then|else|do|!|&&|\|\||\(|\{)\s+/;

// Command positions within a step's shell body: each line split on separators,
// with control-flow prefixes stripped. A pattern anchored with ^ therefore
// matches only where a command actually starts, so text quoted inside an
// `echo`/`printf` argument is not mistaken for the command running.
export function commandPositions(step: WorkflowStep): string[] {
  return executableRunLines(step).flatMap((line) =>
    line
      .split(/;|&&|\|\||(?<![0-9<>])\|(?!\|)/)
      .map((segment) => {
        let candidate = segment.trim();
        let stripped = candidate.replace(COMMAND_PREFIX, '');
        while (stripped !== candidate) {
          candidate = stripped;
          stripped = candidate.replace(COMMAND_PREFIX, '');
        }
        return candidate;
      })
      .filter(Boolean),
  );
}

// Whether a step actually executes a command matching the pattern. Anchor the
// pattern with ^ so it has to match at a command position.
export function stepRunsCommand(step: WorkflowStep, pattern: RegExp): boolean {
  return commandPositions(step).some((position) => pattern.test(position));
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
