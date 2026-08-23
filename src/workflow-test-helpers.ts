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

// Splits one shell line on command separators, tracking quotes so a separator
// inside a string is not treated as a command break, and stopping at an
// unquoted inline comment.
function splitCommandSeparators(line: string): string[] {
  const segments: string[] = [];
  let current = '';
  let quote: "'" | '"' | null = null;

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index]!;

    if (quote) {
      current += char;
      if (char === '\\' && quote === '"' && index + 1 < line.length) {
        current += line[index + 1]!;
        index += 1;
      } else if (char === quote) {
        quote = null;
      }
      continue;
    }

    if (char === "'" || char === '"') {
      quote = char;
      current += char;
      continue;
    }

    // An unquoted # starts a comment when it opens a word; the rest is inert.
    if (char === '#' && (current === '' || /\s$/.test(current))) {
      break;
    }

    const next = line[index + 1];
    if (char === ';') {
      segments.push(current);
      current = '';
      continue;
    }
    if ((char === '&' || char === '|') && next === char) {
      segments.push(current);
      current = '';
      index += 1;
      continue;
    }
    // A lone pipe separates commands; a redirect such as 2>&1 does not.
    if (char === '|' && !/[0-9<>&]$/.test(current)) {
      segments.push(current);
      current = '';
      continue;
    }

    current += char;
  }

  segments.push(current);
  return segments;
}

// Command positions within a step's shell body: each line split on separators,
// with control-flow prefixes stripped. A pattern anchored with ^ therefore
// matches only where a command actually starts, so text quoted inside an
// `echo`/`printf` argument is not mistaken for the command running.
export function commandPositions(step: WorkflowStep): string[] {
  return executableRunLines(step).flatMap((line) =>
    splitCommandSeparators(line)
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
