import {
  applyOperationPlan,
  createOperationRegistry,
  invertPatch,
  patchBetween,
  type Workbook,
  type Patch,
  type Report,
  type Preview,
  type OperationErrorCode,
} from '@excel-agent/engine';

export interface MissionMutationRequest {
  workbook: Workbook;
  steps: { operation: string; args: Record<string, unknown> }[];
  confirmed: boolean;
}
export interface MissionProgress {
  index: number;
  phase: 'verifying' | 'verified' | 'failed';
}
export type MissionMutationResult =
  | {
      ok: true;
      workbook: Workbook;
      patch: Patch;
      inverse: Patch;
      steps: { report: Report; preview: Preview }[];
    }
  | {
      ok: false;
      failedStep: number;
      preview?: Preview;
      error: { code: OperationErrorCode; messages: string[]; rolledBack: boolean };
    };

/** Exactly the engine's atomic-plan path, with no caller history and no observable mutation. */
export function stageMissionMutation(
  request: MissionMutationRequest,
  progress?: (event: MissionProgress) => void,
): MissionMutationResult {
  if (request.steps.length < 1 || request.steps.length > 25)
    return {
      ok: false,
      failedStep: 0,
      error: {
        code: 'validation-error',
        messages: ['A mission must have between one and 25 steps.'],
        rolledBack: false,
      },
    };
  const result = applyOperationPlan(request.workbook, request.steps, {
    registry: createOperationRegistry(),
    confirmed: request.confirmed,
    onStep: (index, phase) => progress?.({ index, phase }),
  });
  if (!result.ok)
    return {
      ok: false,
      error: result.error,
      failedStep: result.failedStep,
      preview: result.preview,
    };
  const patch = patchBetween(request.workbook, result.workbook);
  return {
    ok: true,
    workbook: result.workbook,
    patch,
    inverse: invertPatch(patch),
    steps: result.steps.map((step) => ({ report: step.report, preview: step.preview })),
  };
}
