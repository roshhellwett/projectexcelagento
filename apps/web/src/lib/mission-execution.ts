import {
  stageMissionMutation,
  type MissionMutationRequest,
  type MissionMutationResult,
  type MissionProgress,
} from './mission-execution-core.js';
export * from './mission-execution-core.js';

const aborted = () => new DOMException('Mission stopped. Nothing was committed.', 'AbortError');

/** A dedicated worker per mutation lets Stop actually terminate CPU work without affecting imports. */
export async function executeMissionMutation(
  request: MissionMutationRequest,
  options: {
    signal?: AbortSignal;
    onProgress?: (event: MissionProgress) => void;
    timeoutMs?: number;
  } = {},
): Promise<MissionMutationResult> {
  if (options.signal?.aborted) throw aborted();
  if (typeof Worker === 'undefined') {
    // Tests and browsers without Worker use the identical engine contract, never a fake result.
    return stageMissionMutation(request, options.onProgress);
  }
  const worker = new Worker(new URL('../workers/mission-worker.ts', import.meta.url), {
    type: 'module',
  });
  return new Promise((resolve, reject) => {
    const cleanup = () => {
      clearTimeout(timer);
      options.signal?.removeEventListener('abort', onAbort);
      worker.terminate();
    };
    const onAbort = () => {
      cleanup();
      reject(aborted());
    };
    const timer = setTimeout(() => {
      cleanup();
      reject(new Error('Mission execution timed out. Nothing was committed.'));
    }, options.timeoutMs ?? 120000);
    options.signal?.addEventListener('abort', onAbort, { once: true });
    worker.onmessage = (event: MessageEvent) => {
      if (options.signal?.aborted) return;
      if (event.data?.type === 'progress') {
        options.onProgress?.(event.data.progress as MissionProgress);
        return;
      }
      cleanup();
      if (event.data?.type === 'result') resolve(event.data.result as MissionMutationResult);
      else
        reject(
          new Error(String(event.data?.error ?? 'Mission worker failed. Nothing was committed.')),
        );
    };
    worker.onerror = () => {
      cleanup();
      reject(
        new Error(
          'Mission worker failed. Nothing was committed. Please retry or export your workbook.',
        ),
      );
    };
    try {
      worker.postMessage(request);
    } catch (error) {
      cleanup();
      reject(error);
    }
  });
}
