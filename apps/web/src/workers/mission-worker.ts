import {
  stageMissionMutation,
  type MissionMutationRequest,
} from '../lib/mission-execution-core.js';

self.addEventListener('message', (event: MessageEvent<MissionMutationRequest>) => {
  try {
    const result = stageMissionMutation(event.data, (progress) =>
      self.postMessage({ type: 'progress', progress }),
    );
    self.postMessage({ type: 'result', result });
  } catch (error) {
    self.postMessage({
      type: 'error',
      error: error instanceof Error ? error.message : 'Mission worker failed.',
    });
  }
});
