import type { ProviderResponse } from './types.js';

/** Only provider-reported usage is persisted. Missing usage is unknown, never zero. */
export class InferenceUsage {
  private requests = 0;
  private prompt = 0;
  private completion = 0;
  private total = 0;
  private knownPrompt = true;
  private knownCompletion = true;
  private knownTotal = true;
  private last?: ProviderResponse;

  add(response: ProviderResponse): void {
    this.requests += 1;
    this.last = response;
    const valid = (value: number | undefined): value is number =>
      value !== undefined && Number.isFinite(value) && value >= 0;
    const p = response.usage?.promptTokens;
    const c = response.usage?.completionTokens;
    const t = response.usage?.totalTokens ?? (valid(p) && valid(c) ? p + c : undefined);
    if (valid(p)) this.prompt += p;
    else this.knownPrompt = false;
    if (valid(c)) this.completion += c;
    else this.knownCompletion = false;
    if (valid(t)) this.total += t;
    else this.knownTotal = false;
  }

  snapshot() {
    return {
      ...(this.last ? { provider: this.last.provider, model: this.last.model } : {}),
      ...(this.requests && this.knownPrompt ? { promptTokens: this.prompt } : {}),
      ...(this.requests && this.knownCompletion ? { completionTokens: this.completion } : {}),
      ...(this.requests && this.knownTotal ? { totalTokens: this.total } : {}),
    };
  }
}
