export type UiSound = 'click' | 'send' | 'success' | 'error' | 'toggle';

const STORAGE_KEY = 'excel_agent_sound_enabled';
let context: AudioContext | null = null;

function soundEnabled(): boolean {
  try {
    if (typeof window === 'undefined') return true;
    return window.localStorage.getItem(STORAGE_KEY) !== 'off';
  } catch {
    return true;
  }
}

export function isSoundEnabled(): boolean {
  return soundEnabled();
}

export function setSoundEnabled(enabled: boolean): void {
  try {
    if (typeof window === 'undefined') return;
    window.localStorage.setItem(STORAGE_KEY, enabled ? 'on' : 'off');
  } catch {
    // Sound is an enhancement; an unavailable storage API should not affect the workspace.
  }
}

function getAudioContext(): AudioContext | null {
  if (typeof window === 'undefined' || !soundEnabled()) return null;
  const AudioContextConstructor =
    window.AudioContext ??
    (window as typeof window & { webkitAudioContext?: typeof AudioContext }).webkitAudioContext;
  if (!AudioContextConstructor) return null;
  try {
    context ??= new AudioContextConstructor();
  } catch {
    return null;
  }
  if (context.state === 'suspended') void context.resume().catch(() => undefined);
  return context;
}

/** Small, synthesized cues keep the app self-contained and avoid shipping an audio asset. */
export function playUiSound(sound: UiSound): void {
  const audio = getAudioContext();
  if (!audio) return;

  const now = audio.currentTime;
  const presets: Record<UiSound, { frequencies: number[]; duration: number; volume: number }> = {
    click: { frequencies: [520], duration: 0.045, volume: 0.018 },
    toggle: { frequencies: [420, 620], duration: 0.07, volume: 0.022 },
    send: { frequencies: [360, 520], duration: 0.1, volume: 0.024 },
    success: { frequencies: [520, 660, 820], duration: 0.18, volume: 0.028 },
    error: { frequencies: [220, 170], duration: 0.14, volume: 0.024 },
  };
  const preset = presets[sound];

  preset.frequencies.forEach((frequency, index) => {
    const oscillator = audio.createOscillator();
    const gain = audio.createGain();
    const start = now + index * (preset.duration / Math.max(preset.frequencies.length, 2));
    oscillator.type = sound === 'error' ? 'triangle' : 'sine';
    oscillator.frequency.setValueAtTime(frequency, start);
    gain.gain.setValueAtTime(0.0001, start);
    gain.gain.exponentialRampToValueAtTime(preset.volume, start + 0.008);
    gain.gain.exponentialRampToValueAtTime(0.0001, start + preset.duration);
    oscillator.connect(gain);
    gain.connect(audio.destination);
    oscillator.start(start);
    oscillator.stop(start + preset.duration + 0.015);
  });
}
