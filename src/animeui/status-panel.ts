import { el } from './dom';
import { describeAnimeBrowserError } from './error-message';

export function createStatusPanel() {
  const status = el<HTMLElement>('status-message');
  const panel = el<HTMLElement>('request-error');
  const title = el<HTMLElement>('request-error-title');
  const explanation = el<HTMLElement>('request-error-explanation');
  const guidance = el<HTMLElement>('request-error-guidance');
  const disclosure = el<HTMLDetailsElement>('request-error-details');
  const technical = el<HTMLElement>('request-error-technical');
  const dismiss = el<HTMLButtonElement>('request-error-dismiss');

  function setStatus(message: string, tone: 'info' | 'ok' | 'error' = 'info'): void {
    const failed = tone === 'error' && message.length > 0;
    panel.classList.toggle('hidden', !failed);
    status.parentElement?.setAttribute('data-tone', tone);
    disclosure.open = false;
    if (!failed) {
      status.textContent = message;
      technical.textContent = '';
      return;
    }

    const error = describeAnimeBrowserError(message);
    status.textContent = '';
    title.textContent = error.title;
    explanation.textContent = error.explanation;
    guidance.textContent = error.guidance;
    guidance.classList.toggle('hidden', !error.guidance);
    technical.textContent = error.details;
    disclosure.classList.toggle('hidden', !error.details);
  }

  dismiss.addEventListener('click', () => {
    setStatus('');
    // Keep keyboard focus in the browser after its dismiss button disappears.
    document.querySelector<HTMLButtonElement>('.tab[aria-selected="true"]')?.focus();
  });

  return { setStatus };
}
