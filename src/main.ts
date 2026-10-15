import { App } from './ui/app';
import { launchFiles } from './ui/shell';

const app = new App();
document.addEventListener('copy', (e) => app.onCopyEvent(e, false));
document.addEventListener('cut', (e) => app.onCopyEvent(e, true));

// Open a deck passed as ?file=URL (handy for demos), otherwise start with a blank presentation.
const file = new URLSearchParams(location.search).get('file');
(async () => {
  const launched = await launchFiles().catch(() => []);
  if (launched[0]) {
    await app.open(launched[0].bytes, launched[0].name, undefined, launched[0].path);
    return;
  }
  if (file) {
    try {
      const res = await fetch(file);
      await app.open(new Uint8Array(await res.arrayBuffer()), decodeURIComponent(file.split('/').pop() || 'presentation.pptx'));
      return;
    } catch (e) {
      console.error(e);
    }
  }
  await app.newDocument();
})();

(window as unknown as { opptx: App }).opptx = app;
