// offscreen.js — runs in an offscreen document to create blob URLs for downloads

chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
  if (msg.action === 'OFFSCREEN_DOWNLOAD') {
    const blob = new Blob([msg.content], { type: 'text/calendar;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    chrome.downloads.download(
      { url, filename: msg.filename, saveAs: false, conflictAction: 'overwrite' },
      () => {
        URL.revokeObjectURL(url);
        sendResponse({ ok: true });
      }
    );
    return true; // keep message channel open for async response
  }
});
