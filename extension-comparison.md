# Chrome vs Firefox Extension Comparison

This document provides a side-by-side comparison of the Chrome MV3 and Firefox MV2 versions of the Teams Shifts Exporter extension.

## Summary

| Aspect | Chrome (MV3) | Firefox (MV2) |
|--------|--------------|---------------|
| Manifest Version | MV3 | MV2 |
| Background Script Type | Service Worker (persistent) | Event Page (non-persistent) |
| API Namespace | `chrome.*` | `browser.*` |
| Storage API | `chrome.storage.local` | `browser.storage.local` |
| Alarms API | `chrome.alarms` | `browser.alarms` |
| Tabs API | `chrome.tabs` | `browser.tabs` |
| WebRequest API | `chrome.webRequest` | `browser.webRequest` |
| Scripting API | `chrome.scripting` | `browser.tabs.executeScript` |
| Downloads API | `chrome.downloads` with offscreen document | `browser.downloads` directly |
| Notifications API | `chrome.notifications` | Not implemented |

---

## Key Differences by Area

### 1. Manifest Configuration

#### Chrome (MV3) - `manifest.json`
```json
{
  "manifest_version": 3,
  "background": {
    "service_worker": "background.js"
  },
  "host_permissions": [
    "*://teams.cloud.microsoft/*",
    "*://flw.teams.cloud.microsoft/*",
    "*://outlook.office.com/*",
    "*://outlook.office365.com/*"
  ],
  "permissions": [
    "alarms",
    "notifications",
    "scripting",
    "webRequest",
    "storage",
    "offscreen"
  ]
}
```

#### Firefox (MV2) - `manifest.json`
```json
{
  "manifest_version": 2,
  "background": {
    "scripts": ["background.js"],
    "persistent": false
  },
  "permissions": [
    "alarms",
    "webRequest",
    "webRequestBlocking",
    "storage",
    "tabs",
    "http://*/*",
    "https://*/*"
  ],
  "host_permissions": [
    "*://teams.cloud.microsoft/*",
    "*://flw.teams.cloud.microsoft/*",
    "*://outlook.office.com/*",
    "*://outlook.office365.com/*"
  ]
}
```

### 2. Background Script Differences

#### Module System
- **Chrome (MV3)**: ES Modules with `import`/`export`
- **Firefox (MV2)**: Legacy script with `require`/`module.exports` (though the provided code uses global scope)

#### Promise Handling
- **Chrome (MV3)**: Uses `Promise` directly
- **Firefox (MV2)**: Uses `Promise` directly, but some APIs return different object structures

### 3. Storage API Differences

#### Chrome (MV3)
```javascript
chrome.storage.local.set({ key: value });
chrome.storage.local.get(['key'], (data) => { ... });
chrome.storage.local.remove(['key']);
```

#### Firefox (MV2)
```javascript
browser.storage.local.set({ key: value });
browser.storage.local.get(['key']).then((data) => { ... });
browser.storage.local.remove(['key']);
```

### 4. Alarm API Differences

#### Chrome (MV3)
```javascript
chrome.alarms.create(name, { periodInMinutes: 60 });
chrome.alarms.get(name).then((alarm) => { ... });
chrome.alarms.clear(name);
```

#### Firefox (MV2)
```javascript
browser.alarms.create(name, { periodInMinutes: 60 });
browser.alarms.get(name).then((alarm) => { ... });
browser.alarms.clear(name);
```

### 5. Web Request API Differences

#### Chrome (MV3)
```javascript
chrome.webRequest.onBeforeSendHeaders.addListener(
  handler,
  { urls: ['<all_urls>'] },
  ['requestHeaders', 'blocking']
);
```

#### Firefox (MV2)
```javascript
browser.webRequest.onBeforeSendHeaders.addListener(
  handler,
  { urls: ['<all_urls>'] },
  ['requestHeaders', 'blocking']
);
```

### 6. Scripting API Differences

#### Chrome (MV3) - `chrome.scripting`
```javascript
await chrome.scripting.executeScript({
  target: { tabId: tabId },
  func: someFunction,
  args: [arg1, arg2]
});
```

#### Firefox (MV2) - `browser.tabs.executeScript`
```javascript
await browser.tabs.executeScript(tabId, {
  code: `(${someFunction})(${JSON.stringify(arg1)}, ${JSON.stringify(arg2)})`
});
```

### 7. Downloads API Differences

#### Chrome (MV3)
```javascript
chrome.downloads.download({ url, filename, saveAs: false });
```

#### Firefox (MV2)
```javascript
browser.downloads.download({ url, filename, saveAs: false });
```

### 8. Notification API Differences

#### Chrome (MV3)
```javascript
chrome.notifications.create(id, options, (notificationId) => { ... });
chrome.notifications.clear(id);
```

#### Firefox (MV2)
```javascript
browser.notifications.create(id, options);
browser.notifications.clear(id);
```

### 9. Message Passing Differences

#### Chrome (MV3)
```javascript
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  // ...
  return true; // keep message channel open
});
```

#### Firefox (MV2)
```javascript
browser.runtime.onMessage.addListener((message, sender) => {
  // ...
  return Promise.resolve(response); // return promise for async
});
```

### 10. Runtime API Differences

#### Chrome (MV3)
```javascript
chrome.runtime.getManifest();
chrome.runtime.openOptionsPage();
chrome.runtime.reload();
chrome.runtime.disconnect();
```

#### Firefox (MV2)
```javascript
browser.runtime.getManifest();
browser.runtime.openOptionsPage();
// Note: reload() and disconnect() behave differently
```

### 11. Tabs API Differences

#### Chrome (MV3)
```javascript
chrome.tabs.create({ url, windowId });
chrome.tabs.update(tabId, { active: true });
chrome.tabs.remove(tabId);
chrome.tabs.query({ active: true, currentWindow: true });
```

#### Firefox (MV2)
```javascript
browser.tabs.create({ url, windowId });
browser.tabs.update(tabId, { active: true });
browser.tabs.remove(tabId);
browser.tabs.query({ active: true, currentWindow: true });
```

### 12. Content Security Policy (CSP) Differences

#### Chrome (MV3)
```json
"content_security_policy": {
  "extension_pages": "script-src 'self'; object-src 'self'"
}
```

#### Firefox (MV2)
```json
"content_security_policy": "script-src 'self'; object-src 'self'"
```

### 13. Manifest V3 Specific Features (Chrome Only)

#### Service Worker Lifecycle
- Chrome MV3 uses service workers which have different lifecycle behavior
- Service workers can be terminated when not actively processing events
- No access to `window`, `document`, or DOM APIs in background scripts

#### Event Pages vs Service Workers
- Firefox MV2 uses event pages (can be terminated)
- Chrome MV3 uses service workers (more restrictive lifecycle)

### 14. API Compatibility Notes

#### Firefox-Specific Limitations
1. **No `chrome.scripting` API**: Must use `browser.tabs.executeScript`
2. **Different `browser.runtime.onMessage` behavior**: Must return promises
3. **Different `browser.storage.local.get` return type**: Direct object vs callback
4. **No `chrome.notifications.create` callback**: Returns promise instead

#### Chrome-Specific Features
1. **`chrome.scripting` API**: More powerful script injection
2. **`chrome.runtime.onMessage`**: Return `true` to keep channel open
3. **`chrome.storage.local.get`**: Callback-based API

### 15. Code Structure Differences

#### Chrome (MV3) Module Import/Export
```javascript
// background.js
import { helperFunction } from './helper.js';

// helper.js
export function helperFunction() { ... }
```

#### Firefox (MV2) Module System
```javascript
// background.js
const helperFunction = require('./helper.js').helperFunction;

// helper.js
module.exports = { helperFunction };
```

### 16. Compatibility Layer

The provided Firefox code uses a compatibility layer to make `browser.*` APIs work like `chrome.*`:

```javascript
// Firefox compatibility shim
if (typeof browser === 'undefined') {
  var browser = chrome;
}
```

This allows the same code to work in both browsers with minimal changes.

---

## Summary of Key Differences

| Feature | Chrome (MV3) | Firefox (MV2) |
|---------|--------------|---------------|
| Manifest Version | 3 | 2 |
| Background Script | Service Worker | Event Page |
| Module System | ES Modules | Legacy (or ES Modules with shim) |
| Promise Support | Native | Native |
| Storage API | `chrome.storage.local` | `browser.storage.local` |
| Alarm API | `chrome.alarms` | `browser.alarms` |
| Scripting API | `chrome.scripting` | `browser.tabs.executeScript` |
| Message Passing | Callback with `return true` | Promise return |
| CSP Format | Object | String |
| Notifications | Callback | Promise |
| Downloads | `chrome.downloads` | `browser.downloads` |

---

## Migration Notes

### From Chrome (MV3) to Firefox (MV2)

1. Replace `chrome.*` with `browser.*` or use a compatibility shim
2. Update message listeners to return promises
3. Replace `chrome.scripting.executeScript` with `browser.tabs.executeScript`
4. Update storage API calls to use promises instead of callbacks
5. Adjust CSP format from object to string

### From Firefox (MV2) to Chrome (MV3)

1. Replace `browser.*` with `chrome.*`
2. Update message listeners to return `true` for async responses
3. Replace `browser.tabs.executeScript` with `chrome.scripting.executeScript`
4. Update storage API calls to use callbacks instead of promises
5. Adjust CSP format from string to object
6. Consider using ES modules for better code organization

---

## Testing Considerations

### Chrome (MV3)
- Test service worker lifecycle (termination/restart)
- Verify ES module imports work correctly
- Check CSP compliance
- Test notification callbacks

### Firefox (MV2)
- Test event page lifecycle
- Verify promise-based APIs work correctly
- Check CSP string format
- Test message passing with promise returns

---

## Future Considerations

### Firefox MV3 Support
Firefox is working on Manifest V3 support, which will:
- Use service workers instead of event pages
- Support ES modules
- Use `browser.*` APIs (already supported)
- Support `chrome.*` APIs (via compatibility layer)

### Cross-Browser Development
For best cross-browser compatibility:
1. Use `browser.*` APIs (supported in both)
2. Use ES modules (supported in both)
3. Use promises (supported in both)
4. Avoid browser-specific features
5. Test in both browsers regularly

---

*This comparison is based on the provided extension code and may not include all possible differences between Chrome MV3 and Firefox MV2.*
