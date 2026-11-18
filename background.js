// Clicking the toolbar icon opens the tool in a full tab.
chrome.action.onClicked.addListener(() => {
  chrome.tabs.create({ url: chrome.runtime.getURL('app/page.html') });
});
