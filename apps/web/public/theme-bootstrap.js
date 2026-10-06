/* global document, window */

(function () {
  try {
    var stored = window.localStorage.getItem('excel_agent_theme');
    var prefersDark =
      typeof window.matchMedia === 'function' &&
      window.matchMedia('(prefers-color-scheme: dark)').matches;
    var theme = stored === 'dark' || stored === 'light' ? stored : prefersDark ? 'dark' : 'light';
    document.documentElement.dataset.theme = theme;
    document.documentElement.style.colorScheme = theme;
    var themeColor = document.querySelector('meta[name="theme-color"]');
    if (themeColor) themeColor.setAttribute('content', theme === 'dark' ? '#0f172a' : '#f8fafc');
  } catch (_error) {
    document.documentElement.dataset.theme = 'light';
  }
})();
