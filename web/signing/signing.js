(() => {
  'use strict';

  if (window.__stopSigningInit) {
    return;
  }

  const pluginScriptSourcesDefault = [
    'https://www.cryptopro.ru/sites/default/files/products/cades/cadesplugin_api.js',
    'chrome-extension://iifchhfnnmpdbibifmljnfjhpififfog/nmcades_plugin_api.js',
    'chrome-extension://epiejncknlhcgcanmnmnjnmghjkpgkdd/nmcades_plugin_api.js',
  ];

  const statusBox = document.getElementById('status');
  const errorBox = document.getElementById('error');
  const startBtn = document.getElementById('startBtn');
  const pythonLog = document.getElementById('pythonLog');
  const cadesStatus = document.getElementById('cadesStatus');
  const checkCadesBtn = document.getElementById('checkCadesBtn');
  const fileInfo = document.getElementById('fileInfo');

  const nonce = new URLSearchParams(location.search).get('nonce') || '';
  const state = {
    nonce,
    config: null,
    lastLogId: 0,
  };

  function log(msg) {
    console.log('[BrowserSigning]', msg);
    statusBox.textContent += msg + '\n';
    statusBox.scrollTop = statusBox.scrollHeight;
  }

  function applyPythonLogs(items) {
    if (!items || !items.length) return;
    pythonLog.textContent += items.map((item) => `[Python] ${item}`).join('\n') + '\n';
    pythonLog.scrollTop = pythonLog.scrollHeight;
  }

  function setError(msg) {
    errorBox.textContent = msg || '';
    if (msg) {
      console.error('[BrowserSigning][Error]', msg);
    }
  }

  function setBusy(state) {
    startBtn.disabled = state;
    startBtn.textContent = state ? 'Подписание…' : 'Выбрать сертификат и подписать';
  }

  function setCadesBadge(text, kind = 'muted') {
    cadesStatus.textContent = text;
    cadesStatus.className = `badge ${kind}`;
  }

  function pluginSources() {
    const cfg = state.config;
    if (cfg && Array.isArray(cfg.pluginScriptSources) && cfg.pluginScriptSources.length) {
      return cfg.pluginScriptSources;
    }
    return pluginScriptSourcesDefault;
  }

  function appendScript(url, timeoutMs = 8000) {
    return new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = url;
      script.async = true;

      const timer = setTimeout(() => {
        cleanup();
        reject(new Error('таймаут загрузки скрипта'));
      }, timeoutMs);

      function cleanup() {
        clearTimeout(timer);
        script.onerror = null;
        script.onload = null;
      }

      script.onerror = () => {
        cleanup();
        reject(new Error('ошибка загрузки скрипта'));
      };
      script.onload = () => {
        cleanup();
        resolve();
      };

      document.head.appendChild(script);
    });
  }

  async function ensureCadespluginReady() {
    const plugin = window.cadesplugin;
    if (!plugin) return null;
    if (typeof plugin.then === 'function') {
      await plugin;
    }
    return window.cadesplugin;
  }

  async function loadCadesPlugin() {
    if (window.cadesplugin) {
      return ensureCadespluginReady();
    }

    let lastError = null;
    for (const src of pluginSources()) {
      try {
        log(`Пробуем загрузить cadesplugin_api.js: ${src}`);
        await appendScript(src);
        const plugin = await ensureCadespluginReady();
        if (plugin) return plugin;
        lastError = new Error('cadesplugin_api.js загружен, но объект плагина не появился');
      } catch (e) {
        lastError = e;
        log(`Не удалось загрузить cadesplugin_api.js из ${src}: ${e.message || e}`);
      }
    }

    throw lastError || new Error('Плагин CryptoPro не найден');
  }

  async function checkCades() {
    setCadesBadge('Проверяем доступность API…', 'muted');
    setError('');
    try {
      const plugin = await loadCadesPlugin();
      if (!plugin) {
        setCadesBadge('Расширение/плагин не обнаружены', 'error');
        return;
      }
      await ensureCadespluginReady();
      setCadesBadge('API доступен', 'ok');
      log('Плагин CryptoPro доступен для текущего origin.');
    } catch (e) {
      if (window.cadesplugin) {
        setCadesBadge('Расширение есть, но API недоступен для текущего origin', 'warn');
      } else {
        setCadesBadge('Расширение/плагин не обнаружены', 'error');
      }
      setError(e && e.message ? e.message : String(e));
    }
  }

  async function waitForPluginLoad(timeoutMs = 12000) {
    let timer = null;
    const timeoutPromise = new Promise((_, reject) => {
      timer = setTimeout(
        () => reject(new Error('Плагин CryptoPro не загрузился (таймаут)')),
        timeoutMs,
      );
    });

    try {
      const plugin = await Promise.race([loadCadesPlugin(), timeoutPromise]);
      if (!plugin) {
        throw new Error('Плагин CryptoPro не найден');
      }
      return plugin;
    } finally {
      if (timer) clearTimeout(timer);
    }
  }

  async function sendResult(payload) {
    try {
      await fetch('/result', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
      log('Результат отправлен приложению.');
    } catch (e) {
      setError('Не удалось отправить результат: ' + e);
    }
  }

  async function sign() {
    setError('');
    if (!state.config) {
      setError('Конфигурация страницы не загружена.');
      return;
    }
    log('Проверяем наличие плагина CryptoPro...');
    setBusy(true);
    try {
      const plugin = await waitForPluginLoad();
      if (!plugin) {
        setError('Плагин CryptoPro не найден в браузере.');
        await sendResult({ nonce: state.nonce, status: 'error', error: 'Плагин CryptoPro не найден' });
        return;
      }
      log('Открываем хранилище сертификатов...');
      const store = await plugin.CreateObjectAsync('CAdESCOM.Store');
      await store.Open();
      const certs = await store.Certificates;
      const selected = await certs.Select();
      const count = await selected.Count;
      if (!count) {
        setError('Выбор сертификата отменён.');
        await sendResult({ nonce: state.nonce, status: 'error', error: 'Выбор сертификата отменён' });
        return;
      }
      const cert = await selected.Item(1);
      const signer = await plugin.CreateObjectAsync('CAdESCOM.CPSigner');
      await signer.propset_Certificate(cert);
      log('Сертификат выбран, формируем подпись...');
      const sd = await plugin.CreateObjectAsync('CAdESCOM.CadesSignedData');
      await sd.propset_ContentEncoding(plugin.CADESCOM_BASE64_TO_BINARY);
      await sd.propset_Content(state.config.pdfBase64);
      const signature = await sd.SignCades(signer, plugin.CADESCOM_CADES_BES, true);
      log('Подпись сформирована, отправляем обратно в приложение...');
      await sendResult({ nonce: state.nonce, status: 'ok', signature });
      log('Готово. Теперь можно вернуться в приложение.');
    } catch (e) {
      console.error(e);
      const msg = e && e.message ? e.message : String(e);
      setError(msg);
      await sendResult({ nonce: state.nonce, status: 'error', error: msg });
    } finally {
      setBusy(false);
    }
  }

  async function pollPythonLogs() {
    if (!state.config || !state.config.logEnabled) return;
    try {
      const resp = await fetch(`/logs?nonce=${encodeURIComponent(state.nonce)}&after=${state.lastLogId}`);
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const data = await resp.json();
      applyPythonLogs(data.items || []);
      if (typeof data.last === 'number') {
        state.lastLogId = data.last;
      }
    } catch (e) {
      console.warn('Не удалось получить логи сервера:', e);
    } finally {
      if (state.config && state.config.logEnabled) setTimeout(pollPythonLogs, 1500);
    }
  }

  async function loadConfig() {
    const resp = await fetch(`/config?nonce=${encodeURIComponent(state.nonce)}`);
    if (!resp.ok) {
      throw new Error(`Не удалось загрузить конфигурацию (HTTP ${resp.status})`);
    }
    const data = await resp.json();
    state.config = data;
    state.lastLogId = data.lastLogId || 0;
    if (data.initialLogs) {
      applyPythonLogs(data.initialLogs);
    }
    if (data.pdfName) {
      fileInfo.textContent = `Файл: ${data.pdfName}`;
    }
    log('Конфигурация страницы получена от приложения.');
  }

  async function init() {
    if (!state.nonce) {
      setError('Nonce не передан приложением. Страница запущена напрямую?');
      setBusy(true);
      return;
    }

    try {
      await loadConfig();
      if (state.config && state.config.logEnabled) {
        pollPythonLogs();
      }
      log('Страница готова. Запускаем авто-проверку плагина...');
      checkCades();
    } catch (e) {
      const msg = e && e.message ? e.message : String(e);
      setError(msg);
      setBusy(true);
    }
  }

  startBtn.addEventListener('click', () => {
    log('Запущен процесс подписи по кнопке.');
    sign();
  });

  checkCadesBtn.addEventListener('click', () => {
    log('Ручная проверка плагина.');
    checkCades();
  });

  document.addEventListener('DOMContentLoaded', init);
})();
