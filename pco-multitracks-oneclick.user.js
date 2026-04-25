// ==UserScript==
// @name         PCO -> MultiTracks One-Click Table
// @namespace    http://tampermonkey.net/
// @version      1.2.1
// @description  Adds a one-click button on PCO plan pages that builds a Key/BPM/Time-Sig table sourced from MultiTracks and opens it in a new tab.
// @match        https://services.planningcenteronline.com/plans/*
// @match        https://services.planningcenter.com/plans/*
// @homepageURL  https://github.com/HenryGetz/pco-bpm-bandit
// @supportURL   https://github.com/HenryGetz/pco-bpm-bandit/issues
// @updateURL    https://raw.githubusercontent.com/HenryGetz/pco-bpm-bandit/main/pco-multitracks-oneclick.user.js
// @downloadURL  https://raw.githubusercontent.com/HenryGetz/pco-bpm-bandit/main/pco-multitracks-oneclick.user.js
// @grant        unsafeWindow
// @grant        GM.xmlHttpRequest
// @grant        GM.getValue
// @grant        GM.setValue
// @grant        GM.getValues
// @grant        GM.setValues
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_getValues
// @grant        GM_setValues
// @grant        GM_addStyle
// @connect      api.planningcenteronline.com
// @connect      www.multitracks.com
// @connect      multitracks.com
// ==/UserScript==

(function () {
  'use strict';

  const PCO_API_BASE = 'https://api.planningcenteronline.com/services/v2';
  const MT_BASE = 'https://www.multitracks.com';
  const LOCK_TTL_MS = 10000;
  const LOCK_REFRESH_MS = 3000;
  const MAX_CANDIDATES = 6;
  const CONCURRENCY = 3;
  const MAX_RETRIES_PER_SONG = 2;
  const CACHE_KEY = 'tm_pco_mt_song_cache_v1';
  const CACHE_TTL_MS = 1000 * 60 * 60 * 24 * 45;
  const CACHE_MAX_ENTRIES = 1500;
  const BRAND_FONT_URL =
    'https://db.onlinewebfonts.com/c/59360441654e2d78a904e0094b3537b4?family=Code+Pro+W01+Bold+Lowercase';
  const MULTITRACKS_CIRCLE_ICON_PATH =
    'M15.9671 0C10.5648 0 5.78707 2.69091 2.89717 6.80727H14.2462C14.2462 6.80727 15.0594 7.05455 15.0594 7.96364C15.0594 8.87273 14.2172 9.12 14.2172 9.12H1.53935C1.31425 9.59273 1.11095 10.0873 0.92942 10.5818H20.8538C20.8538 10.5818 21.667 10.8364 21.667 11.7455C21.667 12.6545 20.8247 12.9018 20.8247 12.9018H5.11178H0.290443C0.19605 13.4036 0.11618 13.9127 0.0653524 14.4291H28.7974C28.7974 14.4291 29.6106 14.6764 29.6106 15.5855C29.6106 16.4945 28.7683 16.7418 28.7683 16.7418H13.0554H0C0.0217832 17.2436 0.0726123 17.7382 0.137962 18.2255H26.5537C26.5537 18.2255 27.3669 18.4727 27.3669 19.3818C27.3669 20.2909 26.5246 20.5382 26.5246 20.5382H10.819H0.646239C0.798721 21.0545 0.972984 21.5564 1.17629 22.0436H18.2688C18.2688 22.0436 19.0821 22.2909 19.0821 23.2C19.0821 24.1091 18.2398 24.3564 18.2398 24.3564H2.53411H2.35984C5.16261 28.9455 10.2091 32 15.9671 32C24.7893 32 31.9414 24.8364 31.9414 16C31.9414 7.16364 24.7893 0 15.9671 0Z';
  const MULTITRACKS_CIRCLE_ICON_SVG = `<svg viewBox="0 0 32 32" aria-hidden="true" focusable="false" xmlns="http://www.w3.org/2000/svg"><path d="${MULTITRACKS_CIRCLE_ICON_PATH}" fill="currentColor"/></svg>`;

  const DEBUG_LOGGING = (() => {
    try {
      if (typeof GM_getValue === 'function') {
        return Boolean(GM_getValue('tm_pco_mt_debug', false));
      }
    } catch {
      // ignore
    }
    return false;
  })();

  const RUN_ID = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;

  function log(level, message, meta) {
    const ts = new Date().toISOString();
    const prefix = `[TM-PCO-MT][${RUN_ID}][${level}][${ts}]`;

    if (level === 'ERROR') {
      if (meta !== undefined) console.error(prefix, message, meta);
      else console.error(prefix, message);
      return;
    }

    if (level === 'WARN') {
      if (meta !== undefined) console.warn(prefix, message, meta);
      else console.warn(prefix, message);
      return;
    }

    if (level === 'DEBUG' && !DEBUG_LOGGING) return;

    if (meta !== undefined) console.log(prefix, message, meta);
    else console.log(prefix, message);
  }

  function logInfo(message, meta) {
    log('INFO', message, meta);
  }

  function logDebug(message, meta) {
    log('DEBUG', message, meta);
  }

  function logWarn(message, meta) {
    log('WARN', message, meta);
  }

  function logError(message, meta) {
    log('ERROR', message, meta);
  }

  const interceptorState = {
    serviceTypeIdByPlanId: new Map(),
    itemsPayloadByPlanId: new Map(),
  };

  let songCache = null;
  let songCacheDirty = false;

  const gm = {
    async getValue(key, defaultValue) {
      if (typeof GM !== 'undefined' && typeof GM.getValue === 'function') {
        return GM.getValue(key, defaultValue);
      }
      if (typeof GM_getValue === 'function') {
        return GM_getValue(key, defaultValue);
      }
      return defaultValue;
    },

    async setValue(key, value) {
      if (typeof GM !== 'undefined' && typeof GM.setValue === 'function') {
        return GM.setValue(key, value);
      }
      if (typeof GM_setValue === 'function') {
        return GM_setValue(key, value);
      }
      return undefined;
    },

    async getValues(defaultsObj) {
      if (typeof GM !== 'undefined' && typeof GM.getValues === 'function') {
        return GM.getValues(defaultsObj);
      }
      if (typeof GM_getValues === 'function') {
        return GM_getValues(defaultsObj);
      }
      const out = { ...defaultsObj };
      for (const [k, def] of Object.entries(defaultsObj || {})) {
        // eslint-disable-next-line no-await-in-loop
        out[k] = await this.getValue(k, def);
      }
      return out;
    },

    async setValues(valuesObj) {
      if (typeof GM !== 'undefined' && typeof GM.setValues === 'function') {
        return GM.setValues(valuesObj);
      }
      if (typeof GM_setValues === 'function') {
        return GM_setValues(valuesObj);
      }
      const entries = Object.entries(valuesObj || {});
      for (const [k, v] of entries) {
        // eslint-disable-next-line no-await-in-loop
        await this.setValue(k, v);
      }
      return undefined;
    },

    request(details) {
      const xhr =
        (typeof GM !== 'undefined' && GM.xmlHttpRequest) ||
        (typeof GM_xmlhttpRequest !== 'undefined' && GM_xmlhttpRequest);

      if (!xhr) {
        return Promise.reject(new Error('GM.xmlHttpRequest/GM_xmlhttpRequest is unavailable.'));
      }

      return new Promise((resolve, reject) => {
        xhr({
          anonymous: false,
          withCredentials: true,
          timeout: 45000,
          ...details,
          onload: (response) => resolve(response),
          onerror: (error) => reject(error),
          ontimeout: () => reject(new Error(`Request timed out for ${details.url}`)),
        });
      });
    },
  };

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function parsePlanIdFromUrl(urlText) {
    const text = String(urlText || location.href || '');
    const match = text.match(/\/plans\/(\d{5,})/);
    logDebug('parsePlanIdFromUrl evaluated', { text, planId: match ? match[1] : null });
    return match ? match[1] : null;
  }

  function decodeHtmlEntities(value) {
    const textarea = document.createElement('textarea');
    textarea.innerHTML = String(value || '');
    return textarea.value;
  }

  function stripTags(value) {
    return decodeHtmlEntities(String(value || ''))
      .replace(/<[^>]+>/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function normalizeText(value) {
    return String(value || '')
      .normalize('NFD')
      .replace(/\p{Diacritic}/gu, '')
      .toLowerCase()
      .replace(/&/g, ' and ')
      .replace(/[^a-z0-9/ ]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function stripParenthetical(value) {
    return String(value || '').replace(/\([^)]*\)/g, ' ').replace(/\s+/g, ' ').trim();
  }

  function parseBpm(text) {
    const match = String(text || '').match(/(\d+(?:\.\d+)?)/);
    if (!match) return null;
    const value = Number(match[1]);
    return Number.isFinite(value) ? value : null;
  }

  function parseTimeSignature(text) {
    const match = String(text || '').match(/\b(\d+\s*\/\s*\d+)\b/);
    return match ? match[1].replace(/\s+/g, '') : null;
  }

  function jaccardSimilarity(a, b) {
    const aSet = new Set(normalizeText(a).split(' ').filter(Boolean));
    const bSet = new Set(normalizeText(b).split(' ').filter(Boolean));
    if (!aSet.size && !bSet.size) return 1;
    if (!aSet.size || !bSet.size) return 0;

    let intersection = 0;
    for (const token of aSet) {
      if (bSet.has(token)) intersection += 1;
    }
    return intersection / (aSet.size + bSet.size - intersection);
  }

  function escapeHtml(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function formatBpm(value) {
    if (!Number.isFinite(value)) return '—';
    if (Math.abs(value - Math.round(value)) < 0.001) return String(Math.round(value));
    return value.toFixed(2).replace(/\.00$/, '').replace(/0$/, '');
  }

  function buildMarkdownTable(rows) {
    const header = '|#|Song|Version (MT)|Key|Original BPM|Time Sig|';
    const divider = '|--:|---|---|:-:|--:|:-:|';
    const lines = [header, divider];
    for (const row of rows) {
      lines.push(
        `|${row.order || row.section || '—'}|${row.title}|${row.multitracksTitle || '—'}|${row.key || '—'}|${formatBpm(
          row.bpm
        )}|${row.timeSignature || '—'}|`
      );
    }
    return lines.join('\n');
  }

  function setFloatingButtonLabel(button, labelText) {
    button.innerHTML = `
      <span class="tm-mt-btn-content">
        <span class="tm-mt-btn-icon">${MULTITRACKS_CIRCLE_ICON_SVG}</span>
        <span class="tm-mt-btn-text">${escapeHtml(labelText)}</span>
      </span>
    `;
  }

  function ensureBrandFontLoaded(doc) {
    if (!doc || !doc.head) return;
    if (doc.querySelector('link[data-tm-brand-font="code-pro"]')) return;

    const link = doc.createElement('link');
    link.rel = 'stylesheet';
    link.href = BRAND_FONT_URL;
    link.setAttribute('data-tm-brand-font', 'code-pro');
    doc.head.appendChild(link);
  }

  function cacheKeyForSong(row) {
    return `song::${normalizeText(row.title)}::${normalizeText(row.versionHint || '')}::${normalizeText(
      row.key || ''
    )}`;
  }

  function cacheFallbackKeyForSong(row) {
    return `song::${normalizeText(row.title)}::`;
  }

  async function ensureSongCacheLoaded() {
    if (songCache && typeof songCache === 'object') return songCache;
    const loaded = await gm.getValue(CACHE_KEY, {});
    songCache = loaded && typeof loaded === 'object' ? loaded : {};
    logInfo('Song cache loaded', { entries: Object.keys(songCache).length, cacheKey: CACHE_KEY });
    return songCache;
  }

  function pruneSongCacheInMemory() {
    if (!songCache || typeof songCache !== 'object') return;
    const now = Date.now();
    const keys = Object.keys(songCache);

    for (const key of keys) {
      const entry = songCache[key];
      if (!entry || typeof entry !== 'object') {
        delete songCache[key];
        continue;
      }

      if (now - Number(entry.ts || 0) > CACHE_TTL_MS) {
        delete songCache[key];
      }
    }

    const remaining = Object.keys(songCache);
    if (remaining.length <= CACHE_MAX_ENTRIES) return;

    remaining
      .sort((a, b) => {
        const left = Number(songCache[a]?.ts || 0);
        const right = Number(songCache[b]?.ts || 0);
        return left - right;
      })
      .slice(0, Math.max(0, remaining.length - CACHE_MAX_ENTRIES))
      .forEach((key) => {
        delete songCache[key];
      });
  }

  async function flushSongCache() {
    if (!songCacheDirty) return;
    pruneSongCacheInMemory();
    await gm.setValue(CACHE_KEY, songCache || {});
    songCacheDirty = false;
    logDebug('Song cache flushed', { entries: Object.keys(songCache || {}).length });
  }

  async function lookupSongCache(row) {
    const cache = await ensureSongCacheLoaded();
    const now = Date.now();

    const primaryKey = cacheKeyForSong(row);
    const fallbackKey = cacheFallbackKeyForSong(row);
    const candidates = [primaryKey, fallbackKey];

    for (const key of candidates) {
      const entry = cache[key];
      if (!entry) continue;

      if (now - Number(entry.ts || 0) > CACHE_TTL_MS) {
        delete cache[key];
        songCacheDirty = true;
        continue;
      }

      entry.hits = Number(entry.hits || 0) + 1;
      entry.ts = now;
      songCacheDirty = true;

      logInfo('Song cache hit', {
        key,
        title: row.title,
        confidence: entry.value?.confidence,
        bpm: entry.value?.bpm,
        timeSignature: entry.value?.timeSignature,
      });

      return {
        ...row,
        bpm: entry.value?.bpm ?? null,
        timeSignature: entry.value?.timeSignature ?? null,
        multitracksUrl: entry.value?.multitracksUrl ?? null,
        confidence: entry.value?.confidence || 'none',
        multitracksTitle: entry.value?.multitracksTitle || null,
        multitracksArtist: entry.value?.multitracksArtist || null,
        matchScore: entry.value?.matchScore ?? null,
        cacheHit: true,
      };
    }

    logDebug('Song cache miss', { title: row.title, primaryKey, fallbackKey });
    return null;
  }

  async function rememberSongCache(row, resolved) {
    if (!resolved || !resolved.multitracksUrl) return;

    const cache = await ensureSongCacheLoaded();
    const now = Date.now();
    const key = cacheKeyForSong(row);
    const fallbackKey = cacheFallbackKeyForSong(row);

    const value = {
      bpm: resolved.bpm ?? null,
      timeSignature: resolved.timeSignature ?? null,
      multitracksUrl: resolved.multitracksUrl ?? null,
      confidence: resolved.confidence || 'none',
      multitracksTitle: resolved.multitracksTitle || null,
      multitracksArtist: resolved.multitracksArtist || null,
      matchScore: resolved.matchScore ?? null,
    };

    cache[key] = {
      ts: now,
      hits: Number(cache[key]?.hits || 0),
      value,
    };

    // Title-level fallback is only trusted for high-confidence matches.
    if (value.confidence === 'high' && (!cache[fallbackKey] || Number(cache[fallbackKey].hits || 0) < 2)) {
      cache[fallbackKey] = {
        ts: now,
        hits: Number(cache[fallbackKey]?.hits || 0),
        value,
      };
    }

    songCacheDirty = true;
    logDebug('Song cached', {
      title: row.title,
      key,
      fallbackKey,
      bpm: value.bpm,
      timeSignature: value.timeSignature,
      confidence: value.confidence,
    });
  }

  function getByPath(obj, pathParts) {
    return pathParts.reduce((acc, key) => {
      if (acc && Object.prototype.hasOwnProperty.call(acc, key)) return acc[key];
      return null;
    }, obj);
  }

  function addUserscriptStyles() {
    const css = `
      .tm-mt-btn {
        position: fixed;
        right: 18px;
        bottom: 18px;
        z-index: 2147483000;
        border: 1px solid rgba(40, 40, 40, 0.18);
        border-radius: 999px;
        background: #fff;
        color: rgb(40, 40, 40);
        padding: 10px 14px 10px 10px;
        font: 700 22px/1.2 "Code Pro W01 Bold Lowercase", -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif;
        cursor: pointer;
        box-shadow: 0 8px 24px rgba(2, 6, 23, 0.18);
        min-height: 42px;
        display: flex;
        align-items: center;
      }
      .tm-mt-btn:hover {
        background: rgb(40, 40, 40);
        color: #fff;
      }
      .tm-mt-btn:disabled {
        opacity: 0.62;
        cursor: wait;
      }
      .tm-mt-btn-content {
        display: inline-flex;
        align-items: center;
        gap: 9px;
      }
      .tm-mt-btn-icon {
        display: inline-flex;
        width: 22px;
        height: 22px;
      }
      .tm-mt-btn-icon svg {
        width: 22px;
        height: 22px;
      }
      .tm-mt-btn-text {
        white-space: nowrap;
      }
    `;

    if (typeof GM_addStyle === 'function') {
      GM_addStyle(css);
    } else {
      const style = document.createElement('style');
      style.textContent = css;
      document.head.appendChild(style);
    }
  }

  function installFetchInterceptor() {
    if (!unsafeWindow || unsafeWindow.__TM_PCO_MT_INTERCEPTOR__) return;
    if (typeof unsafeWindow.fetch !== 'function') return;

    const originalFetch = unsafeWindow.fetch.bind(unsafeWindow);

    unsafeWindow.fetch = new Proxy(originalFetch, {
      apply: async (target, thisArg, argumentsList) => {
        const requestInput = argumentsList[0];
        const requestUrl = String(requestInput?.url || requestInput || '');
        const response = await Reflect.apply(target, thisArg, argumentsList);

        try {
          const itemsMatch = requestUrl.match(/service_types\/(\d+)\/plans\/(\d+)\/items/i);
          if (itemsMatch) {
            const [, stId, planId] = itemsMatch;
            interceptorState.serviceTypeIdByPlanId.set(planId, stId);
            logDebug('Fetch interceptor captured items request', { planId, serviceTypeId: stId, url: requestUrl });

            response
              .clone()
              .json()
              .then((data) => {
                interceptorState.itemsPayloadByPlanId.set(planId, data);
                logDebug('Fetch interceptor cached items payload', {
                  planId,
                  itemsCount: Array.isArray(data?.data) ? data.data.length : 0,
                  includedCount: Array.isArray(data?.included) ? data.included.length : 0,
                });
              })
              .catch(() => {
                /* no-op */
              });
          }
        } catch {
          // Never break site fetch.
        }

        return response;
      },
    });

    unsafeWindow.__TM_PCO_MT_INTERCEPTOR__ = true;
    logInfo('Installed fetch interceptor');
  }

  async function acquireLock(lockKey) {
    const now = Date.now();
    const existing = Number(await gm.getValue(lockKey, 0));
    if (now - existing < LOCK_TTL_MS) {
      logWarn('Lock acquisition denied', { lockKey, existing, now, ttlMs: LOCK_TTL_MS });
      return false;
    }
    await gm.setValue(lockKey, now);
    logInfo('Lock acquired', { lockKey, ts: now });
    return true;
  }

  async function refreshLock(lockKey) {
    await gm.setValue(lockKey, Date.now());
    logDebug('Lock refreshed', { lockKey });
  }

  async function releaseLock(lockKey) {
    await gm.setValue(lockKey, 0);
    logInfo('Lock released', { lockKey });
  }

  async function requestJson(url, method = 'GET') {
    logDebug('requestJson start', { method, url });
    const response = await gm.request({
      method,
      url,
      headers: { Accept: 'application/json' },
    });

    if (response.status < 200 || response.status >= 300) {
      logWarn('requestJson non-success status', { method, url, status: response.status });
      throw new Error(`HTTP ${response.status} for ${url}`);
    }
    logDebug('requestJson success', { method, url, status: response.status });

    return JSON.parse(response.responseText);
  }

  async function requestText(url, method = 'GET') {
    logDebug('requestText start', { method, url });
    const response = await gm.request({
      method,
      url,
      headers: {
        Accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
      },
    });

    if (response.status < 200 || response.status >= 300) {
      logWarn('requestText non-success status', { method, url, status: response.status });
      throw new Error(`HTTP ${response.status} for ${url}`);
    }
    logDebug('requestText success', { method, url, status: response.status, bytes: response.responseText?.length || 0 });
    return response.responseText;
  }

  async function discoverServiceTypeId(planId) {
    const intercepted = interceptorState.serviceTypeIdByPlanId.get(planId);
    if (intercepted) {
      logInfo('Using interceptor-cached service type id', { planId, serviceTypeId: intercepted });
      return intercepted;
    }

    const serviceTypes = await requestJson(`${PCO_API_BASE}/service_types?per_page=200`);
    const rows = serviceTypes?.data || [];
    logDebug('Discovering service type by HEAD probes', { planId, candidateCount: rows.length });

    for (const row of rows) {
      const stId = row?.id;
      if (!stId) continue;

      try {
        logDebug('HEAD probing service type', { planId, serviceTypeId: stId });
        // eslint-disable-next-line no-await-in-loop
        const response = await gm.request({
          method: 'HEAD',
          url: `${PCO_API_BASE}/service_types/${stId}/plans/${planId}`,
          headers: { Accept: 'application/json' },
        });

        if (response.status === 200) {
          interceptorState.serviceTypeIdByPlanId.set(planId, stId);
          logInfo('Discovered service type id', { planId, serviceTypeId: stId });
          return stId;
        }
      } catch {
        logDebug('HEAD probe failed for service type', { planId, serviceTypeId: stId });
        // Try next service type.
      }
    }

    logWarn('Service type discovery failed', { planId });
    return null;
  }

  async function fetchPlanItems(planId, serviceTypeId) {
    const intercepted = interceptorState.itemsPayloadByPlanId.get(planId);
    if (intercepted) {
      logInfo('Using interceptor-cached plan items payload', {
        planId,
        serviceTypeId,
        itemCount: Array.isArray(intercepted?.data) ? intercepted.data.length : 0,
      });
      return intercepted;
    }

    const url =
      `${PCO_API_BASE}/service_types/${serviceTypeId}/plans/${planId}/items?` +
      'per_page=200&include=song,arrangement,key,arrangement.keys';
    logInfo('Fetching plan items from API', { planId, serviceTypeId, url });
    return requestJson(url);
  }

  function resolveKeyForItem(item, arrangementObj, includedMap) {
    const itemKey = getByPath(item, ['attributes', 'key_name']);
    if (itemKey) return itemKey;

    const keyRel = getByPath(item, ['relationships', 'key', 'data']);
    if (keyRel) {
      const keyObj = includedMap.get(`${keyRel.type}:${keyRel.id}`);
      const keyName =
        getByPath(keyObj, ['attributes', 'name']) ||
        getByPath(keyObj, ['attributes', 'starting_key']) ||
        getByPath(keyObj, ['attributes', 'ending_key']);
      if (keyName) return keyName;
    }

    const arrKey =
      getByPath(arrangementObj, ['attributes', 'default_key']) ||
      getByPath(arrangementObj, ['attributes', 'chord_chart_key']);
    if (arrKey) return arrKey;

    const relKeys = getByPath(arrangementObj, ['relationships', 'keys', 'data']) || [];
    for (const rel of relKeys) {
      const keyObj = includedMap.get(`${rel.type}:${rel.id}`);
      const keyName =
        getByPath(keyObj, ['attributes', 'name']) ||
        getByPath(keyObj, ['attributes', 'starting_key']) ||
        getByPath(keyObj, ['attributes', 'ending_key']);
      if (keyName) return keyName;
    }

    return null;
  }

  function extractSongRows(itemsPayload) {
    const data = itemsPayload?.data || [];
    const included = itemsPayload?.included || [];
    const includedMap = new Map();

    for (const row of included) {
      includedMap.set(`${row.type}:${row.id}`, row);
    }

    const rows = [];
    let section = 0;

    for (const item of data) {
      const itemType = getByPath(item, ['attributes', 'item_type']);

      if (itemType === 'header') {
        section += 1;
        continue;
      }

      if (itemType !== 'song') {
        continue;
      }

      if (section === 0) section = 1;

      const title = getByPath(item, ['attributes', 'title']) || 'Untitled';
      const arrangementRel = getByPath(item, ['relationships', 'arrangement', 'data']);
      const arrangementObj = arrangementRel
        ? includedMap.get(`${arrangementRel.type}:${arrangementRel.id}`)
        : null;

      const versionHint =
        getByPath(arrangementObj, ['attributes', 'name']) || getByPath(item, ['attributes', 'description']) || '';

      const key = resolveKeyForItem(item, arrangementObj, includedMap);

      rows.push({
        id: item.id,
        section,
        title,
        key,
        versionHint,
      });

      logDebug('Extracted song row', {
        id: item.id,
        section,
        title,
        key,
        versionHint,
      });
    }

    // Normalize section groups so they are contiguous (1..N), then add stable row order.
    const sectionRemap = new Map();
    let nextSection = 1;
    rows.forEach((row, index) => {
      const rawSection = Number(row.section || 1);
      if (!sectionRemap.has(rawSection)) {
        sectionRemap.set(rawSection, nextSection);
        nextSection += 1;
      }
      row.section = sectionRemap.get(rawSection);
      row.order = index + 1;
    });

    logInfo('Extracted songs from plan payload', { count: rows.length });
    return rows;
  }

  async function searchMultiTracks(query) {
    const url = `${MT_BASE}/search/?search=${encodeURIComponent(query)}&tab=0&order=1`;
    logDebug('MultiTracks search start', { query, url });
    const html = await requestText(url);
    const prefetchMatch = html.match(/id="searchPrefetch"[^>]*value="([^"]*)"/i);
    if (!prefetchMatch) {
      logWarn('MultiTracks search had no prefetch payload', { query, url });
      return [];
    }

    const jsonText = decodeHtmlEntities(prefetchMatch[1]);

    let payload;
    try {
      payload = JSON.parse(jsonText);
    } catch {
      logWarn('Failed to parse MultiTracks prefetch JSON', { query, url });
      return [];
    }

    const songs = Array.isArray(payload?.songs) ? payload.songs : [];
    const out = [];
    const seen = new Set();

    for (const song of songs) {
      const songUrl = song?.songURL ? new URL(song.songURL, MT_BASE).toString() : null;
      const title = String(song?.title || '').trim();
      const artist = String(song?.artist || song?.artists?.[0]?.artist || '').trim();
      const album = String(song?.album || '').trim();
      if (!songUrl || !title || seen.has(songUrl)) continue;
      seen.add(songUrl);
      out.push({ songUrl, title, artist, album });
    }

    logDebug('MultiTracks search parsed candidates', { query, count: out.length });

    return out;
  }

  function parseSongMetaFromHtml(html) {
    const groupRegex =
      /<div class="song-banner--meta-list--group"[^>]*>[\s\S]*?<dt[^>]*>([\s\S]*?)<\/dt>[\s\S]*?<dd[^>]*>([\s\S]*?)<\/dd>[\s\S]*?<\/div>/gi;

    let match;
    let bpm = null;
    let timeSignature = null;

    while ((match = groupRegex.exec(html)) !== null) {
      const term = stripTags(match[1]).replace(/:$/, '').toLowerCase();
      const desc = stripTags(match[2]);
      if (!bpm && term.startsWith('bpm')) bpm = parseBpm(desc);
      if (!timeSignature && term.startsWith('time sig')) timeSignature = parseTimeSignature(desc);
    }

    return { bpm, timeSignature };
  }

  async function fetchSongMeta(songUrl) {
    logDebug('Fetching MultiTracks song page for metadata', { songUrl });
    const html = await requestText(songUrl);
    const parsed = parseSongMetaFromHtml(html);
    logDebug('Parsed MultiTracks metadata', { songUrl, bpm: parsed.bpm, timeSignature: parsed.timeSignature });
    return parsed;
  }

  function isLikelyVersionHint(hint) {
    const text = normalizeText(hint);
    if (!text) return false;

    return /\blive\b|\bstudio\b|\bflow\b|\bacoustic\b|\binstrumental\b|\breprise\b|\bspontaneous\b|\bmix\b|\bedit\b|\bversion\b|\barrangement\b|\bradio\b|\bdemo\b|\bloop\b|\bextended\b/.test(
      text
    );
  }

  function computeExtraTitleTokens(target, candidate) {
    const targetTokens = new Set(String(target || '').split(' ').filter(Boolean));
    const candidateTokens = new Set(String(candidate || '').split(' ').filter(Boolean));
    return [...candidateTokens].filter((token) => !targetTokens.has(token));
  }

  function rankCandidates(row, candidates) {
    const targetTitle = normalizeText(row.title);
    const targetCore = normalizeText(stripParenthetical(row.title));
    const hint = normalizeText(row.versionHint);
    const hasHint = Boolean(hint);
    const hintIsVersion = isLikelyVersionHint(hint);
    const expectedArtist = !hintIsVersion && hasHint ? hint : '';
    const wantsLive = /\blive\b/.test(hint);
    const wantsStudio = /\bstudio\b/.test(hint);

    return candidates
      .map((candidate, index) => {
        const title = normalizeText(candidate.title);
        const titleCore = normalizeText(stripParenthetical(candidate.title));
        const artist = normalizeText(candidate.artist);
        const meta = normalizeText(`${candidate.title} ${candidate.artist} ${candidate.album}`);
        const extraTitleTokens = computeExtraTitleTokens(targetCore || targetTitle, titleCore || title);

        let score = 0;
        if (title === targetTitle) score += 260;
        if (titleCore === targetCore) score += 120;
        if (title.includes(targetTitle) || targetTitle.includes(title)) score += 32;
        if (titleCore.includes(targetCore) || targetCore.includes(titleCore)) score += 24;

        score += Math.round(jaccardSimilarity(targetCore || targetTitle, titleCore || title) * 95);

        if (expectedArtist) {
          if (artist === expectedArtist) score += 260;
          else if (artist.includes(expectedArtist) || expectedArtist.includes(artist)) score += 170;
          else if (meta.includes(expectedArtist)) score += 120;
          else score -= 120;
        } else if (hasHint) {
          if (meta.includes(hint)) score += 70;
          else score -= 15;
        }

        if (wantsLive) score += meta.includes('live') ? 22 : -8;
        if (wantsStudio) score += meta.includes('studio') ? 18 : -8;

        if (!hint.includes('flow') && /\bflow\b/.test(meta)) score -= 120;
        if (!hint.includes('instrumental') && /\binstrumental\b/.test(meta)) score -= 55;
        if (!hint.includes('reprise') && /\breprise\b/.test(meta)) score -= 45;

        score -= Math.max(0, extraTitleTokens.length - (wantsLive ? 1 : 0)) * 16;

        score += Math.max(0, 12 - index);

        return { ...candidate, score };
      })
      .sort((a, b) => b.score - a.score);
  }

  function confidenceFromScores(best, second, completeness) {
    if (!best) return 'none';
    const gap = best.score - (second?.score || 0);
    if (best.score >= 320 && completeness === 2 && gap >= 80) return 'high';
    if (best.score >= 200 && completeness >= 1) return 'medium';
    return 'low';
  }

  async function resolveSongOnMultiTracks(row) {
    const cached = await lookupSongCache(row);
    if (cached) return cached;

    const queries = [row.title, `${row.title} ${row.versionHint || ''}`]
      .map((q) => q.trim())
      .filter(Boolean);

    logInfo('Resolving song on MultiTracks', {
      id: row.id,
      title: row.title,
      key: row.key,
      versionHint: row.versionHint,
      queries,
    });

    const seen = new Set();
    const candidates = [];

    for (const query of queries) {
      try {
        // eslint-disable-next-line no-await-in-loop
        const found = await searchMultiTracks(query);
        logDebug('MultiTracks query returned', { title: row.title, query, candidates: found.length });
        for (const candidate of found) {
          if (seen.has(candidate.songUrl)) continue;
          seen.add(candidate.songUrl);
          candidates.push(candidate);
        }
      } catch {
        logWarn('MultiTracks query failed', { title: row.title, query });
        // Keep going with other queries.
      }
    }

    if (!candidates.length) {
      logWarn('No MultiTracks candidates found', { title: row.title, queries });
      return {
        ...row,
        bpm: null,
        timeSignature: null,
        multitracksUrl: null,
        confidence: 'none',
      };
    }

    const ranked = rankCandidates(row, candidates);
    logDebug('Ranked candidates', {
      title: row.title,
      top: ranked.slice(0, 5).map((candidate) => ({
        title: candidate.title,
        artist: candidate.artist,
        score: candidate.score,
        songUrl: candidate.songUrl,
      })),
    });
    const inspected = [];

    for (const candidate of ranked.slice(0, MAX_CANDIDATES)) {
      let meta = { bpm: null, timeSignature: null };
      try {
        // eslint-disable-next-line no-await-in-loop
        meta = await fetchSongMeta(candidate.songUrl);
      } catch {
        logWarn('Failed to fetch metadata for candidate', {
          title: row.title,
          candidateTitle: candidate.title,
          candidateArtist: candidate.artist,
          songUrl: candidate.songUrl,
        });
        // Keep candidate even if metadata fetch failed.
      }

      inspected.push({
        ...candidate,
        bpm: meta.bpm,
        timeSignature: meta.timeSignature,
        completeness: Number(Boolean(meta.bpm)) + Number(Boolean(meta.timeSignature)),
      });
    }

    inspected.sort((a, b) => {
      if (b.completeness !== a.completeness) return b.completeness - a.completeness;
      return b.score - a.score;
    });

    const best = inspected[0] || null;
    const second = inspected[1] || null;

    logInfo('Resolved song result', {
      title: row.title,
      selected: best
        ? {
            songUrl: best.songUrl,
            title: best.title,
            artist: best.artist,
            score: best.score,
            bpm: best.bpm,
            timeSignature: best.timeSignature,
          }
        : null,
      runnerUp: second
        ? {
            songUrl: second.songUrl,
            title: second.title,
            artist: second.artist,
            score: second.score,
            bpm: second.bpm,
            timeSignature: second.timeSignature,
          }
        : null,
    });

    return {
      ...row,
      bpm: best?.bpm ?? null,
      timeSignature: best?.timeSignature ?? null,
      multitracksUrl: best?.songUrl ?? null,
      multitracksTitle: best?.title ?? null,
      multitracksArtist: best?.artist ?? null,
      matchScore: best?.score ?? null,
      confidence: confidenceFromScores(best, second, best?.completeness || 0),
      debugTop: inspected.slice(0, 3),
    };
  }

  async function processQueueWithBackpressure({ queueKey, resultsKey, initialTasks, onProgress }) {
    const queue = Array.isArray(initialTasks) ? initialTasks.map((task) => ({ ...task })) : [];
    const results = {};
    let completed = 0;

    logInfo('Queue processor boot', {
      queueKey,
      resultsKey,
      initialTasks: queue.length,
      concurrency: CONCURRENCY,
      maxRetries: MAX_RETRIES_PER_SONG,
    });

    await gm.setValues({ [queueKey]: queue, [resultsKey]: results });

    async function persistSnapshot(reason) {
      await gm.setValues({ [queueKey]: queue, [resultsKey]: results });
      logDebug('Persisted queue snapshot', {
        reason,
        queueRemaining: queue.length,
        done: completed,
        storedResults: Object.keys(results).length,
      });
    }

    async function worker(workerId) {
      logDebug('Worker started', { workerId });

      while (true) {
        const task = queue.shift();
        if (!task) {
          logDebug('Worker idle-exit (queue empty)', { workerId });
          return;
        }

        const attempt = Number(task.__attempt || 0) + 1;
        task.__attempt = attempt;

        logDebug('Worker dequeued task', {
          workerId,
          taskId: task.id,
          title: task.title,
          attempt,
          queueRemainingAfterDequeue: queue.length,
        });

        await persistSnapshot(`dequeue:worker-${workerId}:task-${task.id}`);

        try {
          const resolved = await resolveSongOnMultiTracks(task);
          await rememberSongCache(task, resolved);
          results[task.id] = resolved;
          completed += 1;

          await persistSnapshot(`resolved:worker-${workerId}:task-${task.id}`);
          onProgress?.(resolved, null, queue.length, completed);

          logDebug('Worker completed task', {
            workerId,
            taskId: task.id,
            title: task.title,
            completed,
            queueRemaining: queue.length,
          });
        } catch (error) {
          logWarn('Worker task failed', {
            workerId,
            taskId: task.id,
            title: task.title,
            attempt,
            error: error?.message || String(error),
          });

          if (attempt < MAX_RETRIES_PER_SONG) {
            queue.push(task);
            await persistSnapshot(`retry-enqueue:worker-${workerId}:task-${task.id}`);
            onProgress?.(task, error, queue.length, completed);

            // Recursive-timeout style backpressure; no setInterval overlap.
            // eslint-disable-next-line no-await-in-loop
            await sleep(300 * attempt);
            continue;
          }

          const failed = {
            ...task,
            bpm: null,
            timeSignature: null,
            multitracksUrl: null,
            confidence: 'none',
            error: error?.message || String(error),
          };

          results[task.id] = failed;
          completed += 1;

          await persistSnapshot(`failed-final:worker-${workerId}:task-${task.id}`);
          onProgress?.(failed, error, queue.length, completed);

          logError('Worker gave up after max retries', {
            workerId,
            taskId: task.id,
            title: task.title,
            completed,
            queueRemaining: queue.length,
          });
        }

        // Recursive timeout tick to keep pressure bounded.
        // eslint-disable-next-line no-await-in-loop
        await sleep(100);
      }
    }

    const workerCount = Math.max(1, Math.min(CONCURRENCY, queue.length || 1));
    const workers = Array.from({ length: workerCount }, (_, idx) => worker(idx + 1));
    await Promise.all(workers);

    logInfo('Queue processor finished', {
      totalResults: Object.keys(results).length,
      completed,
      queueRemaining: queue.length,
    });

    await gm.setValues({ [queueKey]: [], [resultsKey]: results });
    return results;
  }

  function createReportPage(win) {
    const doc = win.document;
    doc.open();
    doc.write(`<!doctype html>
      <html>
      <head>
        <meta charset="utf-8" />
        <title>PCO MultiTracks Table</title>
        <style>
          body { font-family: Inter, -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, sans-serif; margin: 0; background: #0f172a; color: #e2e8f0; }
          .wrap { max-width: 1120px; margin: 0 auto; padding: 24px; }
          .card { background: rgba(15,23,42,.7); border: 1px solid rgba(148,163,184,.25); border-radius: 14px; padding: 18px; margin-bottom: 16px; backdrop-filter: blur(3px); }
          h1 { margin: 0 0 10px; font-size: 22px; }
          .meta { opacity: .9; font-size: 14px; margin: 6px 0; }
          .actions { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 10px; }
          button { border: 0; border-radius: 8px; padding: 8px 12px; font-weight: 600; cursor: pointer; background: #0ea5e9; color: #fff; }
          button.alt { background: #334155; }
          .status { font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; font-size: 12px; white-space: pre-wrap; line-height: 1.45; max-height: 220px; overflow: auto; background: #020617; border-radius: 8px; padding: 10px; }
          .progress-wrap { margin: 10px 0 12px; }
          .progress-meta { font-size: 12px; opacity: .9; margin-bottom: 6px; display: flex; justify-content: space-between; }
          .progress-bar { width: 100%; height: 8px; border-radius: 999px; background: #1e293b; overflow: hidden; }
          .progress-fill { height: 100%; width: 0%; background: linear-gradient(90deg, #0891b2, #22d3ee); transition: width .15s ease; }
          table { width: 100%; border-collapse: collapse; background: #020617; border-radius: 10px; overflow: hidden; }
          th, td { padding: 10px 10px; border-bottom: 1px solid #1e293b; text-align: left; }
          thead th { background: #111827; position: sticky; top: 0; }
          td.num, th.num { text-align: right; }
          td.center, th.center { text-align: center; }
          a { color: #38bdf8; text-decoration: none; }
          .pill { display: inline-block; border-radius: 999px; font-size: 11px; padding: 2px 8px; background: #1e293b; }
          .pill.high { background: #14532d; }
          .pill.medium { background: #78350f; }
          .pill.low { background: #7f1d1d; }
        </style>
      </head>
      <body>
        <div class="wrap">
          <div class="card">
            <h1>PCO → MultiTracks Table</h1>
            <div id="meta" class="meta">Starting…</div>
            <div class="actions">
              <button id="copy-md">Copy Markdown</button>
              <button id="copy-json" class="alt">Copy JSON</button>
              <button id="toggle-debug" class="alt">Verbose: ${DEBUG_LOGGING ? 'ON' : 'OFF'}</button>
            </div>
            <div class="progress-wrap">
              <div class="progress-meta">
                <span id="progress-label">Starting…</span>
                <span id="progress-count">0/0</span>
              </div>
              <div class="progress-bar"><div id="progress-fill" class="progress-fill"></div></div>
            </div>
            <div id="status" class="status">Booting…</div>
          </div>
          <div class="card" id="table-wrap"></div>
        </div>
      </body>
      </html>`);
    doc.close();

    return {
      setMeta(text) {
        const el = doc.getElementById('meta');
        if (el) el.textContent = text;
      },
      setProgress(done, total, label) {
        const safeDone = Math.max(0, Number(done || 0));
        const safeTotal = Math.max(0, Number(total || 0));
        const percent = safeTotal ? Math.min(100, Math.round((safeDone / safeTotal) * 100)) : 0;

        const countEl = doc.getElementById('progress-count');
        if (countEl) countEl.textContent = `${safeDone}/${safeTotal}`;

        const labelEl = doc.getElementById('progress-label');
        if (labelEl && label) labelEl.textContent = label;

        const fillEl = doc.getElementById('progress-fill');
        if (fillEl) fillEl.style.width = `${percent}%`;
      },
      appendStatus(text) {
        const el = doc.getElementById('status');
        if (!el) return;
        el.textContent += `\n${text}`;
        el.scrollTop = el.scrollHeight;
      },
      renderTable(rows) {
        const wrap = doc.getElementById('table-wrap');
        if (!wrap) return;

        const body = rows
          .map((row) => {
            const titleCell = row.multitracksUrl
              ? `<a href="${escapeHtml(row.multitracksUrl)}" target="_blank" rel="noopener noreferrer">${escapeHtml(
                  row.title
                )}</a>`
              : escapeHtml(row.title);

            return `<tr>
              <td class="num">${escapeHtml(row.order || row.section || '—')}</td>
              <td>${titleCell}</td>
              <td>${escapeHtml(row.multitracksTitle || '—')}</td>
              <td class="center">${escapeHtml(row.key || '—')}</td>
              <td class="num">${escapeHtml(formatBpm(row.bpm))}</td>
              <td class="center">${escapeHtml(row.timeSignature || '—')}</td>
            </tr>`;
          })
          .join('');

        wrap.innerHTML = `
          <table>
            <thead>
              <tr>
                <th class="num">#</th>
                <th>Song</th>
                <th>Version (MT)</th>
                <th class="center">Key</th>
                <th class="num">Original BPM</th>
                <th class="center">Time Sig</th>
              </tr>
            </thead>
            <tbody>${body}</tbody>
          </table>`;
      },
      onCopyMarkdown(getText) {
        const btn = doc.getElementById('copy-md');
        if (!btn) return;
        btn.onclick = async () => {
          const text = getText();
          await win.navigator.clipboard.writeText(text);
          btn.textContent = 'Copied Markdown';
          setTimeout(() => {
            btn.textContent = 'Copy Markdown';
          }, 1500);
        };
      },
      onCopyJson(getText) {
        const btn = doc.getElementById('copy-json');
        if (!btn) return;
        btn.onclick = async () => {
          const text = getText();
          await win.navigator.clipboard.writeText(text);
          btn.textContent = 'Copied JSON';
          setTimeout(() => {
            btn.textContent = 'Copy JSON';
          }, 1500);
        };
      },
      onToggleDebug() {
        const btn = doc.getElementById('toggle-debug');
        if (!btn) return;
        btn.onclick = async () => {
          const next = !DEBUG_LOGGING;
          await gm.setValue('tm_pco_mt_debug', next);
          btn.textContent = `Verbose: ${next ? 'ON' : 'OFF'} (next run)`;
        };
      },
    };
  }

  async function runOneClickFlow(buttonEl, reportWindow) {
    logInfo('One-click flow started', { href: location.href });
    const planId = parsePlanIdFromUrl(location.href);
    if (!planId) {
      if (reportWindow && !reportWindow.closed) reportWindow.close();
      alert('Could not detect a numeric plan ID from this URL.');
      return;
    }

    const lockKey = `tm_pco_mt_lock_${planId}`;
    const lockOk = await acquireLock(lockKey);
    if (!lockOk) {
      if (reportWindow && !reportWindow.closed) reportWindow.close();
      alert('Another tab appears to be running this plan already. Try again in a few seconds.');
      return;
    }

    buttonEl.disabled = true;
    setFloatingButtonLabel(buttonEl, 'MultiTracks BPM…');

    if (!reportWindow) {
      await releaseLock(lockKey);
      buttonEl.disabled = false;
      setFloatingButtonLabel(buttonEl, 'MultiTracks BPM');
      alert('Popup blocked. Allow popups for this site and try again.');
      return;
    }

    const report = createReportPage(reportWindow);
    const startedAt = Date.now();
    report.onToggleDebug();
    report.setProgress(0, 0, 'Preparing…');

    const queueKey = `tm_pco_mt_queue_${planId}`;
    const resultsKey = `tm_pco_mt_results_${planId}`;

    const refreshTimer = setInterval(() => {
      refreshLock(lockKey).catch(() => {
        /* no-op */
      });
    }, LOCK_REFRESH_MS);

    try {
      report.setMeta(`Plan ${planId} — Loading from PCO…`);
      report.appendStatus('[1/4] Discovering service type…');
      report.setProgress(0, 4, 'Discovering service type');
      logInfo('Step 1/4: discover service type', { planId });

      const serviceTypeId = await discoverServiceTypeId(planId);
      if (!serviceTypeId) {
        throw new Error('Could not discover PCO service type for this plan.');
      }

      report.appendStatus(`[2/4] Fetching plan items (service type ${serviceTypeId})…`);
      report.setProgress(1, 4, `Fetching plan items (service type ${serviceTypeId})`);
      logInfo('Step 2/4: fetch plan items', { planId, serviceTypeId });
      const itemsPayload = await fetchPlanItems(planId, serviceTypeId);
      const songs = extractSongRows(itemsPayload);
      if (!songs.length) throw new Error('No song rows found in this plan.');

      report.appendStatus(`[3/4] Queueing ${songs.length} songs for MultiTracks resolution…`);
      report.setProgress(0, songs.length, `Resolving ${songs.length} songs on MultiTracks`);
      logInfo('Step 3/4: queue songs', {
        planId,
        songs: songs.map((song) => ({ id: song.id, section: song.section, title: song.title, key: song.key })),
      });
      await gm.setValues({
        [queueKey]: songs,
        [resultsKey]: {},
      });

      report.setMeta(`Plan ${planId} — Resolving ${songs.length} songs from MultiTracks…`);
      const resultsMap = await processQueueWithBackpressure({
        queueKey,
        resultsKey,
        initialTasks: songs,
        onProgress: (row, error, queueRemaining, doneCount) => {
          if (error) {
            report.appendStatus(
              `[retry] ${row.title || row.id} failed (${error.message || error}). queue=${queueRemaining}`
            );
            report.setProgress(doneCount, songs.length, `Retrying ${row.title || row.id}`);
            logWarn('Progress callback error', {
              planId,
              rowId: row.id,
              title: row.title,
              queueRemaining,
              doneCount,
              error: error?.message || String(error),
            });
            return;
          }

          const icon = row.multitracksUrl ? (row.confidence === 'high' ? '✓' : '~') : '✗';
          report.appendStatus(
            `[${doneCount}] ${icon} ${row.title} | key=${row.key || '—'} | bpm=${formatBpm(
              row.bpm
            )} | time=${row.timeSignature || '—'} | queue=${queueRemaining}`
          );
          report.setProgress(doneCount, songs.length, `Resolved ${doneCount}/${songs.length}`);
          logInfo('Progress callback success', {
            planId,
            rowId: row.id,
            title: row.title,
            key: row.key,
            bpm: row.bpm,
            timeSignature: row.timeSignature,
            multitracksUrl: row.multitracksUrl,
            confidence: row.confidence,
            queueRemaining,
            doneCount,
          });
        },
      });

      const ordered = songs.map((song) => resultsMap[song.id] || { ...song, bpm: null, timeSignature: null });
      const markdown = buildMarkdownTable(ordered);
      logInfo('Step 4/4: render completed rows', {
        planId,
        orderedRows: ordered.map((row) => ({
          section: row.section,
          title: row.title,
          key: row.key,
          bpm: row.bpm,
          timeSignature: row.timeSignature,
          confidence: row.confidence,
          multitracksUrl: row.multitracksUrl,
        })),
      });

      report.setMeta(
        `Plan ${planId} — Done in ${Math.round((Date.now() - startedAt) / 1000)}s | ${ordered.length} rows`
      );
      report.appendStatus('[4/4] Rendering table… done.');
      report.setProgress(ordered.length, ordered.length, 'Done');
      report.renderTable(ordered);
      report.onCopyMarkdown(() => markdown);
      report.onCopyJson(() => JSON.stringify(ordered, null, 2));

      // Keep data in storage so you can inspect/reuse from another tab quickly.
      await gm.setValues({
        [`tm_pco_mt_last_plan_${planId}`]: {
          generatedAt: new Date().toISOString(),
          planId,
          rows: ordered,
          markdown,
        },
      });
    } catch (error) {
      report.setMeta(`Plan ${planId} — Failed`);
      report.appendStatus(`[error] ${error.message || String(error)}`);
      logError('One-click flow failed', { planId, error: error?.message || String(error), stack: error?.stack });
    } finally {
      await flushSongCache().catch((error) => {
        logWarn('Failed to flush song cache', { error: error?.message || String(error) });
      });
      clearInterval(refreshTimer);
      await releaseLock(lockKey);
      buttonEl.disabled = false;
      setFloatingButtonLabel(buttonEl, 'MultiTracks BPM');
      logInfo('One-click flow finished', { planId });
    }
  }

  function mountButton() {
    if (document.querySelector('.tm-mt-btn')) return;

    const button = document.createElement('button');
    button.className = 'tm-mt-btn';
    button.type = 'button';
    setFloatingButtonLabel(button, 'MultiTracks BPM');
    button.title = 'One-click: build MultiTracks BPM table for this plan';
    button.addEventListener('click', () => {
      // Must open the tab synchronously in the click handler to preserve user activation.
      const reportWindow = window.open('about:blank', '_blank');
      if (!reportWindow) {
        alert('Popup blocked. Allow popups for this site and try again.');
        return;
      }
      runOneClickFlow(button, reportWindow).catch((error) => {
        logError('Unhandled flow error', { error: error?.message || String(error), stack: error?.stack });
        if (!reportWindow.closed) reportWindow.close();
      });
    });

    document.body.appendChild(button);
  }

  function init() {
    logInfo('Userscript init start', { href: location.href });
    ensureBrandFontLoaded(document);
    installFetchInterceptor();
    addUserscriptStyles();
    mountButton();
    window.addEventListener('load', mountButton, { once: true });
    logInfo('Userscript init complete');
  }

  init();
})();
