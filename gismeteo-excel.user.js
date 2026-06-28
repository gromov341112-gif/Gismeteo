// ==UserScript==
// @name         Gismeteo Precipitation
// @namespace    gismeteo-excel
// @version      2.1
// @description  Export Gismeteo 10-day precipitation, wind, and gust forecasts to a styled Excel report with daily charts and filtered alert lists.
// @author       HARIBB
// @match        https://www.gismeteo.ru/*
// @grant        none
// @icon         https://raw.githubusercontent.com/gromov341112-gif/Gismeteo/main/assets/icon.svg
// @homepageURL  https://gromov341112-gif.github.io/Gismeteo/
// @supportURL   https://github.com/gromov341112-gif/Gismeteo/issues
// @downloadURL  https://raw.githubusercontent.com/gromov341112-gif/Gismeteo/main/gismeteo-excel.user.js
// @updateURL    https://raw.githubusercontent.com/gromov341112-gif/Gismeteo/main/gismeteo-excel.user.js
// @require      https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js
// @require      https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js
// @require      https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js
// ==/UserScript==

(function () {
  'use strict';

  const STORAGE_KEY = 'gm_city_list_v45';
  const FIXED_LIST_STORAGE_KEY = 'gm_fixed_osadki_gust_city_list_v1';
  const BASE = 'https://www.gismeteo.ru';
  const APP_VERSION = '2.1';
  const APP_TITLE = `Gismeteo Precipitation v${APP_VERSION}`;
  const PANEL_COLLAPSED_WIDTH = 'max-content';
  const PANEL_EXPANDED_WIDTH = '336px';
  const LOCATION_WORDS_RE = /^(?:погода\s+)?(?:в|во|на|для)\s+/i;
  const FORECAST_TAIL_RE = /\s+(?:на\s+(?:10\s+дней|3\s+дня|2\s+недели|месяц|неделю|выходные)|сегодня|завтра).*$/i;
  const BAD_LOCATION_RE = /аэропорт|airport|аэродром|aeroport|авиабаза|внуково|шереметьево|домодедово|спиченково|остафьево/i;
  const WEATHER_TEXT_RE = /\b(?:ясно|облачно|пасмурно|дождь|дождя|дождем|дождём|дожди|гроз[а-я]*|снег[а-я]*|туман[а-я]*|град|ливень|ливни|морось|переменная)\b/i;
  let collapsed = true;
  let collectionLocked = false;

  createPanel();

  window.addEventListener('beforeunload', event => {
    if (!collectionLocked) return;

    event.preventDefault();
    event.returnValue = '';
  });

  window.addEventListener('keydown', event => {
    if (!collectionLocked) return;

    const isReloadKey = event.key === 'F5' || ((event.ctrlKey || event.metaKey) && event.key?.toLowerCase() === 'r');

    if (isReloadKey) {
      event.preventDefault();
      event.stopPropagation();
    }
  }, true);

  function createPanel() {
    const style = document.createElement('style');
    style.textContent = `
      #gmRun {
        transition: transform .16s ease, background-color .16s ease;
      }
      #gmRun:hover {
        background: #22c55e !important;
        transform: translateY(-1px);
      }
      #gmRun:active {
        transform: translateY(0) scale(.97);
      }
    `;
    document.head.appendChild(style);

    const box = document.createElement('div');
    box.id = 'gmBox';
    box.style = `
      position: fixed; right: 18px; bottom: 18px; width: ${PANEL_COLLAPSED_WIDTH}; max-width: calc(100vw - 36px);
      background: #111827; z-index: 999999; border-radius: 10px;
      border: 1px solid #263244;
      box-shadow: 0 14px 34px rgba(0,0,0,.38);
      font-family: Arial, sans-serif; color: #e5e7eb; overflow: hidden;
    `;

    box.innerHTML = `
      <div id="gmHeader" style="
        background:#0f172a;color:#f8fafc;padding:8px 10px;cursor:pointer;
        border-bottom:1px solid #263244;
        font-size:13px;font-weight:bold;
      ">
        <span style="white-space:nowrap;">${APP_TITLE}</span>
      </div>

      <div id="gmBody" style="display:none;padding:8px;">
        <div style="margin:0 0 4px;font-size:11px;color:#cbd5e1;font-weight:bold;">Отчет №1</div>
        <textarea id="gmCities" placeholder="Осадки по городам" style="
          width:100%;height:86px;box-sizing:border-box;resize:vertical;
          padding:7px;font-size:12px;background:#020617;color:#e5e7eb;
          border:1px solid #334155;border-radius:7px;outline:none;
        "></textarea>

        <div style="margin:8px 0 4px;font-size:11px;color:#cbd5e1;font-weight:bold;">Отчет №2</div>
        <textarea id="gmFixedCities" placeholder="Порывы по фиксированным городам" style="
          width:100%;height:86px;box-sizing:border-box;resize:vertical;
          padding:7px;font-size:12px;background:#020617;color:#e5e7eb;
          border:1px solid #334155;border-radius:7px;outline:none;
        "></textarea>

        <button id="gmRun" style="
          display:flex;align-items:center;justify-content:center;
          margin:8px auto 0;width:86px;min-height:34px;padding:5px 10px;background:#16a34a;color:#fff;
          border:none;border-radius:999px;cursor:pointer;font-size:12px;text-align:center;
        ">Собрать</button>

        <div id="gmStatus" style="margin-top:7px;font-size:11px;color:#94a3b8;line-height:1.3;"></div>
      </div>
    `;

    document.body.appendChild(box);

    const textarea = document.getElementById('gmCities');
    const fixedTextarea = document.getElementById('gmFixedCities');
    textarea.value = localStorage.getItem(STORAGE_KEY) || localStorage.getItem('gm_city_list_v44') || '';
    fixedTextarea.value = localStorage.getItem(FIXED_LIST_STORAGE_KEY) || '';

    document.getElementById('gmHeader').onclick = () => {
      collapsed = !collapsed;
      document.getElementById('gmBody').style.display = collapsed ? 'none' : 'block';
      box.style.width = collapsed ? PANEL_COLLAPSED_WIDTH : PANEL_EXPANDED_WIDTH;
    };

    document.getElementById('gmRun').onclick = async () => {
      const cities = textarea.value.split('\n').map(x => x.trim()).filter(Boolean);
      const fixedCities = dedupeCityList(fixedTextarea.value.split('\n').map(x => x.trim()).filter(Boolean));
      const collectionCities = mergeCityLists(cities, fixedCities);
      localStorage.setItem(STORAGE_KEY, textarea.value);
      localStorage.setItem(FIXED_LIST_STORAGE_KEY, fixedTextarea.value);

      if (!collectionCities.length) {
        alert('Вставь список городов');
        return;
      }

      await run(collectionCities, fixedCities, cities);
    };
  }

  function mergeCityLists(...lists) {
    return dedupeCityList(lists.flat());
  }

  function dedupeCityList(cities) {
    const seen = new Set();
    const result = [];

    for (const city of cities) {
      const cleanCity = cleanCityName(city);
      const key = normalizeLocation(cleanCity);

      if (!key || seen.has(key)) continue;

      seen.add(key);
      result.push(cleanCity);
    }

    return result;
  }

  async function run(cities, fixedCities = [], reportCities = cities) {
    const btn = document.getElementById('gmRun');
    const status = document.getElementById('gmStatus');

    collectionLocked = true;
    btn.disabled = true;
    btn.textContent = 'Сбор...';

    const blocks = [];

    try {
      for (let i = 0; i < cities.length; i++) {
        const rawCity = cities[i];

        try {
          status.textContent = `${i + 1}/${cities.length}: поиск ${rawCity}`;

          const resolved = await loadCityForecast(rawCity, found => {
            status.textContent = `${i + 1}/${cities.length}: прогноз ${found.name}`;
          });
          const { found, forecastUrl, parsed } = resolved;
          const rainChart = makeRainChartImage(found.name || rawCity, parsed);
          const windChart = makeWindChartImage(found.name || rawCity, parsed);

          blocks.push({
            city: found.name || rawCity,
            url: forecastUrl,
            days: parsed.days,
            rain: parsed.rain,
            wind: parsed.wind,
            gust: parsed.gust,
            weather: parsed.weather,
            rainChart,
            windChart,
            status: 'OK'
          });
        } catch (e) {
          console.error(e);

          blocks.push({
            city: rawCity,
            url: '',
            days: buildDates(),
            rain: Array(10).fill(''),
            wind: Array(10).fill(''),
            gust: Array(10).fill(''),
            weather: Array(10).fill(''),
            rainChart: null,
            windChart: null,
            status: 'FAIL'
          });
        }

        await sleep(1200);
      }

      status.textContent = 'Формирую Excel...';
      await exportExcel(blocks, fixedCities, reportCities);

      status.textContent = `Готово: ${cities.length}`;
    } finally {
      collectionLocked = false;
      btn.disabled = false;
      btn.textContent = 'Собрать';
    }
  }

  async function loadCityForecast(city, onAttempt) {
    const candidates = await findCityCandidates(city);
    let lastError = null;

    for (const found of candidates) {
      const forecastUrl = make10DaysUrl(found.url);

      try {
        if (typeof onAttempt === 'function') onAttempt(found);

        const parsed = await loadForecastWithRetries(forecastUrl, 3, found);

        return { found, forecastUrl, parsed };
      } catch (e) {
        lastError = e;
      }
    }

    throw lastError || new Error(`город не найден: ${city}`);
  }

  async function findCityUrl(city) {
    const candidates = await findCityCandidates(city);

    if (!candidates.length) throw new Error(`город не найден: ${city}`);

    return candidates[0];
  }

  async function findCityCandidates(city) {
    let lastError = null;

    for (let attempt = 1; attempt <= 2; attempt++) {
      try {
        return await findCityCandidatesOnce(city);
      } catch (e) {
        lastError = e;

        if (attempt < 2) await sleep(900);
      }
    }

    throw lastError || new Error(`город не найден: ${city}`);
  }

  async function findCityCandidatesOnce(city) {
    const searchUrl = `${BASE}/search/${encodeURIComponent(city)}/`;
    const loaded = await loadPageWithIframe(searchUrl, 9000);

    try {
      const doc = loaded.doc;
      const query = normalizeLocation(city);
      const allCandidates = [...doc.querySelectorAll('a[href*="/weather-"]')]
        .map(a => buildCityCandidate(a, query))
        .filter(Boolean)
        .filter(x =>
          x.href.includes('/weather-') &&
          !x.href.includes('/maps/') &&
          !isBadLocation(x)
        );

      const candidates = uniqueCityCandidates(allCandidates
        .filter(x => isExactCityMatch(query, x.normalizedLocation))
        .sort((a, b) => b.score - a.score));

      if (!candidates.length) throw new Error(`город не найден: ${city}`);

      return candidates;
    } finally {
      loaded.iframe.remove();
    }
  }

  function uniqueCityCandidates(candidates) {
    const seen = new Set();

    return candidates.filter(candidate => {
      const key = make10DaysUrl(candidate.url);

      if (seen.has(key)) return false;

      seen.add(key);
      return true;
    });
  }

  function buildCityCandidate(a, query) {
    const href = a.getAttribute('href') || '';
    const rawText = a.innerText || a.textContent || '';
    const contextText = getCandidateContextText(a);
    const name = cleanCityName(rawText);
    const location = extractLocationName(name);
    const normalizedLocation = normalizeLocation(location || name);

    if (!href || !normalizedLocation) return null;

    return {
      name: location || name,
      href,
      url: href.startsWith('http') ? href : BASE + href,
      normalizedLocation,
      contextText,
      score: scoreCity(query, normalizedLocation, href, `${name} ${contextText}`)
    };
  }

  function getCandidateContextText(a) {
    const ownText = cleanCityName(a.innerText || a.textContent || '');
    const maxContextLength = Math.max(ownText.length + 240, 300);

    for (let node = a.parentElement; node && node !== document.body; node = node.parentElement) {
      const text = cleanCityName(node.innerText || node.textContent || '');

      if (text && text.length > ownText.length && text.length <= maxContextLength) {
        return text;
      }
    }

    return '';
  }

  function cleanCityName(text) {
    return String(text)
      .replace(/\u00A0/g, ' ')
      .replace(/\s+/g, ' ')
      .replace(/\s*,?\s*Россия.*$/i, '')
      .trim();
  }

  function extractLocationName(text) {
    return cleanCityName(text)
      .replace(LOCATION_WORDS_RE, '')
      .replace(FORECAST_TAIL_RE, '')
      .trim();
  }

  function normalizeText(text) {
    return String(text)
      .toLowerCase()
      .replace(/ё/g, 'е')
      .replace(/[^\p{L}\p{N}\s-]/gu, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function normalizeLocation(text) {
    return normalizeText(extractLocationName(text));
  }

  function isBadLocation(candidate) {
    return BAD_LOCATION_RE.test(`${candidate.name} ${candidate.href} ${candidate.normalizedLocation} ${candidate.contextText}`);
  }

  function isExactCityMatch(query, candidate) {
    if (!query || !candidate) return false;
    if (candidate === query) return true;

    return candidate
      .split(/[\s-]+/)
      .filter(Boolean)
      .join(' ') === query;
  }

  function scoreCity(query, candidate, href, text) {
    let score = 0;

    if (candidate === query) score += 120;
    if (candidate.startsWith(`${query} `) || candidate.startsWith(`${query}-`)) score += 90;
    if (candidate.startsWith(query)) score += 70;

    for (const word of query.split(/[\s-]+/).filter(Boolean)) {
      if (candidate.split(/[\s-]+/).includes(word)) score += 15;
    }

    if (BAD_LOCATION_RE.test(`${href} ${text}`)) {
      score -= 1000;
    }

    return score;
  }

  function make10DaysUrl(url) {
    url = url.split('?')[0];

    url = url
      .replace('/now/', '/')
      .replace('/today/', '/')
      .replace('/tomorrow/', '/')
      .replace('/3-days/', '/')
      .replace('/week/', '/')
      .replace('/10-days/', '/')
      .replace('/2-weeks/', '/')
      .replace('/month/', '/');

    if (!url.endsWith('/')) url += '/';

    return url + '10-days/';
  }

  async function loadForecastWithRetries(url, attempts = 3, expectedCity = null) {
    let lastError = null;

    for (let attempt = 1; attempt <= attempts; attempt++) {
      let loaded = null;

      try {
        loaded = await loadPageWithIframe(url, 15000);
        await waitForForecastReady(loaded.doc, 12000);
        ensureForecastIsMainCity(loaded.doc, expectedCity);

        const parsed = parseForecast(loaded.doc);
        const hasRainValues = parsed.rain.some(value =>
          value !== '' && value !== null && value !== undefined && Number.isFinite(Number(value))
        );
        const hasDays = parsed.days.filter(Boolean).length >= 10;

        if (!hasDays || !hasRainValues) {
          throw new Error('прогноз загрузился без полных данных');
        }

        return parsed;
      } catch (e) {
        if (isBadForecastError(e)) throw e;

        lastError = e;
      } finally {
        if (loaded?.iframe) loaded.iframe.remove();
      }

      if (attempt < attempts) {
        await sleep(1000 * attempt);
      }
    }

    throw lastError || new Error('не удалось загрузить прогноз');
  }

  function ensureForecastIsMainCity(doc, expectedCity) {
    const pageLocationText = getForecastLocationText(doc);

    if (BAD_LOCATION_RE.test(pageLocationText)) {
      throw makeBadForecastError(`прогноз относится не к основному городу: ${pageLocationText || expectedCity?.name || ''}`);
    }
  }

  function getForecastLocationText(doc) {
    const selectors = [
      'h1',
      'meta[property="og:title"]',
      'meta[name="twitter:title"]'
    ];
    const parts = [doc.title || ''];

    for (const selector of selectors) {
      const node = doc.querySelector(selector);
      const text = node ? cleanCityName(node.getAttribute('content') || node.innerText || node.textContent || '') : '';

      if (text) parts.push(text);
    }

    return parts.join(' ').slice(0, 1000);
  }

  function makeBadForecastError(message) {
    const error = new Error(message);
    error.code = 'BAD_FORECAST_LOCATION';
    return error;
  }

  function isBadForecastError(error) {
    return error && error.code === 'BAD_FORECAST_LOCATION';
  }

  function loadPageWithIframe(url, timeout = 9000) {
    return new Promise((resolve, reject) => {
      const iframe = document.createElement('iframe');

      iframe.style.position = 'fixed';
      iframe.style.left = '-99999px';
      iframe.style.top = '0';
      iframe.style.width = '760px';
      iframe.style.height = '1200px';
      iframe.style.opacity = '0';
      iframe.style.pointerEvents = 'none';

      let done = false;

      const finish = () => {
        if (done) return;
        done = true;

        try {
          const doc = iframe.contentDocument;
          if (doc && doc.body && doc.body.innerText.length > 100) {
            resolve({ iframe, doc });
          } else {
            iframe.remove();
            reject(new Error('страница не загрузилась'));
          }
        } catch (e) {
          iframe.remove();
          reject(e);
        }
      };

      const timer = setTimeout(finish, timeout);

      iframe.onload = () => {
        setTimeout(() => {
          clearTimeout(timer);
          finish();
        }, 5000);
      };

      iframe.src = url;
      document.body.appendChild(iframe);
    });
  }

  function parseForecast(doc) {
    const text = doc.body.innerText.replace(/\u00A0/g, ' ').replace(/\s+/g, ' ');

    let rain = extractRainValuesFromDom(doc);
    const marker = 'Осадки в жидком эквиваленте, мм';
    const idx = text.indexOf(marker);

    if (rain.length < 10 && idx !== -1) {
      const chunk = text.slice(idx + marker.length, idx + marker.length + 1000);

      rain = chunk
        .match(/\d+(?:[,.]\d+)?/g)
        ?.map(x => Number(String(x).replace(',', '.')))
        .slice(0, 10) || [];
    }

    while (rain.length < 10) rain.push('');
    const wind = extractWindValuesFromDom(doc);
    while (wind.length < 10) wind.push('');
    const gust = extractGustValuesFromDom(doc);
    while (gust.length < 10) gust.push('');

    return {
      days: extractForecastDays(doc, text),
      rain,
      wind,
      gust,
      weather: extractWeatherTexts(doc)
    };
  }

  function extractRainValuesFromDom(doc) {
    const rows = [
      ...doc.querySelectorAll('[data-row="precipitation"], [data-row*="precipitation"], .widget-row-chart-precipitation')
    ];

    for (const row of rows) {
      const values = getForecastRowItems(row)
        .map(el => normalizeForecastText(el.innerText || el.textContent || ''))
        .map(text => text.match(/\d+(?:[,.]\d+)?/)?.[0])
        .filter(Boolean)
        .map(value => Number(String(value).replace(',', '.')));

      if (values.length >= 10) return values.slice(0, 10);

      const fromText = normalizeForecastText(row.innerText || row.textContent || '')
        .match(/\d+(?:[,.]\d+)?/g)
        ?.map(value => Number(String(value).replace(',', '.')))
        .slice(0, 10) || [];

      if (fromText.length >= 10) return fromText;
    }

    return [];
  }

  function getForecastRowItems(row) {
    const directContainer = row.querySelector(':scope > .widget-items');
    const directItems = directContainer
      ? [...directContainer.children].filter(child => child.classList?.contains('row-item'))
      : [];

    if (directItems.length >= 10) return directItems.slice(0, 10);

    const ownItems = [...row.children].filter(child => child.classList?.contains('row-item'));
    if (ownItems.length >= 10) return ownItems.slice(0, 10);

    const firstItemsContainer = row.querySelector('.widget-items');
    const nestedItems = firstItemsContainer
      ? [...firstItemsContainer.children].filter(child => child.classList?.contains('row-item'))
      : [];

    if (nestedItems.length >= 10) return nestedItems.slice(0, 10);

    return [...row.querySelectorAll('.row-item')].slice(0, 10);
  }

  function extractWindValuesFromDom(doc) {
    return extractWindMetricValuesFromDom(doc, 'wind-speed');
  }

  function extractGustValuesFromDom(doc) {
    return extractWindMetricValuesFromDom(doc, 'wind-gust');
  }

  function extractWindMetricValuesFromDom(doc, metricClass) {
    const rows = [
      ...doc.querySelectorAll('[data-row="wind"], .widget-row-wind')
    ];

    for (const row of rows) {
      const items = getForecastRowItems(row);
      const values = items
        .map(item => {
          const speed = item.querySelector(`.${metricClass} speed-value[value], [class*="${metricClass}"] speed-value[value]`);
          const visibleValue = normalizeForecastText(speed?.innerText || speed?.textContent || '');
          const visibleMatch = visibleValue.match(/\d+(?:[,.]\d+)?/);

          if (visibleMatch) {
            return Number(visibleMatch[0].replace(',', '.'));
          }

          const attrValue = speed?.getAttribute('value');

          if (attrValue !== null && attrValue !== undefined && attrValue !== '') {
            return Number(String(attrValue).replace(',', '.'));
          }

          if (metricClass !== 'wind-speed') return NaN;

          const text = normalizeForecastText(item.innerText || item.textContent || '');
          const match = text.match(/\d+(?:[,.]\d+)?/);

          return match ? Number(match[0].replace(',', '.')) : NaN;
        })
        .filter(value => Number.isFinite(value));

      if (values.length >= 10) return values.slice(0, 10);
    }

    return [];
  }

  function extractForecastDays(doc, text) {
    const months = 'января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря';
    const fromDom = extractForecastDaysFromDom(doc);
    const sequentialFromDom = buildSequentialForecastDays(fromDom[0]);

    if (sequentialFromDom.length >= 10) {
      return sequentialFromDom;
    }

    const dayRe = new RegExp(`(?:В\\s*с|П\\s*н|В\\s*т|С\\s*р|Ч\\s*т|П\\s*т|С\\s*б)\\s+\\d{1,2}\\s+(?:${months})`, 'gi');
    const matches = [
      ...fromDom,
      ...[...String(text).matchAll(dayRe)]
        .map(match => normalizeForecastDay(match[0]))
    ];
    const unique = [];

    for (const day of matches) {
      if (!unique.includes(day)) unique.push(day);
      if (unique.length >= 10) break;
    }

    const sequentialFromText = buildSequentialForecastDays(unique[0]);

    if (sequentialFromText.length >= 10) {
      return sequentialFromText;
    }

    while (unique.length < 10) {
      unique.push(buildDates()[unique.length]);
    }

    return unique.slice(0, 10);
  }

  function buildSequentialForecastDays(firstDay) {
    const firstDate = makeExcelDateFromForecastDay(firstDay, 0);

    if (!firstDate) return [];

    const week = ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];
    const months = [
      'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
      'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'
    ];

    return Array.from({ length: 10 }, (_, index) => {
      const date = new Date(Date.UTC(
        firstDate.getUTCFullYear(),
        firstDate.getUTCMonth(),
        firstDate.getUTCDate() + index
      ));

      return `${week[date.getUTCDay()]} ${date.getUTCDate()} ${months[date.getUTCMonth()]}`;
    });
  }

  function extractForecastDaysFromDom(doc) {
    return [...doc.querySelectorAll('.widget-row-date .row-item')]
      .map(el => normalizeForecastText(el.innerText || el.textContent || ''))
      .map(normalizeForecastDay)
      .filter(day => makeExcelDateFromForecastDay(day, 0))
      .slice(0, 10);
  }

  function normalizeForecastText(text) {
    return String(text || '')
      .replace(/\u00A0/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function normalizeForecastDay(text) {
    const value = normalizeForecastText(text);
    const compactWeekday = value.replace(/\s+/g, '').match(/^(Вс|Пн|Вт|Ср|Чт|Пт|Сб)/i)?.[1] || '';
    const date = value.match(/\b(\d{1,2})\s+([а-яё]+)\b/i);

    return compactWeekday && date
      ? `${capitalizeWeekday(compactWeekday)} ${Number(date[1])} ${date[2].toLowerCase()}`
      : value;
  }

  function capitalizeWeekday(value) {
    const normalized = String(value || '').toLowerCase();
    const map = {
      вс: 'Вс',
      пн: 'Пн',
      вт: 'Вт',
      ср: 'Ср',
      чт: 'Чт',
      пт: 'Пт',
      сб: 'Сб'
    };

    return map[normalized] || value;
  }

  function extractWeatherTexts(doc) {
    const dates = [...doc.querySelectorAll('.widget-row-date .row-item')]
      .map(el => normalizeForecastText(el.innerText || el.textContent || ''))
      .filter(Boolean);
    const limit = dates.length || 10;
    const weather = [...doc.querySelectorAll('div.row-item[data-tooltip]')]
      .map(el => el.dataset.tooltip)
      .filter(text =>
        text &&
        !text.includes('ветер') &&
        !text.includes('возмущ') &&
        !text.includes('буря')
      )
      .map(cleanWeatherText)
      .slice(0, limit);

    return Array.from({ length: 10 }, (_, index) => weather[index] || '');
  }

  function findForecastIconTooltipElements(doc) {
    const rows = [
      ...doc.querySelectorAll('[data-row="icon-tooltip"]'),
      ...doc.querySelectorAll('.widget-row-icon')
    ];

    for (const row of rows) {
      const items = [...row.querySelectorAll(':scope > .row-item[data-tooltip], :scope .widget-items > .row-item[data-tooltip]')]
        .filter(el => WEATHER_TEXT_RE.test(cleanWeatherText(el.getAttribute('data-tooltip'))));

      if (items.length >= 10) return items.slice(0, 10);
    }

    const items = [...doc.querySelectorAll('.widget-row-icon .row-item[data-tooltip], [data-row="icon-tooltip"] .row-item[data-tooltip]')]
      .filter(el => WEATHER_TEXT_RE.test(cleanWeatherText(el.getAttribute('data-tooltip'))));

    return items.slice(0, 10);
  }

  function cleanWeatherText(text) {
    return String(text || '')
      .replace(/\u00A0/g, ' ')
      .replace(/\s+/g, ' ')
      .replace(/[.!?]+$/g, '')
      .trim();
  }

  function isSpecificWeatherText(text) {
    return WEATHER_TEXT_RE.test(text) && !/^погода\b/i.test(text);
  }

  function getWeatherFallbackText(value) {
    const rain = Number(value);

    if (!Number.isFinite(rain) || rain <= 0) return 'Облачно';
    if (rain >= 8) return 'Облачно, дождь, гроза';
    if (rain >= 2) return 'Облачно, дождь';

    return 'Небольшой дождь';
  }

  function waitForForecastReady(doc, timeout = 8000) {
    const started = Date.now();

    return new Promise(resolve => {
      const check = () => {
        const hasWeather = extractWeatherTexts(doc).filter(Boolean).length >= 10;
        const hasRain = doc.body?.innerText?.includes('Осадки в жидком эквиваленте') ||
          extractRainValuesFromDom(doc).length >= 10;
        const hasWind = extractWindValuesFromDom(doc).length >= 10 &&
          extractGustValuesFromDom(doc).length >= 10;

        if ((hasWeather && hasRain && hasWind) || Date.now() - started >= timeout) {
          resolve();
          return;
        }

        setTimeout(check, 300);
      };

      check();
    });
  }

  function makeRainChartImage(city, parsed) {
    const width = 980;
    const height = 360;
    const canvas = document.createElement('canvas');
    const ratio = window.devicePixelRatio || 1;
    canvas.width = width * ratio;
    canvas.height = height * ratio;
    canvas.style.width = `${width}px`;
    canvas.style.height = `${height}px`;

    const ctx = canvas.getContext('2d');
    ctx.scale(ratio, ratio);
    ctx.fillStyle = '#f6f8fb';
    ctx.fillRect(0, 0, width, height);

    roundRect(ctx, 8, 8, width - 16, height - 16, 12, '#ffffff');

    ctx.fillStyle = '#222';
    ctx.font = 'bold 22px Arial';
    ctx.fillText(`Осадки по дням: ${city}`, 30, 44);

    ctx.fillStyle = '#6b7280';
    ctx.font = '12px Arial';
    ctx.fillText('Красным выделены дни с осадками больше 5 мм', 30, 66);

    const rainValues = parsed.rain.map(value => {
      const number = Number(value);
      return Number.isFinite(number) ? number : 0;
    });

    const maxValue = Math.max(5, ...rainValues);
    const chartLeft = 58;
    const chartTop = 92;
    const chartWidth = width - 102;
    const chartHeight = 180;
    const baseline = chartTop + chartHeight;

    ctx.strokeStyle = '#dce4ea';
    ctx.lineWidth = 1;
    ctx.fillStyle = '#7b8790';
    ctx.font = '11px Arial';

    for (let i = 0; i <= 4; i++) {
      const y = chartTop + (chartHeight / 4) * i;
      const value = maxValue - (maxValue / 4) * i;
      ctx.beginPath();
      ctx.moveTo(chartLeft, y);
      ctx.lineTo(chartLeft + chartWidth, y);
      ctx.stroke();
      ctx.fillText(formatRain(value), 20, y + 3);
    }

    const gap = 14;
    const barWidth = (chartWidth - gap * 9) / 10;

    rainValues.forEach((value, index) => {
      const x = chartLeft + index * (barWidth + gap);
      const barHeight = value > 0 ? Math.max(5, (value / maxValue) * chartHeight) : 2;
      const y = baseline - barHeight;
      const isStrong = value > 5;

      const gradient = ctx.createLinearGradient(0, y, 0, baseline);
      gradient.addColorStop(0, isStrong ? '#ff5a5a' : '#6bb7f0');
      gradient.addColorStop(1, isStrong ? '#d90000' : '#a7d8fb');

      roundRect(ctx, x, y, barWidth, barHeight, 7, gradient);

      ctx.fillStyle = isStrong ? '#c00000' : '#0070c0';
      ctx.font = 'bold 13px Arial';
      ctx.textAlign = 'center';
      ctx.fillText(formatRain(value), x + barWidth / 2, y - 7);

      ctx.fillStyle = '#222';
      ctx.font = '12px Arial';
      const day = parsed.days[index] || '';
      const parts = day.split(' ');
      ctx.fillText(parts.slice(0, 2).join(' '), x + barWidth / 2, baseline + 22);
      ctx.fillStyle = '#6b7280';
      ctx.font = '11px Arial';
      ctx.fillText(parts.slice(2).join(' '), x + barWidth / 2, baseline + 40);
    });

    ctx.textAlign = 'left';
    ctx.fillStyle = '#6b7280';
    ctx.font = '12px Arial';
    ctx.fillText('Единица измерения: мм за сутки', 30, height - 28);

    return {
      dataUrl: canvas.toDataURL('image/png'),
      width,
      height
    };
  }

  function makeWindChartImage(city, parsed) {
    const width = 980;
    const height = 360;
    const canvas = document.createElement('canvas');
    const ratio = window.devicePixelRatio || 1;
    canvas.width = width * ratio;
    canvas.height = height * ratio;
    canvas.style.width = `${width}px`;
    canvas.style.height = `${height}px`;

    const ctx = canvas.getContext('2d');
    ctx.scale(ratio, ratio);
    ctx.fillStyle = '#f6f8fb';
    ctx.fillRect(0, 0, width, height);

    roundRect(ctx, 8, 8, width - 16, height - 16, 12, '#ffffff');

    ctx.fillStyle = '#222';
    ctx.font = 'bold 22px Arial';
    ctx.fillText(`Ветер и порывы по дням: ${city}`, 30, 44);

    ctx.fillStyle = '#6b7280';
    ctx.font = '12px Arial';
    ctx.fillText('Красным выделены значения больше 10 м/с', 30, 66);

    const windValues = (parsed.wind || []).map(value => {
      const number = Number(value);
      return Number.isFinite(number) ? number : 0;
    });
    while (windValues.length < 10) windValues.push(0);

    const gustValues = (parsed.gust || []).map(value => {
      const number = Number(value);
      return Number.isFinite(number) ? number : 0;
    });
    while (gustValues.length < 10) gustValues.push(0);

    const maxValue = Math.max(10, ...windValues, ...gustValues);
    const chartLeft = 58;
    const chartTop = 92;
    const chartWidth = width - 102;
    const chartHeight = 180;
    const baseline = chartTop + chartHeight;

    ctx.strokeStyle = '#dce4ea';
    ctx.lineWidth = 1;
    ctx.fillStyle = '#7b8790';
    ctx.font = '11px Arial';

    for (let i = 0; i <= 4; i++) {
      const y = chartTop + (chartHeight / 4) * i;
      const value = maxValue - (maxValue / 4) * i;
      ctx.beginPath();
      ctx.moveTo(chartLeft, y);
      ctx.lineTo(chartLeft + chartWidth, y);
      ctx.stroke();
      ctx.fillText(formatWind(value), 20, y + 3);
    }

    const groupGap = 14;
    const groupWidth = (chartWidth - groupGap * 9) / 10;
    const innerGap = 5;
    const barWidth = (groupWidth - innerGap) / 2;

    windValues.slice(0, 10).forEach((windValue, index) => {
      const groupX = chartLeft + index * (groupWidth + groupGap);
      const gustValue = gustValues[index] || 0;

      drawWindBar(ctx, {
        x: groupX,
        baseline,
        chartHeight,
        maxValue,
        barWidth,
        value: windValue,
        colorTop: '#fbbf24',
        colorBottom: '#d97706'
      });

      drawWindBar(ctx, {
        x: groupX + barWidth + innerGap,
        baseline,
        chartHeight,
        maxValue,
        barWidth,
        value: gustValue,
        colorTop: '#a78bfa',
        colorBottom: '#7c3aed'
      });

      ctx.fillStyle = '#222';
      ctx.font = '12px Arial';
      ctx.textAlign = 'center';
      const day = parsed.days[index] || '';
      const parts = day.split(' ');
      ctx.fillText(parts.slice(0, 2).join(' '), groupX + groupWidth / 2, baseline + 22);
      ctx.fillStyle = '#6b7280';
      ctx.font = '11px Arial';
      ctx.fillText(parts.slice(2).join(' '), groupX + groupWidth / 2, baseline + 40);
    });

    ctx.textAlign = 'left';
    ctx.fillStyle = '#6b7280';
    ctx.font = '12px Arial';
    ctx.fillText('Оранжевый столбец: ветер. Фиолетовый столбец: порывы. Единица измерения: м/с за сутки.', 30, height - 28);

    return {
      dataUrl: canvas.toDataURL('image/png'),
      width,
      height
    };
  }

  function drawWindBar(ctx, options) {
    const value = Number(options.value) || 0;
    const barHeight = value > 0 ? Math.max(5, (value / options.maxValue) * options.chartHeight) : 2;
    const y = options.baseline - barHeight;
    const isStrong = value > 10;
    const gradient = ctx.createLinearGradient(0, y, 0, options.baseline);
    gradient.addColorStop(0, options.colorTop);
    gradient.addColorStop(1, options.colorBottom);

    roundRect(ctx, options.x, y, options.barWidth, barHeight, 7, gradient);

    ctx.fillStyle = isStrong ? '#c00000' : '#0070c0';
    ctx.font = 'bold 13px Arial';
    ctx.textAlign = 'center';
    ctx.fillText(formatWind(value), options.x + options.barWidth / 2, y - 7);
  }

  function roundRect(ctx, x, y, width, height, radius, fillStyle) {
    const r = Math.min(radius, width / 2, height / 2);

    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.arcTo(x + width, y, x + width, y + height, r);
    ctx.arcTo(x + width, y + height, x, y + height, r);
    ctx.arcTo(x, y + height, x, y, r);
    ctx.arcTo(x, y, x + width, y, r);
    ctx.closePath();
    ctx.fillStyle = fillStyle;
    ctx.fill();
  }

  function formatRain(value) {
    return Number(value).toFixed(1).replace('.', ',');
  }

  function formatWind(value) {
    return String(Math.round(Number(value))).replace('.', ',');
  }

  function buildDates() {
    const week = ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'];
    const months = [
      'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
      'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'
    ];

    const result = [];

    for (let i = 0; i < 10; i++) {
      const d = new Date();
      d.setDate(d.getDate() + i);
      result.push(`${week[d.getDay()]} ${d.getDate()} ${months[d.getMonth()]}`);
    }

    return result;
  }

  async function exportExcel(blocks, fixedCities = [], reportCities = []) {
    const workbook = new ExcelJS.Workbook();
    workbook.calcProperties = workbook.calcProperties || {};
    workbook.calcProperties.fullCalcOnLoad = true;
    const ws = workbook.addWorksheet('Отчет', {
      views: [{ showGridLines: false }]
    });

    ws.columns = [
      { width: 18 }, { width: 14 }, { width: 14 }, { width: 14 },
      { width: 14 }, { width: 14 }, { width: 14 }, { width: 14 },
      { width: 14 }, { width: 14 }, { width: 14 }, { width: 11 }
    ];

    const dataBorder = {
      top: { style: 'thin', color: { argb: 'FFD6DEE6' } },
      left: { style: 'thin', color: { argb: 'FFD6DEE6' } },
      bottom: { style: 'thin', color: { argb: 'FFD6DEE6' } },
      right: { style: 'thin', color: { argb: 'FFD6DEE6' } }
    };

    const okCount = blocks.filter(b => b.status === 'OK').length;
    const totalCount = blocks.length;
    const allOk = okCount === totalCount;

    ws.mergeCells(1, 1, 1, 5);
    ws.getCell(1, 1).value = {
      richText: [
        {
          text: 'Получены данные по городам ',
          font: { name: 'Arial', size: 11, bold: true, color: { argb: 'FF00B050' } }
        },
        {
          text: String(okCount),
          font: {
            name: 'Arial',
            size: 11,
            bold: true,
            color: { argb: allOk ? 'FF00B050' : 'FFFF0000' }
          }
        },
        {
          text: `/${totalCount}`,
          font: { name: 'Arial', size: 11, bold: true, color: { argb: 'FF00B050' } }
        }
      ]
    };
    ws.getCell(1, 1).alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getRow(1).height = 22;
    let row = writeFailRows(ws, blocks);
    const blockGapRows = 1;

    for (let blockIndex = 0; blockIndex < blocks.length; blockIndex++) {
      const block = blocks[blockIndex];

      if (blockIndex > 0) {
        addDashedSeparator(ws, row);
        ws.getRow(row).height = 10;
        row += 2;
      }

      ws.getRow(row).height = 18;

      ws.getCell(row, 1).value = block.url
        ? {
          text: '🔗',
          hyperlink: block.url,
          tooltip: `Открыть прогноз: ${block.city}`
        }
        : '';
      ws.getCell(row, 1).font = {
        name: 'Segoe UI Emoji',
        size: 12,
        underline: false,
        color: { argb: 'FF0563C1' }
      };
      ws.getCell(row, 1).alignment = { horizontal: 'left', vertical: 'middle' };
      ws.getCell(row, 1).style = {
        ...ws.getCell(row, 1).style,
        font: {
          name: 'Segoe UI Emoji',
          size: 12,
          underline: false,
          color: { argb: 'FF0563C1' }
        }
      };
      ws.getCell(row, 1).border = {
        bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } }
      };

      const cityRow = row + 1;
      const rainRow = row + 2;
      const windRow = row + 3;
      const gustRow = row + 4;

      ws.getRow(cityRow).height = 22;
      ws.getRow(rainRow).height = 22;
      ws.getRow(windRow).height = 22;
      ws.getRow(gustRow).height = 22;

      ws.getCell(cityRow, 1).value = block.city;
      ws.getCell(rainRow, 1).value = 'Осадки, мм';
      ws.getCell(windRow, 1).value = 'Ветер, м/с';
      ws.getCell(gustRow, 1).value = 'Порывы, м/с';

      for (let i = 0; i < 10; i++) {
        const rainNumber = Number(block.rain[i]);
        const windNumber = Number(block.wind?.[i]);
        const gustNumber = Number(block.gust?.[i]);

        ws.getCell(cityRow, i + 2).value = block.days[i];
        ws.getCell(rainRow, i + 2).value = Number.isFinite(rainNumber) ? rainNumber : '';
        ws.getCell(rainRow, i + 2).numFmt = '0.0';
        ws.getCell(windRow, i + 2).value = Number.isFinite(windNumber) ? windNumber : '';
        ws.getCell(windRow, i + 2).numFmt = '0';
        ws.getCell(gustRow, i + 2).value = Number.isFinite(gustNumber) ? gustNumber : '';
        ws.getCell(gustRow, i + 2).numFmt = '0';
      }

      ws.mergeCells(cityRow, 12, gustRow, 12);
      ws.getCell(cityRow, 12).value = block.status;

      for (let c = 1; c <= 12; c++) {
        const cityCell = ws.getCell(cityRow, c);
        cityCell.border = dataBorder;
        cityCell.alignment = { horizontal: 'center', vertical: 'middle' };
        cityCell.font = {
          name: 'Arial',
          size: 10,
          bold: c === 12,
          color: {
            argb: c === 12
              ? block.status === 'OK'
                ? 'FF00B050'
                : 'FFFF0000'
              : 'FF000000'
          }
        };
        cityCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: c === 1
              ? 'FFFFE3C7'
              : c === 12
                ? 'FFFFFFFF'
                : 'FFF8FAFC'
          }
        };

        const rainCell = ws.getCell(rainRow, c);
        rainCell.border = dataBorder;
        rainCell.alignment = { horizontal: 'center', vertical: 'middle' };
        rainCell.font = {
          name: 'Arial',
          size: 10,
          bold: c === 12,
          color: {
            argb: c === 12
              ? block.status === 'OK'
                ? 'FF00B050'
                : 'FFFF0000'
              : 'FF000000'
          }
        };

        const rainValue = c >= 2 && c <= 11
          ? Number(block.rain[c - 2])
          : 0;

        rainCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: c >= 2 && c <= 11
              ? rainValue > 5
                ? 'FFFFD6D6'
                : 'FFE1F0FB'
              : c === 1
                ? 'FFF8FAFC'
                : 'FFFFFFFF'
          }
        };

        const windCell = ws.getCell(windRow, c);
        windCell.border = dataBorder;
        windCell.alignment = { horizontal: 'center', vertical: 'middle' };
        windCell.font = {
          name: 'Arial',
          size: 10,
          color: { argb: 'FF000000' }
        };

        const windValue = c >= 2 && c <= 11
          ? Number(block.wind?.[c - 2])
          : 0;

        windCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: c >= 2 && c <= 11
              ? windValue > 10
                ? 'FFFFD6D6'
                : 'FFE1F0FB'
              : c === 1
                ? 'FFF8FAFC'
                : 'FFFFFFFF'
          }
        };

        const gustCell = ws.getCell(gustRow, c);
        gustCell.border = dataBorder;
        gustCell.alignment = { horizontal: 'center', vertical: 'middle' };
        gustCell.font = {
          name: 'Arial',
          size: 10,
          color: { argb: 'FF000000' }
        };

        const gustValue = c >= 2 && c <= 11
          ? Number(block.gust?.[c - 2])
          : 0;

        gustCell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: c >= 2 && c <= 11
              ? gustValue > 10
                ? 'FFFFD6D6'
                : 'FFE1F0FB'
              : c === 1
                ? 'FFF8FAFC'
                : 'FFFFFFFF'
          }
        };
      }

      const statusCell = ws.getCell(cityRow, 12);
      statusCell.font = {
        name: 'Arial',
        size: 10,
        bold: true,
        color: { argb: block.status === 'OK' ? 'FF00B050' : 'FFFF0000' }
      };
      statusCell.alignment = { horizontal: 'center', vertical: 'middle' };
      statusCell.border = dataBorder;

      row += 6;

      const visualTopRow = row - 1;
      let visualRows = 0;
      let nextChartTopRow = visualTopRow;

      if (block.rainChart) {
        const chartId = workbook.addImage({
          base64: block.rainChart.dataUrl,
          extension: 'png'
        });

        const chartWidth = block.rainChart.width || 980;
        const chartHeight = block.rainChart.height || 360;

        ws.addImage(chartId, {
          tl: { col: 0, row: visualTopRow },
          ext: { width: chartWidth, height: chartHeight }
        });

        const chartRows = pixelsToWorksheetRows(chartHeight);
        nextChartTopRow = visualTopRow + chartRows + 1;
        visualRows += chartRows;
      }

      if (block.windChart) {
        const chartId = workbook.addImage({
          base64: block.windChart.dataUrl,
          extension: 'png'
        });

        const chartWidth = block.windChart.width || 980;
        const chartHeight = block.windChart.height || 360;

        ws.addImage(chartId, {
          tl: { col: 0, row: nextChartTopRow },
          ext: { width: chartWidth, height: chartHeight }
        });

        visualRows += (block.rainChart ? 1 : 0) + pixelsToWorksheetRows(chartHeight);
      }

      if (visualRows > 0) {
        row += visualRows;
      }

      row += blockGapRows;
    }

    const osadkiDays = getOrderedDays(blocks);
    const rainFilter = buildListSheet(workbook, blocks, dataBorder, osadkiDays, reportCities);
    const gustFilter = buildGustSheet(workbook, blocks, dataBorder, osadkiDays, fixedCities, reportCities);

    let buffer = await workbook.xlsx.writeBuffer();
    buffer = await applyCurrentDayFilters(buffer, [rainFilter, gustFilter]);

    saveAs(
      new Blob([buffer]),
      `gismeteo_osadki_${new Date().toISOString().slice(0, 10)}.xlsx`
    );
  }

  function writeFailRows(ws, blocks) {
    const failedBlocks = blocks.filter(block => block.status === 'FAIL');

    if (!failedBlocks.length) {
      addDashedSeparator(ws, 2);
      ws.getRow(2).height = 10;
      return 4;
    }

    let row = 2;

    for (const block of failedBlocks) {
      ws.mergeCells(row, 1, row, 5);
      const cell = ws.getCell(row, 1);
      cell.value = `! ${block.city} - FAIL !`;
      cell.font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFF0000' } };
      cell.alignment = { horizontal: 'left', vertical: 'middle' };
      ws.getRow(row).height = 20;
      row += 1;
    }

    addDashedSeparator(ws, row - 1);

    return row + 1;
  }

  function pixelsToWorksheetRows(pixels) {
    return Math.max(1, Math.ceil(Number(pixels || 0) / 20));
  }

  function addDashedSeparator(ws, row, fromColumn = 1, toColumn = 12) {
    for (let column = fromColumn; column <= toColumn; column++) {
      const cell = ws.getCell(row, column);
      cell.border = {
        ...cell.border,
        bottom: { style: 'mediumDashed', color: { argb: 'FF64748B' } }
      };
    }
  }

  function buildListSheet(workbook, blocks, dataBorder, osadkiDays, reportCities = []) {
    const ws = workbook.addWorksheet('Осадки', {
      views: [{ showGridLines: false }]
    });

    ws.columns = [
      { width: 18 },
      { width: 24 },
      { width: 58 },
      { width: 9 }
    ];

    ws.mergeCells(1, 1, 1, 4);
    ws.getCell(1, 1).value = 'Города с осадками больше 5 мм';
    ws.getCell(1, 1).font = { name: 'Arial', size: 13, bold: true, color: { argb: 'FF222222' } };
    ws.getCell(1, 1).alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getRow(1).height = 24;

    const reportCitySet = buildFixedCitySet(reportCities);
    const listRows = buildStrongRainRows(blocks, osadkiDays, reportCitySet);
    const currentDay = osadkiDays[0] || buildDates()[0];
    const currentDate = makeExcelDateFromForecastDay(currentDay, 0);

    ws.getCell(2, 1).value = `По умолчанию показан текущий день: ${formatShortDate(currentDay)}. Выберите другую дату через фильтр.`;
    ws.getCell(2, 1).font = { name: 'Arial', size: 10, color: { argb: 'FF6B7280' } };

    const lastRow = writeListTable(ws, listRows, 4, dataBorder, true, currentDate, {
      valueFill: 'FFFFD6D6'
    });

    return {
      sheetName: 'Осадки',
      ref: `A4:D${Math.max(4, lastRow - 1)}`,
      currentDate
    };
  }

  function buildGustSheet(workbook, blocks, dataBorder, osadkiDays, fixedCities = [], reportCities = []) {
    const ws = workbook.addWorksheet('Осадки-Порывы', {
      views: [{ showGridLines: false }]
    });

    ws.columns = [
      { width: 18 },
      { width: 24 },
      { width: 58 },
      { width: 9 },
      { width: 3 },
      { width: 28 },
      { width: 1 },
      { width: 14 },
      { width: 24 },
      { width: 58 },
      { width: 9 },
      { width: 8 },
      { width: 8 },
      { width: 8 },
      { width: 8 },
      { width: 14 }
    ];
    ws.getColumn(7).hidden = true;
    ws.getColumn(8).hidden = true;
    ws.getColumn(9).hidden = true;
    ws.getColumn(10).hidden = true;
    ws.getColumn(11).hidden = true;
    ws.getColumn(12).hidden = true;
    ws.getColumn(13).hidden = true;
    ws.getColumn(14).hidden = true;
    ws.getColumn(15).hidden = true;
    ws.getColumn(16).hidden = true;

    ws.mergeCells(1, 1, 1, 4);
    ws.getCell(1, 1).value = 'Города с осадками и порывами больше 10 м/с';
    ws.getCell(1, 1).font = { name: 'Arial', size: 13, bold: true, color: { argb: 'FF222222' } };
    ws.getCell(1, 1).alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getRow(1).height = 24;

    const currentDay = osadkiDays[0] || buildDates()[0];
    const currentDate = makeExcelDateFromForecastDay(currentDay, 0);
    const fixedList = dedupeCityList(fixedCities);
    const fixedCitySet = buildFixedCitySet(fixedList);
    const reportList = dedupeCityList(reportCities);
    const reportCitySet = buildFixedCitySet(reportList);
    const sourceCitySet = mergeCitySets(reportCitySet, fixedCitySet);
    const cityOrder = buildFixedCityOrderMap([...fixedList, ...reportList]);
    const listRows = buildRainGustRows(blocks, osadkiDays, sourceCitySet, cityOrder);

    ws.getCell(2, 1).value = `По умолчанию выбран текущий день: ${formatShortDate(currentDay)}. Выберите другую дату в поле ниже.`;
    ws.getCell(2, 1).font = { name: 'Arial', size: 10, color: { argb: 'FF6B7280' } };

    writeRainGustDatePicker(ws, osadkiDays, currentDay, dataBorder);
    writeFixedCityBox(ws, fixedList, dataBorder);
    const headerRow = getRainGustTableHeaderRow(fixedList);
    const lastRow = writeDynamicRainGustTable(ws, listRows, fixedCitySet, reportCitySet, headerRow, dataBorder, currentDate, {
      valueHeader: 'Осадки/Порывы',
      emptyText: 'Нет дней с осадками или порывами',
      getValueFill: item => item.rain > 5 || item.gust > 10 ? 'FFFFD6D6' : null,
      formatValue: item => makeRainGustListText(item.weather, item.rain, item.gust),
      fixedRowsCount: fixedList.length
    });

    return {
      sheetName: 'Осадки-Порывы',
      ref: `A${headerRow}:D${Math.max(headerRow, lastRow - 1)}`,
      currentDate: null
    };
  }

  function getRainGustTableHeaderRow(fixedCities) {
    const fixedRowsCount = fixedCities.length || 6;
    return 6 + fixedRowsCount + 2;
  }

  function writeDynamicRainGustTable(ws, items, fixedCitySet, reportCitySet, headerRow, dataBorder, currentDate, options = {}) {
    const headers = ['Дата', 'Город', options.valueHeader || 'Осадки/Порывы', ''];
    const dataStartRow = headerRow + 1;
    const maxRows = Math.max(1, items.length);
    const lastDataRow = dataStartRow + maxRows - 1;
    const rawStartRow = dataStartRow;
    const rawLastRow = rawStartRow + Math.max(0, items.length - 1);
    const fixedLastRow = 6 + (options.fixedRowsCount || 1);
    const fixedRange = `$A$7:$A$${fixedLastRow}`;

    headers.forEach((header, index) => {
      const cell = ws.getCell(headerRow, index + 1);
      cell.value = header;
      cell.border = dataBorder;
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF8FAFC' }
      };
      cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF111827' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    items.forEach((item, index) => {
      const row = rawStartRow + index;

      ws.getCell(row, 8).value = item.dayDate || makeExcelDateFromForecastDay(item.day, item.dayIndex);
      ws.getCell(row, 8).numFmt = 'dd.mm.yyyy';
      ws.getCell(row, 9).value = item.city;
      ws.getCell(row, 10).value = typeof options.formatValue === 'function'
        ? options.formatValue(item)
        : makeRainGustListText(item.weather, item.rain, item.gust);
      ws.getCell(row, 11).value = item.url || '';
      ws.getCell(row, 12).value = item.rain > 5 || item.gust > 10 ? 1 : 0;
      ws.getCell(row, 13).value = {
        formula: `IF(COUNTA(${fixedRange})=0,1,IF(COUNTIF(${fixedRange},I${row})>0,1,0))`,
        result: fixedCitySet
          ? (isCityInFixedSet(item.city, fixedCitySet) ? 1 : 0)
          : 1
      };
      ws.getCell(row, 14).value = {
        formula: `IF(AND(M${row}=1,INT(H${row})=INT($F$4)),SUMPRODUCT(($M$${rawStartRow}:M${row}=1)*(INT($H$${rawStartRow}:H${row})=INT($F$4))),"")`,
        result: initialRainGustVisible(item.city, fixedCitySet, reportCitySet) && isSameExcelDay(item.dayDate, currentDate)
          ? initialVisibleDateIndex(items, index, fixedCitySet, currentDate)
          : ''
      };
      ws.getCell(row, 15).value = isCityInFixedSet(item.city, reportCitySet) ? 1 : 0;
    });

    if (!items.length) {
      ws.mergeCells(dataStartRow, 1, dataStartRow, 4);
      ws.getCell(dataStartRow, 1).value = options.emptyText || 'Нет дней с осадками или порывами';
      ws.getCell(dataStartRow, 1).font = { name: 'Arial', size: 10, color: { argb: 'FF6B7280' } };
      ws.getCell(dataStartRow, 1).alignment = { horizontal: 'left', vertical: 'middle' };
      ws.getCell(dataStartRow, 1).border = dataBorder;
      return dataStartRow + 1;
    }

    for (let row = dataStartRow; row <= lastDataRow; row++) {
      const visibleIndex = row - dataStartRow + 1;
      const initialEntry = items
        .map((item, index) => ({ item, index }))
        .filter(entry =>
          initialRainGustVisible(entry.item.city, fixedCitySet, reportCitySet) &&
          isSameExcelDay(entry.item.dayDate, currentDate)
        )[visibleIndex - 1];
      const initialItem = initialEntry?.item;
      const initialSourceRow = initialEntry ? rawStartRow + initialEntry.index : '';
      const matchFormula = `IFERROR(MATCH(${visibleIndex},$N$${rawStartRow}:$N$${rawLastRow},0)+${rawStartRow - 1},"")`;

      ws.getCell(row, 7).value = { formula: matchFormula, result: initialSourceRow };
      ws.getCell(row, 1).value = {
        formula: `IF($G${row}="","",INDEX($H:$H,$G${row}))`,
        result: initialItem?.dayDate || ''
      };
      ws.getCell(row, 2).value = {
        formula: `IF($G${row}="","",INDEX($I:$I,$G${row}))`,
        result: initialItem?.city || ''
      };
      ws.getCell(row, 3).value = {
        formula: `IF($G${row}="","",INDEX($J:$J,$G${row}))`,
        result: initialItem
          ? (typeof options.formatValue === 'function'
            ? options.formatValue(initialItem)
            : makeRainGustListText(initialItem.weather, initialItem.rain, initialItem.gust))
          : ''
      };
      ws.getCell(row, 4).value = {
        formula: `IF($G${row}="","",HYPERLINK(INDEX($K:$K,$G${row}),"🔗"))`,
        result: initialItem?.url ? '🔗' : ''
      };

      const dayDate = initialItem?.dayDate || null;

      for (let c = 1; c <= 4; c++) {
        const cell = ws.getCell(row, c);
        cell.border = initialItem ? dataBorder : {};
        if (c === 1) {
          cell.numFmt = 'dd.mm.yyyy';
        }
        cell.font = {
          name: c === 4 ? 'Segoe UI Emoji' : 'Arial',
          size: c === 4 ? 12 : 10,
          underline: false,
          color: { argb: c === 4 ? 'FF0563C1' : 'FF000000' }
        };
        cell.alignment = {
          horizontal: 'center',
          vertical: 'middle'
        };
        if (initialItem) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: row % 2 === 0 ? 'FFFFFFFF' : 'FFF8FAFC' }
          };
        }
      }

      ws.getRow(row).hidden = false;
    }

    if (typeof ws.addConditionalFormatting === 'function') {
      ws.addConditionalFormatting({
        ref: `A${dataStartRow}:D${lastDataRow}`,
        rules: [
          {
            type: 'expression',
            formulae: [`$B${dataStartRow}<>""`],
            style: {
              border: {
                top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
                left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
                bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } },
                right: { style: 'thin', color: { argb: 'FFD1D5DB' } }
              }
            }
          }
        ]
      });
      ws.addConditionalFormatting({
        ref: `C${dataStartRow}:C${lastDataRow}`,
        rules: [
          {
            type: 'expression',
            formulae: [`AND($G${dataStartRow}<>"",INDEX($L:$L,$G${dataStartRow})=1)`],
            style: {
              fill: {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFFFD6D6' },
                bgColor: { argb: 'FFFFD6D6' }
              }
            }
          }
        ]
      });
    }

    return lastDataRow + 1;
  }

  function initialVisibleIndex(items, currentIndex, fixedCitySet) {
    let count = 0;

    for (let index = 0; index <= currentIndex; index++) {
      const item = items[index];

      if (!fixedCitySet || isCityInFixedSet(item.city, fixedCitySet)) {
        count += 1;
      }
    }

    return count;
  }

  function initialVisibleDateIndex(items, currentIndex, fixedCitySet, currentDate) {
    let count = 0;

    for (let index = 0; index <= currentIndex; index++) {
      const item = items[index];

      if (
        (!fixedCitySet || isCityInFixedSet(item.city, fixedCitySet)) &&
        isSameExcelDay(item.dayDate, currentDate)
      ) {
        count += 1;
      }
    }

    return count;
  }

  function writeListTable(ws, items, headerRow, dataBorder, withFilter, currentDate, options = {}) {
    const headers = ['Дата', 'Город', options.valueHeader || 'Осадки', ''];

    headers.forEach((header, index) => {
      const cell = ws.getCell(headerRow, index + 1);
      cell.value = header;
      cell.border = dataBorder;
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFF8FAFC' }
      };
      cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF111827' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    let row = headerRow + 1;

    if (!items.length) {
      ws.mergeCells(row, 1, row, 4);
      ws.getCell(row, 1).value = options.emptyText || 'Нет дней с осадками больше 5 мм';
      ws.getCell(row, 1).font = { name: 'Arial', size: 10, color: { argb: 'FF6B7280' } };
      ws.getCell(row, 1).alignment = { horizontal: 'left', vertical: 'middle' };
      ws.getCell(row, 1).border = dataBorder;
      row += 1;
    }

    for (const item of items) {
      const dayDate = item.dayDate || makeExcelDateFromForecastDay(item.day, item.dayIndex);
      ws.getCell(row, 1).value = dayDate || (item.dateText || formatShortDate(item.day, item.dayIndex));
      if (dayDate) {
        ws.getCell(row, 1).numFmt = 'dd.mm.yyyy';
      }
      ws.getCell(row, 2).value = item.city;
      ws.getCell(row, 3).value = typeof options.formatValue === 'function'
        ? options.formatValue(item)
        : makeRainListText(item.weather, item.rain);
      ws.getCell(row, 4).value = item.url
        ? {
          text: '🔗',
          hyperlink: item.url,
          tooltip: `Открыть прогноз: ${item.city}`
        }
        : '';

      for (let c = 1; c <= 4; c++) {
        const cell = ws.getCell(row, c);
        cell.border = dataBorder;
        if (c === 1 && dayDate) {
          cell.numFmt = 'dd.mm.yyyy';
        }
        cell.font = {
          name: c === 4 ? 'Segoe UI Emoji' : 'Arial',
          size: c === 4 ? 12 : 10,
          underline: false,
          color: { argb: c === 4 ? 'FF0563C1' : 'FF000000' }
        };
        cell.alignment = {
          horizontal: 'center',
          vertical: 'middle'
        };
        if (c === 4) {
          cell.style = {
            ...cell.style,
            font: {
              name: 'Segoe UI Emoji',
              size: 12,
              underline: false,
              color: { argb: 'FF0563C1' }
            }
          };
        }
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {
            argb: c === 3
              ? (typeof options.getValueFill === 'function' ? options.getValueFill(item) : options.valueFill) ||
                (row % 2 === 0 ? 'FFFFFFFF' : 'FFF8FAFC')
              : row % 2 === 0
                ? 'FFFFFFFF'
                : 'FFF8FAFC'
          }
        };
      }

      ws.getRow(row).hidden = Boolean(dayDate && currentDate && !isSameExcelDay(dayDate, currentDate));
      row += 1;
    }

    if (withFilter) {
      ws.autoFilter = {
        from: { row: headerRow, column: 1 },
        to: { row: Math.max(headerRow, row - 1), column: 4 }
      };
    }

    return row;
  }

  function makeRainListText(weather, rain) {
    const text = cleanWeatherText(weather || '').toLowerCase();
    const value = formatRain(rain);

    return text ? `${text}, ${value} мм` : `${value} мм`;
  }

  function makeWindListText(weather, wind) {
    const text = cleanWeatherText(weather || '').toLowerCase();
    const value = `${formatWind(wind)} м/с`;

    return text ? `${text}, ${value}` : value;
  }

  function makeRainGustListText(weather, rain, gust) {
    const text = cleanWeatherText(weather || '').toLowerCase();
    const value = `${formatRain(rain)} мм, ${formatWind(gust)} м/с`;

    return text ? `${text}, ${value}` : value;
  }

  function writeRainGustDatePicker(ws, osadkiDays, currentDay, dataBorder) {
    const dates = (osadkiDays?.length ? osadkiDays : buildDates())
      .map((day, index) => makeExcelDateFromForecastDay(day, index))
      .filter(Boolean);
    const currentDate = makeExcelDateFromForecastDay(currentDay, 0) || dates[0] || null;

    const dateCell = ws.getCell(4, 1);
    dateCell.value = currentDate;
    dateCell.numFmt = 'dd.mm.yyyy';
    dateCell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF111827' } };
    dateCell.alignment = { horizontal: 'center', vertical: 'middle' };
    dateCell.border = dataBorder;
    dateCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFF8FAFC' }
    };

    dates.forEach((date, index) => {
      const row = index + 5;
      const cell = ws.getCell(row, 16);
      cell.value = date;
      cell.numFmt = 'dd.mm.yyyy';
    });

    if (dates.length) {
      dateCell.dataValidation = {
        type: 'list',
        allowBlank: false,
        formulae: [`$P$5:$P$${dates.length + 4}`]
      };
    }
  }

  function writeFixedCityBox(ws, fixedCities, dataBorder) {
    ws.mergeCells(6, 1, 6, 3);
    ws.getCell(6, 1).value = 'Фиксированный список';
    ws.getCell(6, 1).font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF111827' } };
    ws.getCell(6, 1).alignment = { horizontal: 'center', vertical: 'middle' };

    for (let col = 1; col <= 3; col++) {
      ws.getCell(6, col).border = dataBorder;
    }

    const list = fixedCities.map(city => cleanCityName(city)).filter(Boolean);
    const rowsCount = list.length || 6;

    for (let index = 0; index < rowsCount; index++) {
      const row = index + 7;
      ws.mergeCells(row, 1, row, 3);
      const cell = ws.getCell(row, 1);
      cell.value = list[index] || '';
      cell.font = { name: 'Arial', size: 10, color: { argb: 'FF111827' } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };

      for (let col = 1; col <= 3; col++) {
        ws.getCell(row, col).border = dataBorder;
      }
    }
  }

  function buildFixedCitySet(fixedCities) {
    const normalized = fixedCities
      .map(city => normalizeLocation(city))
      .filter(Boolean);

    return normalized.length ? new Set(normalized) : null;
  }

  function mergeCitySets(...sets) {
    const merged = new Set();

    sets.filter(Boolean).forEach(set => {
      set.forEach(value => merged.add(value));
    });

    return merged.size ? merged : null;
  }

  function initialRainGustVisible(city, fixedCitySet, reportCitySet) {
    return fixedCitySet
      ? isCityInFixedSet(city, fixedCitySet)
      : true;
  }

  function buildFixedCityOrderMap(fixedCities) {
    const order = new Map();

    fixedCities.forEach((city, index) => {
      const normalized = normalizeLocation(city);

      if (normalized && !order.has(normalized)) {
        order.set(normalized, index);
      }
    });

    return order.size ? order : null;
  }

  function getFixedCityOrder(city, fixedCityOrder) {
    if (!fixedCityOrder) return Number.MAX_SAFE_INTEGER;

    const normalized = normalizeLocation(city);

    if (fixedCityOrder.has(normalized)) return fixedCityOrder.get(normalized);

    for (const [item, order] of fixedCityOrder.entries()) {
      if (
        normalized === item ||
        normalized.startsWith(`${item} `) ||
        normalized.startsWith(`${item}-`) ||
        item.startsWith(`${normalized} `) ||
        item.startsWith(`${normalized}-`)
      ) {
        return order;
      }
    }

    return Number.MAX_SAFE_INTEGER;
  }

  function isCityInFixedSet(city, fixedCitySet) {
    if (!fixedCitySet) return true;

    const normalized = normalizeLocation(city);

    if (fixedCitySet.has(normalized)) return true;

    return [...fixedCitySet].some(item =>
      normalized === item ||
      normalized.startsWith(`${item} `) ||
      normalized.startsWith(`${item}-`) ||
      item.startsWith(`${normalized} `) ||
      item.startsWith(`${normalized}-`)
    );
  }

  function formatShortDate(day, fallbackIndex = 0) {
    const date = makeExcelDateFromForecastDay(day, fallbackIndex);

    if (!date) return day;

    return [
      String(date.getDate()).padStart(2, '0'),
      String(date.getMonth() + 1).padStart(2, '0'),
      String(date.getFullYear())
    ].join('.');
  }

  function buildStrongRainRows(blocks, osadkiDays, citySet = null) {
    const items = [];
    const fallbackDays = buildDates();

    for (const block of blocks) {
      if (!isCityInFixedSet(block.city, citySet)) continue;

      for (let dayIndex = 0; dayIndex < 10; dayIndex++) {
        const rain = Number(block.rain[dayIndex]);

        if (!Number.isFinite(rain) || rain <= 5) continue;

        const day = osadkiDays?.[dayIndex] || block.days?.[dayIndex] || fallbackDays[dayIndex];
        const dayDate = makeExcelDateFromForecastDay(day, dayIndex);
        const dateText = formatShortDate(day, dayIndex);

        items.push({
          day,
          dayDate,
          dateText,
          sortKey: dayDate ? dayDate.getTime() : Number.MAX_SAFE_INTEGER + dayIndex,
          city: block.city,
          rain,
          weather: block.weather?.[dayIndex] || '',
          url: block.url,
          dayIndex
        });
      }
    }

    return items.sort((a, b) =>
      a.sortKey - b.sortKey ||
      String(a.city).localeCompare(String(b.city), 'ru')
    );
  }

  function buildRainGustRows(blocks, osadkiDays, fixedCitySet = null, fixedCityOrder = null) {
    const items = [];
    const fallbackDays = buildDates();

    for (const block of blocks) {
      if (!isCityInFixedSet(block.city, fixedCitySet)) continue;

      for (let dayIndex = 0; dayIndex < 10; dayIndex++) {
        const rain = Number(block.rain?.[dayIndex]);
        const gust = Number(block.gust?.[dayIndex]);
        const hasRain = Number.isFinite(rain) && rain > 0;
        const hasGust = Number.isFinite(gust) && gust > 0;

        if (!hasRain && !hasGust) continue;

        const day = osadkiDays?.[dayIndex] || block.days?.[dayIndex] || fallbackDays[dayIndex];
        const dayDate = makeExcelDateFromForecastDay(day, dayIndex);
        const dateText = formatShortDate(day, dayIndex);

        items.push({
          day,
          dayDate,
          dateText,
          sortKey: dayDate ? dayDate.getTime() : Number.MAX_SAFE_INTEGER + dayIndex,
          city: block.city,
          rain: Number.isFinite(rain) ? rain : 0,
          gust: Number.isFinite(gust) ? gust : 0,
          weather: block.weather?.[dayIndex] || '',
          url: block.url,
          dayIndex
        });
      }
    }

    return items.sort((a, b) =>
      a.sortKey - b.sortKey ||
      getFixedCityOrder(a.city, fixedCityOrder) - getFixedCityOrder(b.city, fixedCityOrder) ||
      String(a.city).localeCompare(String(b.city), 'ru')
    );
  }

  function makeExcelDateFromForecastDay(day, fallbackIndex) {
    const months = {
      января: 0,
      февраля: 1,
      марта: 2,
      апреля: 3,
      мая: 4,
      июня: 5,
      июля: 6,
      августа: 7,
      сентября: 8,
      октября: 9,
      ноября: 10,
      декабря: 11
    };
    const value = normalizeForecastText(day).toLowerCase();
    const monthNames = Object.keys(months).join('|');
    const match = value.match(new RegExp(`(\\d{1,2})\\s+(${monthNames})`, 'i'));

    if (!match) return null;

    const today = new Date();
    const month = months[match[2]];
    const date = Number(match[1]);

    if (!Number.isFinite(month) || !Number.isFinite(date)) return null;

    let year = today.getFullYear();
    const candidate = new Date(Date.UTC(year, month, date));

    if (fallbackIndex > 0 && candidate < new Date(today.getFullYear(), today.getMonth(), today.getDate() - 20)) {
      year += 1;
    }

    return new Date(Date.UTC(year, month, date));
  }

  function isSameExcelDay(a, b) {
    return a instanceof Date &&
      b instanceof Date &&
      a.getUTCFullYear() === b.getUTCFullYear() &&
      a.getUTCMonth() === b.getUTCMonth() &&
      a.getUTCDate() === b.getUTCDate();
  }

  async function applyCurrentDayFilters(buffer, filters) {
    let result = buffer;

    for (const filter of filters.filter(Boolean)) {
      result = await applyCurrentDayFilter(result, filter);
    }

    return result;
  }

  async function applyCurrentDayFilter(buffer, filter) {
    if (!filter?.currentDate || typeof JSZip === 'undefined') return buffer;

    try {
      const zip = await JSZip.loadAsync(buffer);
      const sheetPath = await findWorksheetPath(zip, filter.sheetName);
      if (!sheetPath) return buffer;

      const sheetFile = zip.file(sheetPath);
      if (!sheetFile) return buffer;

      let xml = await sheetFile.async('string');
      const ref = escapeXmlAttr(filter.ref);
      const year = filter.currentDate.getUTCFullYear();
      const month = filter.currentDate.getUTCMonth() + 1;
      const day = filter.currentDate.getUTCDate();
      const autoFilterXml = `<autoFilter ref="${ref}"><filterColumn colId="0"><filters><dateGroupItem year="${year}" month="${month}" day="${day}" dateTimeGrouping="day"/></filters></filterColumn></autoFilter>`;

      if (/<autoFilter\b[^>]*\/>/.test(xml)) {
        xml = xml.replace(/<autoFilter\b[^>]*\/>/, autoFilterXml);
      } else if (/<autoFilter\b[\s\S]*?<\/autoFilter>/.test(xml)) {
        xml = xml.replace(/<autoFilter\b[\s\S]*?<\/autoFilter>/, autoFilterXml);
      } else if (xml.includes('<pageMargins')) {
        xml = xml.replace('<pageMargins', `${autoFilterXml}<pageMargins`);
      } else {
        xml = xml.replace('</worksheet>', `${autoFilterXml}</worksheet>`);
      }

      zip.file(sheetPath, xml);

      return await zip.generateAsync({ type: 'arraybuffer' });
    } catch (e) {
      console.warn('Не удалось применить фильтр текущего дня', e);
      return buffer;
    }
  }

  async function findWorksheetPath(zip, sheetName) {
    const workbookFile = zip.file('xl/workbook.xml');
    const relsFile = zip.file('xl/_rels/workbook.xml.rels');

    if (!workbookFile || !relsFile) return null;

    const workbookXml = await workbookFile.async('string');
    const relsXml = await relsFile.async('string');
    const escapedName = escapeRegExp(sheetName);
    const sheetMatch = workbookXml.match(new RegExp(`<sheet\\b[^>]*name="${escapedName}"[^>]*r:id="([^"]+)"`, 'i'));

    if (!sheetMatch) return null;

    const relMatch = relsXml.match(new RegExp(`<Relationship\\b[^>]*Id="${escapeRegExp(sheetMatch[1])}"[^>]*Target="([^"]+)"`, 'i'));

    if (!relMatch) return null;

    const target = relMatch[1].replace(/^\/+/, '');

    return target.startsWith('xl/')
      ? target
      : `xl/${target}`;
  }

  function escapeXmlAttr(value) {
    return String(value)
      .replace(/&/g, '&amp;')
      .replace(/"/g, '&quot;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  function escapeRegExp(value) {
    return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function getOrderedDays(blocks) {
    const first = blocks.find(block => block.days?.length)?.days || buildDates();

    return first.slice(0, 10);
  }

  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
})();
