// ==UserScript==
// @name         Gismeteo Precipitation
// @namespace    gismeteo-excel
// @version      1.2
// @description  Export Gismeteo 10-day precipitation forecasts to a styled Excel report with daily charts and a filtered heavy-rain list.
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
  const BASE = 'https://www.gismeteo.ru';
  const LOCATION_WORDS_RE = /^(?:погода\s+)?(?:в|во|на|для)\s+/i;
  const FORECAST_TAIL_RE = /\s+(?:на\s+(?:10\s+дней|3\s+дня|2\s+недели|месяц|неделю|выходные)|сегодня|завтра).*$/i;
  const BAD_LOCATION_RE = /аэропорт|airport|аэродром|aeroport|авиабаза|внуково|шереметьево|домодедово|спиченково|остафьево/i;
  const WEATHER_TEXT_RE = /\b(?:ясно|облачно|пасмурно|дождь|дождя|дождем|дождём|дожди|гроз[а-я]*|снег[а-я]*|туман[а-я]*|град|ливень|ливни|морось|переменная)\b/i;
  let collapsed = true;

  createPanel();

  function createPanel() {
    const style = document.createElement('style');
    style.textContent = `
      #gmRun {
        transition: transform .16s ease, background-color .16s ease, box-shadow .16s ease;
        box-shadow: 0 6px 14px rgba(22, 163, 74, .24);
      }
      #gmRun:hover {
        background: #22c55e !important;
        transform: translateY(-1px);
        box-shadow: 0 9px 20px rgba(34, 197, 94, .32);
      }
      #gmRun:active {
        transform: translateY(0) scale(.97);
        box-shadow: 0 4px 10px rgba(22, 163, 74, .24);
      }
    `;
    document.head.appendChild(style);

    const box = document.createElement('div');
    box.id = 'gmBox';
    box.style = `
      position: fixed; right: 18px; bottom: 18px; width: 176px;
      background: #111827; z-index: 999999; border-radius: 10px;
      border: 1px solid #263244;
      box-shadow: 0 14px 34px rgba(0,0,0,.38);
      font-family: Arial, sans-serif; color: #e5e7eb; overflow: hidden;
    `;

    box.innerHTML = `
      <div id="gmHeader" style="
        background:#0f172a;color:#f8fafc;padding:8px 10px;cursor:pointer;
        border-bottom:1px solid #263244;
        font-size:13px;font-weight:bold;display:flex;justify-content:space-between;
      ">
        <span>Gismeteo Precipitation</span><span id="gmToggle">＋</span>
      </div>

      <div id="gmBody" style="display:none;padding:8px;">
        <textarea id="gmCities" style="
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
    textarea.value = localStorage.getItem(STORAGE_KEY) || localStorage.getItem('gm_city_list_v44') || '';

    document.getElementById('gmHeader').onclick = () => {
      collapsed = !collapsed;
      document.getElementById('gmBody').style.display = collapsed ? 'none' : 'block';
      document.getElementById('gmToggle').textContent = collapsed ? '＋' : '−';
      box.style.width = collapsed ? '176px' : '306px';
    };

    document.getElementById('gmRun').onclick = async () => {
      const cities = textarea.value.split('\n').map(x => x.trim()).filter(Boolean);
      localStorage.setItem(STORAGE_KEY, textarea.value);

      if (!cities.length) {
        alert('Вставь список городов');
        return;
      }

      await run(cities);
    };
  }

  async function run(cities) {
    const btn = document.getElementById('gmRun');
    const status = document.getElementById('gmStatus');

    btn.disabled = true;
    btn.textContent = 'Сбор...';

    const blocks = [];

    for (let i = 0; i < cities.length; i++) {
      const rawCity = cities[i];

      try {
        status.textContent = `${i + 1}/${cities.length}: поиск ${rawCity}`;

        const found = await findCityUrl(rawCity);

        status.textContent = `${i + 1}/${cities.length}: прогноз ${found.name}`;

        const forecastUrl = make10DaysUrl(found.url);
        const parsed = await loadForecastWithRetries(forecastUrl, 3);
        const rainChart = makeRainChartImage(found.name || rawCity, parsed);

        blocks.push({
          city: found.name || rawCity,
          url: forecastUrl,
          days: parsed.days,
          rain: parsed.rain,
          weather: parsed.weather,
          rainChart,
          status: 'OK'
        });
      } catch (e) {
        console.error(e);

        blocks.push({
          city: rawCity,
          url: '',
          days: buildDates(),
          rain: Array(10).fill(''),
          weather: Array(10).fill(''),
          rainChart: null,
          status: 'FAIL'
        });
      }

      await sleep(1200);
    }

    status.textContent = 'Формирую Excel...';
    await exportExcel(blocks);

    status.textContent = `Готово: ${cities.length}`;
    btn.disabled = false;
    btn.textContent = 'Собрать';
  }

  async function findCityUrl(city) {
    const searchUrl = `${BASE}/search/${encodeURIComponent(city)}/`;
    const loaded = await loadPageWithIframe(searchUrl, 9000);
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

    const strictCandidates = allCandidates
      .filter(x => isStrictCityMatch(query, x.normalizedLocation))
      .sort((a, b) => b.score - a.score);

    const candidates = strictCandidates.length
      ? strictCandidates
      : allCandidates
        .filter(x => isWeakCityMatch(query, x.normalizedLocation))
        .sort((a, b) => b.score - a.score);

    loaded.iframe.remove();

    if (!candidates.length) throw new Error(`город не найден: ${city}`);

    return candidates[0];
  }

  function buildCityCandidate(a, query) {
    const href = a.getAttribute('href') || '';
    const rawText = a.innerText || a.textContent || '';
    const name = cleanCityName(rawText);
    const location = extractLocationName(name);
    const normalizedLocation = normalizeLocation(location || name);

    if (!href || !normalizedLocation) return null;

    return {
      name: location || name,
      href,
      url: href.startsWith('http') ? href : BASE + href,
      normalizedLocation,
      score: scoreCity(query, normalizedLocation, href, name)
    };
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
    return BAD_LOCATION_RE.test(`${candidate.name} ${candidate.href} ${candidate.normalizedLocation}`);
  }

  function isStrictCityMatch(query, candidate) {
    if (!query || !candidate) return false;
    if (candidate === query) return true;
    if (candidate.startsWith(`${query} `) || candidate.startsWith(`${query}-`)) return true;

    return candidate
      .split(/[\s-]+/)
      .filter(Boolean)
      .join(' ') === query;
  }

  function isWeakCityMatch(query, candidate) {
    if (!query || !candidate) return false;
    if (isStrictCityMatch(query, candidate)) return true;

    const queryWords = query.split(/[\s-]+/).filter(Boolean);
    const candidateWords = candidate.split(/[\s-]+/).filter(Boolean);

    return queryWords.length > 0 && queryWords.every(word =>
      candidateWords.some(candidateWord => candidateWord.startsWith(word))
    );
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

  async function loadForecastWithRetries(url, attempts = 3) {
    let lastError = null;

    for (let attempt = 1; attempt <= attempts; attempt++) {
      let loaded = null;

      try {
        loaded = await loadPageWithIframe(url, 15000);
        await waitForForecastReady(loaded.doc, 12000);

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

    return {
      days: extractForecastDays(doc, text),
      rain,
      weather: extractWeatherTexts(doc)
    };
  }

  function extractRainValuesFromDom(doc) {
    const rows = [
      ...doc.querySelectorAll('[data-row*="precipitation"], .widget-row-chart-precipitation')
    ];

    for (const row of rows) {
      const values = [...row.querySelectorAll('.row-item')]
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

        if ((hasWeather && hasRain) || Date.now() - started >= timeout) {
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

    const values = parsed.rain.map(value => {
      const number = Number(value);
      return Number.isFinite(number) ? number : 0;
    });
    const maxValue = Math.max(5, ...values);
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

    values.forEach((value, index) => {
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

  async function exportExcel(blocks) {
    const workbook = new ExcelJS.Workbook();
    const ws = workbook.addWorksheet('Осадки', {
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

    let row = 3;
    const blockGapRows = 2;

    for (const block of blocks) {
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

      ws.getRow(cityRow).height = 22;
      ws.getRow(rainRow).height = 22;

      ws.getCell(cityRow, 1).value = block.city;
      ws.getCell(rainRow, 1).value = 'Осадки, мм';

      for (let i = 0; i < 10; i++) {
        const rainNumber = Number(block.rain[i]);

        ws.getCell(cityRow, i + 2).value = block.days[i];
        ws.getCell(rainRow, i + 2).value = Number.isFinite(rainNumber) ? rainNumber : '';
        ws.getCell(rainRow, i + 2).numFmt = '0.0';
      }

      ws.mergeCells(cityRow, 12, rainRow, 12);
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
                ? 'FFF4FBF6'
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
      }

      row += 4;

      const visualTopRow = row - 1;
      let visualHeight = 0;

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

        visualHeight = Math.max(visualHeight, chartHeight);
      }

      if (visualHeight > 0) {
        row += Math.max(1, Math.ceil(visualHeight / 20) - 1);
      }

      row += blockGapRows;
    }

    const osadkiDays = getOrderedDays(blocks);
    const listFilter = buildListSheet(workbook, blocks, dataBorder, osadkiDays);

    let buffer = await workbook.xlsx.writeBuffer();
    buffer = await applyCurrentDayFilter(buffer, listFilter);

    saveAs(
      new Blob([buffer]),
      `gismeteo_osadki_${new Date().toISOString().slice(0, 10)}.xlsx`
    );
  }

  function buildListSheet(workbook, blocks, dataBorder, osadkiDays) {
    const ws = workbook.addWorksheet('Список', {
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

    const listRows = buildStrongRainRows(blocks, osadkiDays);
    const currentDay = osadkiDays[0] || buildDates()[0];
    const currentDate = makeExcelDateFromForecastDay(currentDay, 0);

    ws.getCell(2, 1).value = `По умолчанию показан текущий день: ${formatShortDate(currentDay)}. Выберите другую дату через фильтр.`;
    ws.getCell(2, 1).font = { name: 'Arial', size: 10, color: { argb: 'FF6B7280' } };

    const lastRow = writeListTable(ws, listRows, 4, dataBorder, true, currentDate);

    return {
      sheetName: 'Список',
      ref: `A4:D${Math.max(4, lastRow - 1)}`,
      currentDate
    };
  }

  function writeListTable(ws, items, headerRow, dataBorder, withFilter, currentDate) {
    const headers = ['Дата', 'Город', 'Осадки', ''];

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
      ws.getCell(row, 1).value = 'Нет дней с осадками больше 5 мм';
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
      ws.getCell(row, 3).value = makeRainListText(item.weather, item.rain);
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
              ? 'FFFFD6D6'
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

  function formatShortDate(day, fallbackIndex = 0) {
    const date = makeExcelDateFromForecastDay(day, fallbackIndex);

    if (!date) return day;

    return [
      String(date.getDate()).padStart(2, '0'),
      String(date.getMonth() + 1).padStart(2, '0'),
      String(date.getFullYear())
    ].join('.');
  }

  function buildStrongRainRows(blocks, osadkiDays) {
    const items = [];
    const fallbackDays = buildDates();

    for (const block of blocks) {
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
