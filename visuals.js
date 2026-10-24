/* Local-only OOXML drawing reader and SVG chart renderer for Visor XLSX. */
(() => {
  'use strict';
  const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const SVG_NS = 'http://www.w3.org/2000/svg';
  const EMU_PER_PIXEL = 9525;
  const MAX_MEDIA_BYTES = 48 * 1024 * 1024;
  const MAX_OBJECTS_PER_SHEET = 300;
  const COLORS = ['#3578c6','#22a06b','#e08b27','#8b61c2','#d14b58','#16a0a6','#6d7b89','#cf6eaa'];

  function fileText(wb, path) {
    const file = wb.files && wb.files[path];
    if (!file || !file.content) return '';
    return typeof file.content === 'string' ? file.content : new TextDecoder('utf-8').decode(file.content);
  }

  function normalizePath(path) {
    const normalized = [];
    for (const part of String(path || '').split('/')) {
      if (!part || part === '.') continue;
      if (part === '..') normalized.pop();
      else normalized.push(part);
    }
    return normalized.join('/');
  }

  function resolveTarget(sourcePart, target) {
    if (!target || /^[a-z]+:/i.test(target)) return '';
    if (target[0] === '/') return normalizePath(target.slice(1));
    const base = sourcePart.split('/').slice(0, -1).join('/');
    return normalizePath((base ? base + '/' : '') + target);
  }

  function relationshipPart(part) {
    const slash = part.lastIndexOf('/');
    const directory = slash < 0 ? '' : part.slice(0, slash + 1);
    const name = part.slice(slash + 1);
    return directory + '_rels/' + name + '.rels';
  }

  function parseXml(wb, path) {
    const xml = fileText(wb, path);
    if (!xml || typeof DOMParser === 'undefined') return null;
    try {
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      return doc.getElementsByTagName('parsererror').length ? null : doc;
    } catch (_) { return null; }
  }

  function allElements(node, localName) {
    if (!node || !node.getElementsByTagName) return [];
    return Array.from(node.getElementsByTagName('*')).filter((el) => el.localName === localName);
  }

  function firstElement(node, localName) {
    return allElements(node, localName)[0] || null;
  }

  function childElement(node, localName) {
    if (!node) return null;
    return Array.from(node.children || []).find((el) => el.localName === localName) || null;
  }

  function nodeText(node) {
    return node ? String(node.textContent || '').trim() : '';
  }

  function relationshipId(node, name) {
    if (!node) return '';
    return node.getAttributeNS(REL_NS, name) || node.getAttribute('r:' + name) || '';
  }

  function relationships(wb, sourcePart) {
    const doc = parseXml(wb, relationshipPart(sourcePart));
    const result = new Map();
    if (!doc) return result;
    for (const relation of allElements(doc, 'Relationship')) {
      const id = relation.getAttribute('Id');
      const target = relation.getAttribute('Target');
      const external = relation.getAttribute('TargetMode') === 'External';
      if (id && target && !external) result.set(id, {
        path: resolveTarget(sourcePart, target),
        type: relation.getAttribute('Type') || ''
      });
    }
    return result;
  }

  function marker(anchor, name) {
    const root = childElement(anchor, name);
    if (!root) return null;
    const number = (key) => {
      const value = Number(nodeText(childElement(root, key)));
      return Number.isFinite(value) ? value : 0;
    };
    return { col:number('col'), row:number('row'), colOff:number('colOff'), rowOff:number('rowOff') };
  }

  function anchorInfo(anchor) {
    const kind = anchor.localName;
    if (kind === 'absoluteAnchor') {
      const pos = childElement(anchor, 'pos');
      const ext = childElement(anchor, 'ext');
      return {
        absolute: true,
        x: Number(pos && pos.getAttribute('x')) / EMU_PER_PIXEL || 0,
        y: Number(pos && pos.getAttribute('y')) / EMU_PER_PIXEL || 0,
        width: Number(ext && ext.getAttribute('cx')) / EMU_PER_PIXEL || 320,
        height: Number(ext && ext.getAttribute('cy')) / EMU_PER_PIXEL || 220
      };
    }
    const from = marker(anchor, 'from') || { col:0, row:0, colOff:0, rowOff:0 };
    const to = marker(anchor, 'to');
    const ext = childElement(anchor, 'ext');
    return {
      absolute: false,
      from,
      to,
      width: Number(ext && ext.getAttribute('cx')) / EMU_PER_PIXEL || 0,
      height: Number(ext && ext.getAttribute('cy')) / EMU_PER_PIXEL || 0
    };
  }

  function chartFormulaValues(formula, wb, defaultSheetName) {
    const source = String(formula || '').replace(/^=/, '').trim();
    const match = source.match(/^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]{1,3}\$?\d+)(?::\$?([A-Z]{1,3}\$?\d+))?$/i);
    if (!match) return [];
    const sheetName = (match[1] ? match[1].replace(/''/g, "'") : match[2]).replace(/^\[[^\]]+\]/, '');
    const ws = wb.Sheets[sheetName] || (sheetName === defaultSheetName ? wb.Sheets[defaultSheetName] : null);
    if (!ws) return [];
    const start = match[3].replace(/\$/g, '');
    const end = (match[4] || match[3]).replace(/\$/g, '');
    let range;
    try { range = XLSX.utils.decode_range(start + ':' + end); } catch (_) { return []; }
    const values = [];
    for (let row = range.s.r; row <= range.e.r; row++) {
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cell = ws[XLSX.utils.encode_cell({ r:row, c:col })];
        values.push(cell ? (cell.t === 'd' ? (cell.w || cell.v) : (cell.v == null ? (cell.w || '') : cell.v)) : '');
      }
    }
    return values;
  }

  function chartData(node, wb, defaultSheetName) {
    if (!node) return [];
    for (const name of ['strCache','numCache','strLit','numLit']) {
      const cache = firstElement(node, name);
      if (!cache) continue;
      const points = allElements(cache, 'pt');
      if (!points.length) {
        const value = firstElement(cache, 'v');
        return value ? [nodeText(value)] : [];
      }
      const output = [];
      for (const point of points) {
        const value = firstElement(point, 'v');
        if (!value) continue;
        const index = Number(point.getAttribute('idx'));
        output[Number.isInteger(index) && index >= 0 ? index : output.length] = nodeText(value);
      }
      return output;
    }
    const formula = firstElement(node, 'f');
    return formula ? chartFormulaValues(nodeText(formula), wb, defaultSheetName) : [];
  }

  function chartLabel(node, wb, defaultSheetName, fallback) {
    const values = chartData(node, wb, defaultSheetName);
    if (values.length) return String(values[0] == null ? fallback : values[0]);
    const literal = firstElement(node, 'v') || firstElement(node, 't');
    return nodeText(literal) || fallback;
  }

  function readChart(wb, path, sheetName) {
    const doc = parseXml(wb, path);
    if (!doc) return null;
    const title = allElements(firstElement(doc, 'title') || doc, 't').map(nodeText).filter(Boolean).join(' ');
    const plot = firstElement(doc, 'plotArea') || doc;
    const chartTypeNames = ['barChart','lineChart','areaChart','pieChart','doughnutChart','ofPieChart','scatterChart','radarChart','bubbleChart','stockChart','surfaceChart','histogramChart','boxWhiskerChart','funnelChart','waterfallChart','treemapChart','sunburstChart'];
    const chartNode = chartTypeNames.map((name) => firstElement(plot, name)).find(Boolean);
    if (!chartNode) return { title:title || 'Gráfico', type:'unsupported', series:[] };
    let type = chartNode.localName.replace('Chart', '').toLowerCase();
    if (type === 'bar') {
      const direction = firstElement(chartNode, 'barDir');
      type = direction && direction.getAttribute('val') === 'bar' ? 'horizontalBar' : 'bar';
    }
    const series = allElements(chartNode, 'ser').map((ser, index) => {
      const tx = childElement(ser, 'tx');
      const cat = childElement(ser, 'cat') || childElement(ser, 'xVal');
      const val = childElement(ser, 'val') || childElement(ser, 'yVal');
      return {
        name: chartLabel(tx, wb, sheetName, 'Serie ' + (index + 1)),
        categories: chartData(cat, wb, sheetName),
        values: chartData(val, wb, sheetName)
      };
    }).filter((item) => item.values.length);
    return { title:title || 'Gráfico', type, series };
  }

  function objectName(anchor, fallback) {
    const properties = firstElement(anchor, 'cNvPr');
    return properties && (properties.getAttribute('name') || properties.getAttribute('descr')) || fallback;
  }

  function imageMime(path) {
    const extension = path.split('.').pop().toLowerCase();
    return ({ png:'image/png', jpg:'image/jpeg', jpeg:'image/jpeg', gif:'image/gif', bmp:'image/bmp', webp:'image/webp', svg:'image/svg+xml' })[extension] || '';
  }

  function mediaBytes(file) {
    const content = file && file.content;
    if (!content) return 0;
    if (typeof content === 'string') return content.length;
    return Number(content.byteLength || content.length || 0);
  }

  function readAnchorObject(wb, anchor, anchorMap, drawingRels, defaultName, mediaState) {
    const picture = firstElement(anchor, 'pic');
    if (picture) {
      const blip = firstElement(picture, 'blip');
      const relation = blip && (relationshipId(blip, 'embed') || relationshipId(blip, 'link'));
      const target = drawingRels.get(relation);
      const file = target && wb.files && wb.files[target.path];
      const mime = target && imageMime(target.path);
      if (target && file && mime) {
        const bytes = mediaBytes(file);
        if (bytes && mediaState.totalBytes + bytes <= MAX_MEDIA_BYTES) {
          try {
            const url = URL.createObjectURL(new Blob([file.content], { type:mime }));
            mediaState.urls.push(url);
            mediaState.totalBytes += bytes;
            return { kind:'image', name:objectName(picture, defaultName), src:url, anchor:anchorMap };
          } catch (_) { /* Keep a visible fallback card below. */ }
        }
        return { kind:'unsupported', name:objectName(picture, defaultName), message:bytes > MAX_MEDIA_BYTES ? 'Imagen omitida por tamaño' : 'No se pudo cargar la imagen', anchor:anchorMap };
      }
      return { kind:'unsupported', name:objectName(picture, defaultName), message:'Formato de imagen no compatible', anchor:anchorMap };
    }

    const chartReference = firstElement(anchor, 'chart');
    if (chartReference) {
      const target = drawingRels.get(relationshipId(chartReference, 'id'));
      const chart = target && readChart(wb, target.path, defaultName);
      return { kind:'chart', name:chart && chart.title || objectName(anchor, 'Gráfico'), chart:chart || { title:'Gráfico', type:'unsupported', series:[] }, anchor:anchorMap };
    }

    const shape = firstElement(anchor, 'sp');
    if (shape) {
      const text = allElements(firstElement(shape, 'txBody') || shape, 't').map(nodeText).filter(Boolean).join('\n');
      if (text) return { kind:'text', name:objectName(shape, defaultName), text, anchor:anchorMap };
    }

    const frame = firstElement(anchor, 'graphicFrame');
    if (frame) return { kind:'unsupported', name:objectName(frame, defaultName), message:'Objeto de Excel no compatible', anchor:anchorMap };
    return null;
  }

  function extract(wb) {
    const bySheet = new Map();
    const urls = [];
    const mediaState = { totalBytes:0, urls };
    try {
      const sheetEntries = wb.Workbook && wb.Workbook.Sheets ? wb.Workbook.Sheets : [];
      const workbookRels = relationships(wb, 'xl/workbook.xml');
      for (let index = 0; index < sheetEntries.length; index++) {
        const metadata = sheetEntries[index];
        const sheetName = metadata.name || wb.SheetNames[index];
        const sheetRel = workbookRels.get(metadata.id);
        if (!sheetRel) continue;
        const sheetDoc = parseXml(wb, sheetRel.path);
        if (!sheetDoc) continue;
        const sheetRels = relationships(wb, sheetRel.path);
        const drawings = allElements(sheetDoc, 'drawing');
        const objects = [];
        for (const drawingReference of drawings) {
          const drawingRel = sheetRels.get(relationshipId(drawingReference, 'id'));
          if (!drawingRel) continue;
          const drawingDoc = parseXml(wb, drawingRel.path);
          if (!drawingDoc) continue;
          const drawingRels = relationships(wb, drawingRel.path);
          const anchors = allElements(drawingDoc, 'twoCellAnchor')
            .concat(allElements(drawingDoc, 'oneCellAnchor'), allElements(drawingDoc, 'absoluteAnchor'));
          for (const anchor of anchors) {
            if (objects.length >= MAX_OBJECTS_PER_SHEET) break;
            const map = anchorInfo(anchor);
            const object = readAnchorObject(wb, anchor, map, drawingRels, sheetName, mediaState);
            if (object) objects.push(object);
          }
        }
        if (objects.length) bySheet.set(sheetName, objects);
      }
    } catch (error) {
      for (const url of urls) URL.revokeObjectURL(url);
      throw error;
    }
    return { bySheet, urls, count:Array.from(bySheet.values()).reduce((total, items) => total + items.length, 0) };
  }

  function svgElement(parent, name, attrs, text) {
    const element = document.createElementNS(SVG_NS, name);
    for (const [key, value] of Object.entries(attrs || {})) element.setAttribute(key, String(value));
    if (text != null) element.textContent = String(text);
    if (parent) parent.appendChild(element);
    return element;
  }

  function numeric(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
    const normalized = String(value == null ? '' : value).trim().replace(/\s/g, '').replace(',', '.');
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function shortLabel(value, maxLength) {
    const text = String(value == null ? '' : value);
    return text.length > maxLength ? text.slice(0, maxLength - 1) + '…' : text;
  }

  function drawLegend(svg, series) {
    let x = 72;
    for (let i = 0; i < Math.min(series.length, 5); i++) {
      svgElement(svg, 'rect', { x, y:333, width:10, height:10, rx:2, fill:COLORS[i % COLORS.length] });
      svgElement(svg, 'text', { x:x + 15, y:342, fill:'#42505e', 'font-size':11 }, shortLabel(series[i].name, 18));
      x += Math.min(185, 40 + String(series[i].name).length * 7);
    }
  }

  function chartFrame(chart) {
    const svg = document.createElementNS(SVG_NS, 'svg');
    svg.setAttribute('class', 'chart-svg');
    svg.setAttribute('viewBox', '0 0 640 360');
    svg.setAttribute('role', 'img');
    svg.setAttribute('aria-label', chart.title || 'Gráfico');
    svgElement(svg, 'rect', { x:0, y:0, width:640, height:360, fill:'#fff' });
    svgElement(svg, 'text', { x:28, y:31, fill:'#253342', 'font-size':17, 'font-weight':700 }, shortLabel(chart.title || 'Gráfico', 60));
    return svg;
  }

  function renderDataFallback(svg, chart) {
    const series = chart.series || [];
    const message = chart.type === 'unsupported'
      ? 'Este tipo de gráfico no se puede dibujar en el visor.'
      : (series.length ? 'Datos guardados en el gráfico' : 'Excel no guardó datos de vista previa para este gráfico.');
    svgElement(svg, 'text', { x:28, y:58, fill:'#687684', 'font-size':12 }, message);
    let y = 86;
    for (let s = 0; s < series.length && s < 4; s++) {
      const item = series[s];
      svgElement(svg, 'text', { x:28, y, fill:COLORS[s % COLORS.length], 'font-size':12, 'font-weight':700 }, shortLabel(item.name, 34));
      y += 18;
      for (let i = 0; i < item.values.length && i < 10; i++) {
        const label = item.categories[i] == null ? String(i + 1) : String(item.categories[i]);
        svgElement(svg, 'text', { x:44, y, fill:'#42505e', 'font-size':11 }, shortLabel(label, 35) + '  ·  ' + String(item.values[i]));
        y += 16;
      }
      y += 7;
    }
    return svg;
  }

  function renderCartesian(svg, chart, horizontal) {
    const series = chart.series || [];
    if (!series.length || !series[0].values.length) return renderDataFallback(svg, chart);
    const first = series[0];
    const count = Math.max(1, first.categories.length, ...series.map((item) => item.values.length));
    const values = series.flatMap((item) => item.values.map(numeric));
    const max = Math.max(0, ...values);
    const min = Math.min(0, ...values);
    const domainMax = max === min ? max + 1 : max;
    const domain = Math.max(1, domainMax - min);
    const plot = { x:62, y:58, w:534, h:244 };
    for (let step = 0; step <= 4; step++) {
      const y = plot.y + plot.h - plot.h * step / 4;
      svgElement(svg, 'line', { x1:plot.x, y1:y, x2:plot.x + plot.w, y2:y, stroke:'#e5e9ed', 'stroke-width':1 });
      const label = (min + (domainMax - min) * step / 4).toFixed(Math.abs(domainMax - min) < 10 ? 1 : 0);
      svgElement(svg, 'text', { x:plot.x - 9, y:y + 4, fill:'#778390', 'font-size':10, 'text-anchor':'end' }, label);
    }
    if (horizontal) {
      const slot = plot.h / count;
      const group = Math.max(1, series.length);
      const zeroX = plot.x + plot.w * (0 - min) / domain;
      for (let category = 0; category < count; category++) {
        const label = first.categories[category] == null ? String(category + 1) : first.categories[category];
        const y = plot.y + category * slot + slot / 2;
        svgElement(svg, 'text', { x:plot.x - 9, y:y + 3, fill:'#65717e', 'font-size':10, 'text-anchor':'end' }, shortLabel(label, 12));
        series.forEach((item, si) => {
          const value = numeric(item.values[category] || 0);
          const barH = Math.max(2, slot * .72 / group);
          const valueX = plot.x + plot.w * (value - min) / domain;
          svgElement(svg, 'rect', { x:Math.min(zeroX, valueX), y:y - slot * .36 + si * barH, width:Math.max(1, Math.abs(valueX - zeroX)), height:barH - 2, rx:2, fill:COLORS[si % COLORS.length] });
        });
      }
    } else {
      const slot = plot.w / count;
      const groupW = slot * .76;
      const barW = Math.max(2, groupW / Math.max(1, series.length));
      const zeroY = plot.y + plot.h - plot.h * (0 - min) / domain;
      for (let category = 0; category < count; category++) {
        const label = first.categories[category] == null ? String(category + 1) : first.categories[category];
        const x = plot.x + category * slot + slot / 2;
        if (category % Math.max(1, Math.ceil(count / 12)) === 0) svgElement(svg, 'text', { x, y:plot.y + plot.h + 18, fill:'#65717e', 'font-size':10, 'text-anchor':'middle' }, shortLabel(label, 12));
        series.forEach((item, si) => {
          const value = numeric(item.values[category] || 0);
          const valueY = plot.y + plot.h - plot.h * (value - min) / domain;
          svgElement(svg, 'rect', { x:x - groupW / 2 + si * barW, y:Math.min(valueY, zeroY), width:Math.max(2, barW - 3), height:Math.max(1, Math.abs(zeroY - valueY)), rx:2, fill:COLORS[si % COLORS.length] });
        });
      }
    }
    drawLegend(svg, series);
    return svg;
  }

  function renderLine(svg, chart) {
    const series = chart.series || [];
    if (!series.length || !series[0].values.length) return renderDataFallback(svg, chart);
    const count = Math.max(1, ...series.map((item) => item.values.length));
    const values = series.flatMap((item) => item.values.map(numeric));
    const min = Math.min(0, ...values);
    const rawMax = Math.max(0, ...values);
    const max = rawMax === min ? min + 1 : rawMax;
    const plot = { x:60, y:58, w:540, h:244 };
    for (let step = 0; step <= 4; step++) {
      const y = plot.y + plot.h - plot.h * step / 4;
      svgElement(svg, 'line', { x1:plot.x, y1:y, x2:plot.x + plot.w, y2:y, stroke:'#e5e9ed', 'stroke-width':1 });
    }
    series.forEach((item, si) => {
      const coords = item.values.map((value, index) => ({
        x:plot.x + (count <= 1 ? plot.w / 2 : index * plot.w / (count - 1)),
        y:plot.y + plot.h - plot.h * (numeric(value) - min) / Math.max(1, max - min)
      }));
      const d = coords.map((point, index) => (index ? 'L' : 'M') + point.x + ' ' + point.y).join(' ');
      svgElement(svg, 'path', { d, fill:'none', stroke:COLORS[si % COLORS.length], 'stroke-width':3, 'stroke-linecap':'round', 'stroke-linejoin':'round' });
      coords.forEach((point) => svgElement(svg, 'circle', { cx:point.x, cy:point.y, r:3.5, fill:'#fff', stroke:COLORS[si % COLORS.length], 'stroke-width':2 }));
    });
    const cats = series[0].categories || [];
    for (let i = 0; i < count; i += Math.max(1, Math.ceil(count / 12))) {
      const x = plot.x + (count <= 1 ? plot.w / 2 : i * plot.w / (count - 1));
      svgElement(svg, 'text', { x, y:plot.y + plot.h + 18, fill:'#65717e', 'font-size':10, 'text-anchor':'middle' }, shortLabel(cats[i] == null ? String(i + 1) : cats[i], 12));
    }
    drawLegend(svg, series);
    return svg;
  }

  function renderPie(svg, chart, doughnut) {
    const item = chart.series && chart.series[0];
    if (!item || !item.values.length) return renderDataFallback(svg, chart);
    const positive = item.values.map(numeric).map((value) => Math.max(0, value));
    const total = positive.reduce((sum, value) => sum + value, 0);
    if (!total) return renderDataFallback(svg, chart);
    const cx = 285, cy = 190, radius = 126, inner = doughnut ? 66 : 0;
    let angle = -Math.PI / 2;
    positive.forEach((value, index) => {
      const delta = value / total * Math.PI * 2;
      const next = angle + delta;
      const x1 = cx + radius * Math.cos(angle), y1 = cy + radius * Math.sin(angle);
      const x2 = cx + radius * Math.cos(next), y2 = cy + radius * Math.sin(next);
      const large = delta > Math.PI ? 1 : 0;
      let d;
      if (inner) {
        const ix1 = cx + inner * Math.cos(angle), iy1 = cy + inner * Math.sin(angle);
        const ix2 = cx + inner * Math.cos(next), iy2 = cy + inner * Math.sin(next);
        d = `M ${x1} ${y1} A ${radius} ${radius} 0 ${large} 1 ${x2} ${y2} L ${ix2} ${iy2} A ${inner} ${inner} 0 ${large} 0 ${ix1} ${iy1} Z`;
      } else d = `M ${cx} ${cy} L ${x1} ${y1} A ${radius} ${radius} 0 ${large} 1 ${x2} ${y2} Z`;
      svgElement(svg, 'path', { d, fill:COLORS[index % COLORS.length], stroke:'#fff', 'stroke-width':2 });
      const label = item.categories[index] == null ? String(index + 1) : item.categories[index];
      const y = 100 + index * 26;
      if (index < 8) {
        svgElement(svg, 'rect', { x:455, y:y - 11, width:11, height:11, rx:2, fill:COLORS[index % COLORS.length] });
        svgElement(svg, 'text', { x:473, y:y - 1, fill:'#42505e', 'font-size':11 }, shortLabel(label, 22) + ' · ' + Math.round(value / total * 100) + '%');
      }
      angle = next;
    });
    return svg;
  }

  function renderChart(chart) {
    const svg = chartFrame(chart || { title:'Gráfico' });
    const type = chart && chart.type;
    if (type === 'pie' || type === 'ofpie' || type === 'doughnut') return renderPie(svg, chart, type === 'doughnut');
    if (type === 'line' || type === 'area' || type === 'scatter' || type === 'radar' || type === 'bubble' || type === 'stock') return renderLine(svg, chart);
    if (type === 'bar' || type === 'horizontalBar') return renderCartesian(svg, chart, type === 'horizontalBar');
    return renderDataFallback(svg, chart || { series:[] });
  }

  window.XLSXViewerVisuals = { extract, renderChart };
})();
