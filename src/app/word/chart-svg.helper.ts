export type WordChartType = 'column' | 'bar' | 'line' | 'pie';

export interface WordChartSeries {
  name: string;
  values: number[];
}

export interface WordChartData {
  categories: string[];
  series: WordChartSeries[];
}

export const DEFAULT_CHART_DATA: WordChartData = {
  categories: ['Category 1', 'Category 2', 'Category 3', 'Category 4'],
  series: [
    { name: 'Series 1', values: [4.3, 2.5, 3.5, 4.5] },
    { name: 'Series 2', values: [2.4, 4.4, 1.8, 2.8] },
    { name: 'Series 3', values: [2.0, 2.0, 3.0, 5.0] },
  ],
};

const CHART_COLORS = ['#4472c4', '#ed7d31', '#70ad47', '#ffc000', '#5b9bd5', '#a5a5a5'];

function escapeXml(unsafe: string): string {
  return String(unsafe || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function getPieSlicePath(cx: number, cy: number, r: number, startAngleDeg: number, endAngleDeg: number): string {
  const rad = Math.PI / 180;
  const x1 = cx + r * Math.cos(startAngleDeg * rad);
  const y1 = cy + r * Math.sin(startAngleDeg * rad);
  const x2 = cx + r * Math.cos(endAngleDeg * rad);
  const y2 = cy + r * Math.sin(endAngleDeg * rad);
  const largeArcFlag = endAngleDeg - startAngleDeg > 180 ? 1 : 0;
  return `M ${cx} ${cy} L ${x1.toFixed(2)} ${y1.toFixed(2)} A ${r} ${r} 0 ${largeArcFlag} 1 ${x2.toFixed(2)} ${y2.toFixed(2)} Z`;
}

export function renderChartSvg(
  type: WordChartType,
  title: string,
  data: WordChartData = DEFAULT_CHART_DATA,
  width = 480,
  height = 300
): string {
  const chartType = (type || 'column').toLowerCase() as WordChartType;
  const chartTitle = title !== undefined && title !== null ? title : 'Chart Title';
  const categories = data?.categories?.length
    ? data.categories
    : ['Category 1', 'Category 2', 'Category 3', 'Category 4'];
  const series = data?.series?.length
    ? data.series
    : DEFAULT_CHART_DATA.series;

  let contentSvg = '';

  if (chartType === 'column') {
    // Column Chart: Clustered Column
    const plotLeft = 45;
    const plotRight = width - 25;
    const plotTop = 45;
    const plotBottom = height - 60;
    const plotWidth = plotRight - plotLeft;
    const plotHeight = plotBottom - plotTop;

    const maxVal = 6;
    const ticks = [0, 1, 2, 3, 4, 5, 6];

    let gridLinesSvg = '';
    ticks.forEach((tick) => {
      const y = plotBottom - (tick / maxVal) * plotHeight;
      gridLinesSvg += `<line x1="${plotLeft}" y1="${y.toFixed(1)}" x2="${plotRight}" y2="${y.toFixed(1)}" class="grid-line" stroke="#e1dfdd" stroke-width="1" />`;
      gridLinesSvg += `<text x="${plotLeft - 7}" y="${(y + 3.5).toFixed(1)}" text-anchor="end" class="chart-axis-text" font-family="'Segoe UI', Aptos, sans-serif" font-size="10" fill="#616161">${tick}</text>`;
    });

    let barsSvg = '';
    const numCats = categories.length || 1;
    const catSlotWidth = plotWidth / numCats;
    const numSeries = series.length || 1;
    const barWidth = Math.min(22, (catSlotWidth * 0.7) / numSeries);
    const totalGroupWidth = numSeries * barWidth;

    categories.forEach((cat, cIdx) => {
      const cx = plotLeft + (cIdx + 0.5) * catSlotWidth;
      const groupStartX = cx - totalGroupWidth / 2;

      series.forEach((s, sIdx) => {
        const val = Number(s.values[cIdx]) || 0;
        const barH = Math.max(0, Math.min(plotHeight, (val / maxVal) * plotHeight));
        const bx = groupStartX + sIdx * barWidth;
        const by = plotBottom - barH;
        const color = CHART_COLORS[sIdx % CHART_COLORS.length];
        barsSvg += `<rect x="${bx.toFixed(1)}" y="${by.toFixed(1)}" width="${(barWidth - 2).toFixed(1)}" height="${barH.toFixed(1)}" fill="${color}" rx="0.5" />`;
      });

      barsSvg += `<text x="${cx.toFixed(1)}" y="${(plotBottom + 16).toFixed(1)}" text-anchor="middle" class="chart-axis-text" font-family="'Segoe UI', Aptos, sans-serif" font-size="10" fill="#616161">${escapeXml(cat)}</text>`;
    });

    let legendSvg = '';
    const legendY = height - 16;
    const itemWidth = 85;
    const totalLegendWidth = numSeries * itemWidth;
    const legendStartX = (width - totalLegendWidth) / 2;

    series.forEach((s, sIdx) => {
      const lx = legendStartX + sIdx * itemWidth;
      const color = CHART_COLORS[sIdx % CHART_COLORS.length];
      legendSvg += `<rect x="${lx.toFixed(1)}" y="${(legendY - 8).toFixed(1)}" width="9" height="9" fill="${color}" rx="1" />`;
      legendSvg += `<text x="${(lx + 13).toFixed(1)}" y="${legendY.toFixed(1)}" class="chart-legend-text" font-family="'Segoe UI', Aptos, sans-serif" font-size="10.5" fill="#323130">${escapeXml(s.name)}</text>`;
    });

    contentSvg = `
      <g class="chart-grid">${gridLinesSvg}</g>
      <g class="chart-bars">${barsSvg}</g>
      <g class="chart-legend">${legendSvg}</g>
    `;
  } else if (chartType === 'line') {
    // Line Chart
    const plotLeft = 45;
    const plotRight = width - 25;
    const plotTop = 45;
    const plotBottom = height - 60;
    const plotWidth = plotRight - plotLeft;
    const plotHeight = plotBottom - plotTop;

    const maxVal = 6;
    const ticks = [0, 1, 2, 3, 4, 5, 6];

    let gridLinesSvg = '';
    ticks.forEach((tick) => {
      const y = plotBottom - (tick / maxVal) * plotHeight;
      gridLinesSvg += `<line x1="${plotLeft}" y1="${y.toFixed(1)}" x2="${plotRight}" y2="${y.toFixed(1)}" class="grid-line" stroke="#e1dfdd" stroke-width="1" />`;
      gridLinesSvg += `<text x="${plotLeft - 7}" y="${(y + 3.5).toFixed(1)}" text-anchor="end" class="chart-axis-text" font-family="'Segoe UI', Aptos, sans-serif" font-size="10" fill="#616161">${tick}</text>`;
    });

    const numCats = categories.length || 1;
    const catSlotWidth = plotWidth / numCats;

    let catLabelsSvg = '';
    categories.forEach((cat, cIdx) => {
      const cx = plotLeft + (cIdx + 0.5) * catSlotWidth;
      catLabelsSvg += `<text x="${cx.toFixed(1)}" y="${(plotBottom + 16).toFixed(1)}" text-anchor="middle" class="chart-axis-text" font-family="'Segoe UI', Aptos, sans-serif" font-size="10" fill="#616161">${escapeXml(cat)}</text>`;
    });

    let linesSvg = '';
    series.forEach((s, sIdx) => {
      const color = CHART_COLORS[sIdx % CHART_COLORS.length];
      const pts: string[] = [];
      categories.forEach((_, cIdx) => {
        const cx = plotLeft + (cIdx + 0.5) * catSlotWidth;
        const val = Number(s.values[cIdx]) || 0;
        const cy = plotBottom - Math.max(0, Math.min(plotHeight, (val / maxVal) * plotHeight));
        pts.push(`${cx.toFixed(1)},${cy.toFixed(1)}`);
      });

      linesSvg += `<polyline points="${pts.join(' ')}" fill="none" stroke="${color}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" />`;
    });

    let legendSvg = '';
    const legendY = height - 16;
    const numSeries = series.length || 1;
    const itemWidth = 85;
    const totalLegendWidth = numSeries * itemWidth;
    const legendStartX = (width - totalLegendWidth) / 2;

    series.forEach((s, sIdx) => {
      const lx = legendStartX + sIdx * itemWidth;
      const color = CHART_COLORS[sIdx % CHART_COLORS.length];
      legendSvg += `<line x1="${lx.toFixed(1)}" y1="${(legendY - 3).toFixed(1)}" x2="${(lx + 10).toFixed(1)}" y2="${(legendY - 3).toFixed(1)}" stroke="${color}" stroke-width="2.5" />`;
      legendSvg += `<text x="${(lx + 14).toFixed(1)}" y="${legendY.toFixed(1)}" class="chart-legend-text" font-family="'Segoe UI', Aptos, sans-serif" font-size="10.5" fill="#323130">${escapeXml(s.name)}</text>`;
    });

    contentSvg = `
      <g class="chart-grid">${gridLinesSvg}</g>
      <g class="chart-categories">${catLabelsSvg}</g>
      <g class="chart-lines">${linesSvg}</g>
      <g class="chart-legend">${legendSvg}</g>
    `;
  } else if (chartType === 'bar') {
    // Horizontal Bar Chart
    const plotLeft = 85;
    const plotRight = width - 25;
    const plotTop = 45;
    const plotBottom = height - 60;
    const plotWidth = plotRight - plotLeft;
    const plotHeight = plotBottom - plotTop;

    const maxVal = 6;
    const ticks = [0, 1, 2, 3, 4, 5, 6];

    let gridLinesSvg = '';
    ticks.forEach((tick) => {
      const x = plotLeft + (tick / maxVal) * plotWidth;
      gridLinesSvg += `<line x1="${x.toFixed(1)}" y1="${plotTop}" x2="${x.toFixed(1)}" y2="${plotBottom}" class="grid-line" stroke="#e1dfdd" stroke-width="1" />`;
      gridLinesSvg += `<text x="${x.toFixed(1)}" y="${(plotBottom + 14).toFixed(1)}" text-anchor="middle" class="chart-axis-text" font-family="'Segoe UI', Aptos, sans-serif" font-size="10" fill="#616161">${tick}</text>`;
    });

    const numCats = categories.length || 1;
    const catSlotHeight = plotHeight / numCats;
    const numSeries = series.length || 1;
    const barHeight = Math.min(16, (catSlotHeight * 0.7) / numSeries);
    const totalGroupHeight = numSeries * barHeight;

    let barsSvg = '';
    categories.forEach((cat, cIdx) => {
      const cy = plotTop + (cIdx + 0.5) * catSlotHeight;
      const groupStartY = cy - totalGroupHeight / 2;

      barsSvg += `<text x="${plotLeft - 7}" y="${(cy + 3.5).toFixed(1)}" text-anchor="end" class="chart-axis-text" font-family="'Segoe UI', Aptos, sans-serif" font-size="10" fill="#616161">${escapeXml(cat)}</text>`;

      series.forEach((s, sIdx) => {
        const val = Number(s.values[cIdx]) || 0;
        const barW = Math.max(0, Math.min(plotWidth, (val / maxVal) * plotWidth));
        const by = groupStartY + sIdx * barHeight;
        const color = CHART_COLORS[sIdx % CHART_COLORS.length];
        barsSvg += `<rect x="${plotLeft}" y="${by.toFixed(1)}" width="${barW.toFixed(1)}" height="${(barHeight - 2).toFixed(1)}" fill="${color}" rx="0.5" />`;
      });
    });

    let legendSvg = '';
    const legendY = height - 16;
    const itemWidth = 85;
    const totalLegendWidth = numSeries * itemWidth;
    const legendStartX = (width - totalLegendWidth) / 2;

    series.forEach((s, sIdx) => {
      const lx = legendStartX + sIdx * itemWidth;
      const color = CHART_COLORS[sIdx % CHART_COLORS.length];
      legendSvg += `<rect x="${lx.toFixed(1)}" y="${(legendY - 8).toFixed(1)}" width="9" height="9" fill="${color}" rx="1" />`;
      legendSvg += `<text x="${(lx + 13).toFixed(1)}" y="${legendY.toFixed(1)}" class="chart-legend-text" font-family="'Segoe UI', Aptos, sans-serif" font-size="10.5" fill="#323130">${escapeXml(s.name)}</text>`;
    });

    contentSvg = `
      <g class="chart-grid">${gridLinesSvg}</g>
      <g class="chart-bars">${barsSvg}</g>
      <g class="chart-legend">${legendSvg}</g>
    `;
  } else if (chartType === 'pie') {
    // Pie Chart: uses first series
    const cx = width / 2;
    const cy = 135;
    const r = 85;

    const firstSeries = series[0]?.values || [4.3, 2.5, 3.5, 4.5];
    const vals = categories.map((_, i) => Math.max(0, Number(firstSeries[i]) || 0));
    const total = vals.reduce((a, b) => a + b, 0) || 1;

    let slicesSvg = '';
    let currentAngle = -90;

    categories.forEach((cat, idx) => {
      const val = vals[idx];
      const sliceAngle = (val / total) * 360;
      const color = CHART_COLORS[idx % CHART_COLORS.length];

      if (sliceAngle >= 359.99) {
        slicesSvg += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${color}" />`;
      } else if (sliceAngle > 0) {
        const path = getPieSlicePath(cx, cy, r, currentAngle, currentAngle + sliceAngle);
        slicesSvg += `<path d="${path}" fill="${color}" stroke="#ffffff" stroke-width="1.5" />`;
      }
      currentAngle += sliceAngle;
    });

    let legendSvg = '';
    const legendY = height - 20;
    const numCats = categories.length || 1;
    const itemWidth = Math.min(105, (width - 40) / numCats);
    const totalLegendWidth = numCats * itemWidth;
    const legendStartX = (width - totalLegendWidth) / 2;

    categories.forEach((cat, idx) => {
      const lx = legendStartX + idx * itemWidth;
      const color = CHART_COLORS[idx % CHART_COLORS.length];
      legendSvg += `<rect x="${lx.toFixed(1)}" y="${(legendY - 8).toFixed(1)}" width="9" height="9" fill="${color}" rx="1" />`;
      legendSvg += `<text x="${(lx + 13).toFixed(1)}" y="${legendY.toFixed(1)}" class="chart-legend-text" font-family="'Segoe UI', Aptos, sans-serif" font-size="10.5" fill="#323130">${escapeXml(cat)}</text>`;
    });

    contentSvg = `
      <g class="chart-pie-slices">${slicesSvg}</g>
      <g class="chart-legend">${legendSvg}</g>
    `;
  }

  return `
    <svg viewBox="0 0 ${width} ${height}" class="word-chart-svg" width="100%" height="100%" xmlns="http://www.w3.org/2000/svg">
      <text x="${(width / 2).toFixed(1)}" y="24" text-anchor="middle" class="chart-title-text" font-family="'Segoe UI', Aptos, Calibri, sans-serif" font-size="14" font-weight="600" fill="#242424">${escapeXml(chartTitle)}</text>
      ${contentSvg}
    </svg>
  `.trim();
}
