import { attr, children, first, numberAttr, warn } from './ooxml-common.js';

const clamp = value => Math.max(0, Math.min(1, value));
const preset = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', cyan: '00FFFF', magenta: 'FF00FF', gray: '808080', dkBlue: '00008B', ltBlue: 'ADD8E6', transparent: '000000' };

function rgbToHsl(rgb) {
  const [r, g, b] = rgb, low = Math.min(...rgb), high = Math.max(...rgb), delta = high - low;
  const lightness = (high + low) / 2;
  if (!delta) return [0, 0, lightness];
  let hue = high === r ? (g - b) / delta + (g < b ? 6 : 0) : high === g ? (b - r) / delta + 2 : (r - g) / delta + 4;
  return [hue / 6, delta / (1 - Math.abs(2 * lightness - 1)), lightness];
}
function hslToRgb([hue, saturation, lightness]) {
  hue = (hue % 1 + 1) % 1; saturation = clamp(saturation); lightness = clamp(lightness);
  const chroma = (1 - Math.abs(2 * lightness - 1)) * saturation, sector = hue * 6;
  const x = chroma * (1 - Math.abs(sector % 2 - 1)), m = lightness - chroma / 2;
  const rgb = sector < 1 ? [chroma, x, 0] : sector < 2 ? [x, chroma, 0] : sector < 3 ? [0, chroma, x] : sector < 4 ? [0, x, chroma] : sector < 5 ? [x, 0, chroma] : [chroma, 0, x];
  return rgb.map(component => component + m);
}

// Color transformations apply in XML order. Luminance/saturation are HSL,
// whereas tint/shade mix the RGB color with white/black respectively.
export function readDrawingColor(container, theme = {}, warnings = []) {
  const source = children(container).find(node => ['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'hslClr'].includes(node.localName));
  if (!source) return null;
  const value = attr(source, 'val');
  let hex;
  if (source.localName === 'schemeClr') {
    const name = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' }[value] || value;
    hex = theme[name]?.replace(/^#/, '');
  } else if (source.localName === 'sysClr') hex = attr(source, 'lastClr');
  else if (source.localName === 'prstClr') hex = preset[value];
  else hex = value;
  let rgb;
  if (source.localName === 'hslClr') rgb = hslToRgb([numberAttr(source, 'hue') / 21600000, numberAttr(source, 'sat') / 100000, numberAttr(source, 'lum') / 100000]);
  else if (/^[0-9a-f]{6}$/i.test(hex || '')) rgb = [0, 2, 4].map(offset => parseInt(hex.slice(offset, offset + 2), 16) / 255);
  else { warn(warnings, '部分 PPTX 主题颜色或预设颜色尚未解析。'); return null; }
  let opacity = value === 'transparent' ? 0 : 1;
  for (const modifier of children(source)) {
    const n = numberAttr(modifier, 'val') / 100000, type = modifier.localName;
    if (type === 'alpha') opacity = clamp(n);
    else if (type === 'alphaMod') opacity = clamp(opacity * n);
    else if (type === 'alphaOff') opacity = clamp(opacity + n);
    else if (type === 'tint') rgb = rgb.map(component => clamp(component * n + 1 - n));
    else if (type === 'shade') rgb = rgb.map(component => clamp(component * n));
    else if (['lum', 'lumMod', 'lumOff', 'sat', 'satMod', 'satOff'].includes(type)) {
      const hsl = rgbToHsl(rgb), index = type.startsWith('lum') ? 2 : 1;
      hsl[index] = clamp(type.endsWith('Mod') ? hsl[index] * n : type.endsWith('Off') ? hsl[index] + n : n);
      rgb = hslToRgb(hsl);
    } else if (type === 'hue' || type === 'hueOff' || type === 'hueMod') {
      const hsl = rgbToHsl(rgb);
      hsl[0] = type === 'hue' ? numberAttr(modifier, 'val') / 21600000 : type === 'hueOff' ? hsl[0] + numberAttr(modifier, 'val') / 21600000 : hsl[0] * n;
      rgb = hslToRgb(hsl);
    } else warn(warnings, `PPTX 颜色变换 ${type} 尚未还原。`);
  }
  return { color: '#' + rgb.map(component => Math.round(clamp(component) * 255).toString(16).padStart(2, '0')).join('').toUpperCase(), ...(opacity !== 1 ? { opacity } : {}) };
}

export function cssColor(color) {
  if (!color) return undefined;
  if (color.opacity === undefined || color.opacity === 1) return color.color;
  const rgb = [1, 3, 5].map(offset => parseInt(color.color.slice(offset, offset + 2), 16));
  return `rgba(${rgb.join(',')},${color.opacity})`;
}

export function readLinearFill(properties, theme, warnings) {
  const gradient = first(properties, 'gradFill');
  if (!gradient) return null;
  const linear = first(gradient, 'lin');
  if (!linear || first(gradient, 'path')) { warn(warnings, 'PPTX 路径渐变尚未还原。'); return null; }
  const stopNodes = children(first(gradient, 'gsLst'), 'gs');
  if (stopNodes.length < 2 || stopNodes.length > 1000) throw new Error('PPTX 线性渐变停靠点数量无效。');
  const stops = stopNodes.map(node => {
    const offset = numberAttr(node, 'pos', -1) / 100000;
    if (offset < 0 || offset > 1) throw new Error('PPTX 渐变停靠点位置无效。');
    const color = readDrawingColor(node, theme, warnings);
    return color ? { offset, ...color } : null;
  });
  if (stops.some(stop => !stop)) return null;
  const tile = first(gradient, 'tileRect');
  if (tile && ['l', 't', 'r', 'b'].some(name => numberAttr(tile, name))) warn(warnings, 'PPTX 渐变平铺区域尚未完整还原。');
  return { type: 'linear', angle: ((numberAttr(linear, 'ang') / 60000) % 360 + 360) % 360, scaled: ['1', 'true'].includes(attr(linear, 'scaled')), stops: stops.sort((a, b) => a.offset - b.offset) };
}
