function inside(point, polygon) {
  let result = false;
  for (let i = 0, j = polygon.length - 1; i < polygon.length; j = i, i += 1) {
    const a = polygon[i];
    const b = polygon[j];
    if ((a.y > point.y) !== (b.y > point.y) && point.x < ((b.x - a.x) * (point.y - a.y)) / (b.y - a.y) + a.x) result = !result;
  }
  return result;
}

function darkness(image, x, y) {
  const ix = Math.round(x);
  const iy = Math.round(y);
  if (ix < 0 || iy < 0 || ix >= image.width || iy >= image.height) return false;
  const offset = (iy * image.width + ix) * 4;
  const { data } = image;
  return (data[offset] * 0.299 + data[offset + 1] * 0.587 + data[offset + 2] * 0.114) < 170;
}

function edgeCoverage(image, x1, y1, x2, y2, vertical = false) {
  const length = Math.hypot(x2 - x1, y2 - y1);
  const count = Math.max(2, Math.ceil(length));
  let dark = 0;
  for (let i = 0; i <= count; i += 1) {
    const x = x1 + ((x2 - x1) * i) / count;
    const y = y1 + ((y2 - y1) * i) / count;
    if ((vertical && [-2, -1, 0, 1, 2].some((delta) => darkness(image, x + delta, y))) ||
        (!vertical && [-1, 0, 1].some((delta) => darkness(image, x, y + delta)))) dark += 1;
  }
  return dark / (count + 1);
}

function interiorDarkness(image, box) {
  let dark = 0;
  let total = 0;
  for (let iy = 1; iy <= 5; iy += 1) {
    for (let ix = 1; ix <= 5; ix += 1) {
      const x = box.x + (box.width * ix) / 6;
      const y = box.y + (box.height * iy) / 6;
      if (darkness(image, x, y)) dark += 1;
      total += 1;
    }
  }
  return dark / total;
}

export function intersectionOverUnion(a, b) {
  const width = Math.max(0, Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x));
  const height = Math.max(0, Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y));
  const intersection = width * height;
  return intersection / (a.width * a.height + b.width * b.height - intersection || 1);
}

function horizontalSegments(image, bounds) {
  const result = [];
  const minLength = Math.max(12, Math.round(bounds.width * 0.04));
  for (let y = bounds.y + 2; y < bounds.y + bounds.height - 2; y += 1) {
    let start = -1;
    let lastDark = -1;
    for (let x = bounds.x + 2; x <= bounds.x + bounds.width - 2; x += 1) {
      const dark = darkness(image, x, y);
      if (dark) {
        if (start < 0) start = x;
        lastDark = x;
      }
      if (start >= 0 && (!dark && x - lastDark > 2 || x === bounds.x + bounds.width - 2)) {
        if (lastDark - start >= minLength && lastDark - start <= bounds.width * 0.75) {
          result.push({ x1: start, x2: lastDark, y });
        }
        start = -1;
      }
    }
  }
  result.sort((a, b) => a.y - b.y || a.x1 - b.x1);
  const unique = [];
  for (const line of result) {
    if (unique.some((prior) => Math.abs(prior.y - line.y) <= 3 && Math.abs(prior.x1 - line.x1) <= 3 && Math.abs(prior.x2 - line.x2) <= 3)) continue;
    unique.push(line);
  }
  return unique.slice(0, 450);
}

export function detectOpeningRectangles(image, walls, { maxCandidates = 40 } = {}) {
  if (!image?.data || !image.width || !image.height || !Array.isArray(walls)) return [];
  const candidates = [];
  for (const wall of walls) {
    if (!Array.isArray(wall.points) || wall.points.length < 3) continue;
    const xs = wall.points.map((point) => point.x);
    const ys = wall.points.map((point) => point.y);
    const left = Math.max(0, Math.floor(Math.min(...xs)));
    const top = Math.max(0, Math.floor(Math.min(...ys)));
    const right = Math.min(image.width - 1, Math.ceil(Math.max(...xs)));
    const bottom = Math.min(image.height - 1, Math.ceil(Math.max(...ys)));
    const bounds = { x: left, y: top, width: right - left, height: bottom - top };
    if (bounds.width < 30 || bounds.height < 30) continue;
    const lines = horizontalSegments(image, bounds);
    for (let i = 0; i < lines.length; i += 1) {
      const topLine = lines[i];
      for (let j = i + 1; j < lines.length; j += 1) {
        const bottomLine = lines[j];
        const height = bottomLine.y - topLine.y;
        if (height > bounds.height * 0.65) break;
        if (height < 12 || Math.abs(topLine.x1 - bottomLine.x1) > 4 || Math.abs(topLine.x2 - bottomLine.x2) > 4) continue;
        const width = (topLine.x2 - topLine.x1 + bottomLine.x2 - bottomLine.x1) / 2;
        if (width < 12 || width / height < 0.35 || width / height > 5 || width * height > bounds.width * bounds.height * 0.25) continue;
        const box = { x: Math.round((topLine.x1 + bottomLine.x1) / 2), y: topLine.y, width: Math.round(width), height };
        if (!inside({ x: box.x + box.width / 2, y: box.y + box.height / 2 }, wall.points)) continue;
        const leftEdge = edgeCoverage(image, box.x, box.y, box.x, box.y + box.height, true);
        const rightEdge = edgeCoverage(image, box.x + box.width, box.y, box.x + box.width, box.y + box.height, true);
        const topEdge = edgeCoverage(image, box.x, box.y, box.x + box.width, box.y);
        const bottomEdge = edgeCoverage(image, box.x, box.y + box.height, box.x + box.width, box.y + box.height);
        const border = Math.min(leftEdge, rightEdge, topEdge, bottomEdge);
        const interior = interiorDarkness(image, box);
        if (border < 0.72 || interior > 0.44) continue;
        candidates.push({ ...box, wallId: wall.id, confidence: Math.round(Math.min(0.95, border * (1 - interior) * 0.95) * 100) / 100 });
      }
    }
  }
  candidates.sort((a, b) => b.confidence - a.confidence || b.width * b.height - a.width * a.height);
  const unique = [];
  for (const candidate of candidates) {
    if (unique.some((existing) => intersectionOverUnion(existing, candidate) > 0.55)) continue;
    unique.push(candidate);
    if (unique.length >= maxCandidates) break;
  }
  return unique.sort((a, b) => a.y - b.y || a.x - b.x);
}
