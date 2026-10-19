const MARKER_PREFIX = "SM-ASSET-NAME|1|";

export function parseMarkerAssetsFromTextItems(textItems) {
  const assets = [];
  const seen = new Set();
  const strings = textItems.map((item) => String(item ?? ""));

  for (let index = 0; index < strings.length; index += 1) {
    let candidate = strings[index].replace(/\s*\|\s*/g, "|").trim();
    if (!candidate.includes("SM-ASSET-NAME")) continue;
    for (let next = index; next < Math.min(strings.length, index + 8); next += 1) {
      if (next > index) candidate += strings[next].replace(/\s*\|\s*/g, "|").trim();
      const start = candidate.indexOf(MARKER_PREFIX);
      if (start < 0) continue;
      const fields = candidate.slice(start + MARKER_PREFIX.length).split("|");
      const assetNumber = fields[0]?.trim();
      const assetName = fields.slice(1).join("|").trim();
      if (!assetNumber || !assetName) continue;
      if (!seen.has(assetNumber)) {
        seen.add(assetNumber);
        assets.push({ assetNumber, assetName, items: [] });
      }
      break;
    }
  }
  return assets;
}
