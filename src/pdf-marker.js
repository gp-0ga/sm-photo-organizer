const LEGACY_MARKER_PREFIX = "SM-ASSET-NAME|1|";
const MARKER_PREFIX = "SM-ASSET-NAME|2|";
const MARKER_END = "|END";

function encodeBase64Url(value) {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function decodeBase64Url(value) {
  const normalized = value.replace(/-/g, "+").replace(/_/g, "/");
  const binary = atob(normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "="));
  return new TextDecoder().decode(Uint8Array.from(binary, (character) => character.charCodeAt(0)));
}

export function markerAssetPrintPayload(assetNumber, assetName) {
  return `${MARKER_PREFIX}${encodeBase64Url(JSON.stringify([String(assetNumber), String(assetName)]))}${MARKER_END}`;
}

export function parseMarkerAssetsFromTextItems(textItems) {
  const assets = [];
  const seen = new Set();
  const strings = textItems.map((item) => String(item?.str ?? item ?? ""));

  for (let index = 0; index < strings.length; index += 1) {
    let candidate = strings[index].replace(/\s*\|\s*/g, "|").trim();
    if (!candidate.includes("SM-ASSET")) continue;
    for (let next = index; next < Math.min(strings.length, index + 24); next += 1) {
      if (next > index) candidate += strings[next].replace(/\s*\|\s*/g, "|").trim();
      const version2Start = candidate.indexOf(MARKER_PREFIX);
      const end = candidate.indexOf(MARKER_END, version2Start + MARKER_PREFIX.length);
      let assetNumber;
      let assetName;
      if (version2Start >= 0 && end >= 0) {
        try {
          [assetNumber, assetName] = JSON.parse(decodeBase64Url(candidate.slice(version2Start + MARKER_PREFIX.length, end)));
        } catch {
          break;
        }
      } else {
        const legacyStart = candidate.indexOf(LEGACY_MARKER_PREFIX);
        if (legacyStart < 0) continue;
        const fields = candidate.slice(legacyStart + LEGACY_MARKER_PREFIX.length).split("|");
        assetNumber = fields[0]?.trim();
        assetName = fields.slice(1).join("|").trim();
      }
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
