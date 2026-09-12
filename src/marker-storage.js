export const storedAssetsKey = "asset-marker.assets.v1";
export const storedSelectionKey = "asset-marker.selection.v1";

export function saveMarkerAssets(storage, assets, now = new Date()) {
  const markerAssets = assets.map(({ assetNumber, assetName }) => ({ assetNumber, assetName }));
  storage.setItem(storedAssetsKey, JSON.stringify({ version: 1, assets: markerAssets, savedAt: now.toISOString() }));
}

export function loadMarkerAssets(storage) {
  try {
    const stored = JSON.parse(storage.getItem(storedAssetsKey) || "null");
    if (!stored || stored.version !== 1 || !Array.isArray(stored.assets) || !stored.assets.length) return null;
    const assets = stored.assets.filter((asset) => typeof asset?.assetNumber === "string" && typeof asset?.assetName === "string");
    return assets.length ? assets : null;
  } catch {
    return null;
  }
}

export function saveMarkerSelection(storage, assetNumber) {
  storage.setItem(storedSelectionKey, assetNumber);
}

export function loadMarkerSelection(storage) {
  try {
    return storage.getItem(storedSelectionKey);
  } catch {
    return null;
  }
}

export function clearMarkerSelection(storage) {
  storage.removeItem(storedSelectionKey);
}
