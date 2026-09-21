const DB_NAME = "asset-photo-organizer";
const DB_VERSION = 1;
const STORE_NAME = "sessions";
const SESSION_ID = "current";

function openDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(STORE_NAME)) {
        request.result.createObjectStore(STORE_NAME, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("保存領域を開けませんでした。"));
  });
}

function requestAsPromise(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("作業内容を保存できませんでした。"));
  });
}

function serializeFile(file) {
  if (!file) return null;
  return {
    name: file.name,
    type: file.type,
    lastModified: file.lastModified,
    blob: file,
  };
}

function restoreFile(record) {
  if (!record?.blob) return null;
  return new File([record.blob], record.name || "保存ファイル", {
    type: record.type || record.blob.type || "application/octet-stream",
    lastModified: record.lastModified || Date.now(),
  });
}

function createRecord({ id, name, assets, photos, healthWorkbookFile, albumTemplateFile, photoTargetKb, bookmarkPhotoId }) {
  return {
    id,
    kind: id === SESSION_ID ? "current" : "snapshot",
    name: name || "",
    savedAt: new Date().toISOString(),
    assets,
    photos: photos.map((photo) => ({ ...photo, file: serializeFile(photo.file) })),
    healthWorkbookFile: serializeFile(healthWorkbookFile),
    albumTemplateFile: serializeFile(albumTemplateFile),
    photoTargetKb: String(photoTargetKb ?? "200"),
    bookmarkPhotoId: bookmarkPhotoId ?? null,
  };
}

export async function savePhotoSession({ assets, photos, healthWorkbookFile, albumTemplateFile, photoTargetKb, bookmarkPhotoId, id = SESSION_ID, name = "" }) {
  if (!globalThis.indexedDB) throw new Error("このブラウザは端末内保存に対応していません。");
  const database = await openDatabase();
  const record = createRecord({ id, name, assets, photos, healthWorkbookFile, albumTemplateFile, photoTargetKb, bookmarkPhotoId });
  await requestAsPromise(database.transaction(STORE_NAME, "readwrite").objectStore(STORE_NAME).put(record));
  database.close();
  return record.savedAt;
}

export async function loadPhotoSession(id = SESSION_ID) {
  if (!globalThis.indexedDB) return null;
  const database = await openDatabase();
  const record = await requestAsPromise(database.transaction(STORE_NAME, "readonly").objectStore(STORE_NAME).get(id));
  database.close();
  if (!record) return null;
  return {
    ...record,
    photos: (record.photos ?? []).map((photo) => ({
      ...photo,
      file: restoreFile(photo.file),
    })).filter((photo) => photo.file),
    healthWorkbookFile: restoreFile(record.healthWorkbookFile),
    albumTemplateFile: restoreFile(record.albumTemplateFile),
  };
}

export async function savePhotoSnapshot({ name, assets, photos, healthWorkbookFile, albumTemplateFile, photoTargetKb, bookmarkPhotoId }) {
  const id = `snapshot-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  const savedAt = await savePhotoSession({ id, name, assets, photos, healthWorkbookFile, albumTemplateFile, photoTargetKb, bookmarkPhotoId });
  return { id, savedAt };
}

export async function listPhotoSnapshots() {
  if (!globalThis.indexedDB) return [];
  const database = await openDatabase();
  const records = await requestAsPromise(database.transaction(STORE_NAME, "readonly").objectStore(STORE_NAME).getAll());
  database.close();
  return records.filter((record) => record.kind === "snapshot").sort((a, b) => String(b.savedAt).localeCompare(String(a.savedAt)));
}

export async function deletePhotoSnapshot(id) {
  if (!globalThis.indexedDB || !id) return;
  const database = await openDatabase();
  await requestAsPromise(database.transaction(STORE_NAME, "readwrite").objectStore(STORE_NAME).delete(id));
  database.close();
}

export async function clearPhotoSession() {
  if (!globalThis.indexedDB) return;
  const database = await openDatabase();
  await requestAsPromise(database.transaction(STORE_NAME, "readwrite").objectStore(STORE_NAME).delete(SESSION_ID));
  database.close();
}
