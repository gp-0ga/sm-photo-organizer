export function photoBookKitArchiveEntries({
  healthWorkbook,
  instruction,
  guide,
  photoBookBat,
  photoBookRunner,
  photoBookBuilder,
}) {
  if (!healthWorkbook) throw new Error("先に健全度判定表を読み込んでください。");
  return [
    { path: "START-PHOTO-BOOK.bat", blob: new Blob([photoBookBat], { type: "text/plain;charset=us-ascii" }) },
    { path: "photo-book-runner.ps1", blob: new Blob([photoBookRunner], { type: "text/plain;charset=utf-8" }) },
    { path: "build-photo-album.ps1", blob: new Blob([photoBookBuilder], { type: "text/plain;charset=utf-8" }) },
    { path: "photo-book-instruction.json", blob: new Blob([JSON.stringify(instruction, null, 2)], { type: "application/json" }) },
    { path: "health.xlsx", blob: healthWorkbook },
    // UTF-8 BOMを付け、Windowsのメモ帳でも日本語の案内文を正しく開けるようにする。
    { path: "README.txt", blob: new Blob(["\uFEFF", guide], { type: "text/plain;charset=utf-8" }) },
  ];
}
