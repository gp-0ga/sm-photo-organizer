import QRCode from "qrcode";
import { markerPayload } from "./domain.js";

export function createQr(payload) {
  return QRCode.toDataURL(payload, {
    errorCorrectionLevel: "H",
    margin: 3,
    width: 360,
    color: { dark: "#111827", light: "#ffffff" },
  });
}

export function createAssetQr(assetNumber) {
  return createQr(markerPayload(assetNumber));
}
