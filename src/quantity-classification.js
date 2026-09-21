export const STANDARD_CATEGORIES = [
  {
    id: "waterproofing_renovation",
    label: "防水改修",
    order: 2,
    divisions: {
      removal: { label: "撤去", sourcePage: 51, guidance: "防水層と防水保護層を分け、設計寸法による面積または体積で計測。部分改修のカッター入れは長さ。", items: ["防水保護コンクリート撤去", "防水立上り部保護撤去", "防水層撤去", "シーリング撤去", "手すり撤去", "笠木撤去", "ルーフドレン撤去", "とい撤去"] },
      renovation: { label: "改修", sourcePage: 51, guidance: "下地処理は工法・部位別、防水層と保護層は面積・長さ・箇所数で計測。ひび割れ補修は工法別の長さ。", items: ["既存下地補修", "アスファルト防水", "ルーフィングシート防水", "塗膜防水", "伸縮目地", "成形緩衝材", "防水入隅処理", "シーリング", "防水立上り部保護", "防水層押え金物", "防水保護コンクリート", "ルーフドレン", "手すり", "笠木", "とい"] },
    },
  },
  {
    id: "exterior_wall_renovation",
    label: "外壁改修",
    order: 3,
    divisions: {
      removal: { label: "撤去", sourcePage: 51, guidance: "既存仕上は設計寸法による面積、部分改修のカッター入れは設計寸法による長さで計測。", items: ["壁タイル撤去", "壁モルタル撤去", "役物モルタル撤去", "既存塗膜等の除去", "カッター入れ", "シーリング撤去"] },
      renovation: { label: "改修", sourcePage: 51, guidance: "調査・ひび割れ・欠損・浮き・新設仕上を工法と部位ごとに分け、面積・長さ・箇所数で計測。", items: ["施工数量調査", "外壁清掃", "ひび割れ部改修", "欠損部改修", "浮き部改修", "下地調整", "仕上塗材塗り", "壁タイル張り", "壁モルタル塗り", "役物モルタル塗り", "シーリング"] },
    },
  },
  {
    id: "opening_renovation",
    label: "建具改修",
    order: 4,
    divisions: {
      removal: { label: "撤去", sourcePage: 52, guidance: "建具の種別ごとに、内法寸法による箇所数・面積・長さで計測。かぶせ工法の既存枠補強等は原則対象外。", items: ["アルミニウム製建具撤去", "樹脂製建具撤去", "鋼製建具撤去", "鋼製軽量建具撤去", "ステンレス製建具撤去", "木製建具撤去", "シャッター撤去", "オーバーヘッドドア撤去", "カッター入れ", "シーリング撤去", "ガラス撤去"] },
      renovation: { label: "改修", sourcePage: 52, guidance: "建具の種別ごとに内法寸法で計測。新設・補修は箇所数・面積・長さ、建具周囲補修は長さ。", items: ["AW アルミニウム製窓", "AG アルミニウム製ガラリ", "AD アルミニウム製ドア", "PW 樹脂製窓", "PD 樹脂製ドア", "SD 鋼製ドア", "SG 鋼製ガラリ", "LD 鋼製軽量ドア", "SSD ステンレス製ドア", "SSW ステンレス製窓", "WD 木製ドア", "自動ドア開閉装置", "シャッター", "オーバーヘッドドア", "ガラス", "シーリング"] },
    },
  },
  {
    id: "interior_renovation",
    label: "内装改修",
    order: 5,
    divisions: {
      removal: { label: "撤去", sourcePage: 52, guidance: "仕上材と下地材を部位・種別ごとに分け、設計寸法による面積・長さ・箇所数で計測。カッター入れは長さ。", items: ["床仕上撤去", "幅木撤去", "壁仕上撤去", "壁紙撤去", "壁下地撤去", "天井仕上撤去", "天井下地撤去", "間仕切撤去", "天井点検口撤去", "ブラインドボックス撤去"] },
      renovation: { label: "改修", sourcePage: 52, guidance: "床・壁・天井と下地処理を工法別に分け、設計寸法による面積・長さ・箇所数で計測。", items: ["床下地補修", "床見切縁", "ビニル床タイル張り", "ビニル床シート張り", "タイルカーペット張り", "カーペット敷き", "合成樹脂塗床", "床フローリング張り", "畳敷き", "床タイル張り", "床モルタル塗り", "ビニル幅木", "壁下地補修", "壁タイル張り", "壁モルタル塗り", "軽量鉄骨壁下地", "壁ボード張り", "壁紙張り", "天井ボード張り", "軽量鉄骨天井下地", "天井廻り縁", "天井点検口"] },
    },
  },
  {
    id: "painting_renovation",
    label: "塗装改修",
    order: 6,
    divisions: {
      renovation: { label: "改修", sourcePage: 52, guidance: "塗装仕様ごとに分け、設計寸法による面積・長さ・箇所数で計測。取り合い部は仕様・部位別に区分。", items: ["FE フタル酸樹脂エナメル塗り", "SOP 合成樹脂調合ペイント塗り", "EP 合成樹脂エマルションペイント塗り", "EP-G つや有り合成樹脂エマルションペイント塗り", "DP 耐候性塗料塗り", "NAD アクリル樹脂系非水分散形塗料塗り", "CL クリヤラッカー塗り", "UC ウレタン樹脂ワニス塗り", "OS オイルステイン塗り"] },
    },
  },
];

export const STANDARD_UNITS = ["㎡", "m", "m3", "か所", "枚", "本", "式"];
export const WORK_TYPES = ["建築改修工事", "建築工事"];
export const SUBJECTS = ["庁舎", "屋外施設", "その他"];

const CATEGORY_PARTS = {
  waterproofing_renovation: ["外部防水", "内部防水", "屋上", "立上り", "笠木", "とい"],
  exterior_wall_renovation: ["外壁", "北面", "東面", "南面", "西面", "バルコニー", "軒天"],
  opening_renovation: ["外部建具", "内部建具", "窓", "出入口", "ガラリ", "シャッター"],
  interior_renovation: ["床", "幅木・壁", "天井", "間仕切", "その他"],
  painting_renovation: ["外部塗装", "内部塗装", "鉄部", "木部", "細幅物"],
};

export function partsForCategory(categoryId) {
  return CATEGORY_PARTS[categoryId] || [];
}

export function categoryById(id) {
  return STANDARD_CATEGORIES.find((category) => category.id === id) || null;
}

export function standardClassificationLabel(classification = {}) {
  return [classification.categoryLabel, classification.divisionLabel, classification.part, classification.item].filter(Boolean).join(" / ");
}

function roundPositive(value, digits) {
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0) return 0;
  const factor = 10 ** digits;
  return Math.round((number + Number.EPSILON) * factor) / factor;
}

// 公共建築数量積算基準（令和5年改定）第1編2(6)(7)による。
export function roundMeasurementQuantity(value) {
  return roundPositive(value, 2);
}

export function roundBreakdownQuantity(value) {
  return roundPositive(value, Number(value) >= 100 ? 0 : 1);
}

export function breakdownQuantityDigits(value) {
  return Number(value) >= 100 ? 0 : 1;
}

export function measurementSuggestion(item = "") {
  const text = String(item);
  if (/ルーフドレン|点検口|建具|\b(?:AW|AG|AD|PW|PD|SD|SG|LD|SSD|SSW|WD)\b/.test(text)) {
    return { mode: "個数", unit: "か所" };
  }
  if (/シーリング|目地|カッター|ひび割れ|入隅|押え金物|見切|幅木|廻り縁|笠木|とい|手すり|周囲/.test(text)) {
    return { mode: "線数量", unit: "m" };
  }
  if (/コンクリート/.test(text) && /撤去|保護/.test(text)) return { mode: "体積または面積", unit: "m3" };
  return { mode: "面積", unit: "㎡" };
}
