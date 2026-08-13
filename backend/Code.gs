// ════════════════════════════════════════════════
// FİNCANLAR ERP — Cari Modülü (Apps Script backend)
// Mevcut Stok Paneli ile AYNI Google E-Tablosunu kullanır ama tamamen ayrı bir
// Apps Script projesi/deploy'udur — buradaki bir hata canlı Stok Panelini etkilemez.
// ════════════════════════════════════════════════

const SHEET_ID = "17-eyhwLd-3vIkH4HArhPnc3Ty7gkrZVYMERYJynXG4Q";

const SHEETS = {
  cariHesaplar:   "CariHesaplar",
  cariHareketler: "CariHareketler",
};

// ── YARDIMCI FONKSİYONLAR ──
function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e8edf5");
  }
  return sheet;
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function logError(err) {
  try { Logger.log("HATA: " + err.message + "\n" + err.stack); } catch(e) {}
}

// ── GİRİŞ NOKTALARI ──
function doGet(e) {
  if (e.parameter && e.parameter.payload) {
    try {
      const parsed = JSON.parse(decodeURIComponent(e.parameter.payload));
      e = Object.assign({}, e, { postData: { contents: JSON.stringify(parsed) } });
    } catch(err) {}
  }
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    const body = e.postData ? JSON.parse(e.postData.contents) : e.parameter;
    const action = body.action;
    let result;
    switch (action) {
      case "getCariListesi": result = getCariListesi(); break;
      case "getCariDetay":   result = getCariDetay(body.cariId); break;
      case "saveCari":       result = saveCari(body); break;
      case "silCari":        result = silCari(body); break;
      case "cariHareketEkle": result = cariHareketEkle(body); break;
      case "cariHareketSil":  result = cariHareketSil(body); break;
      default: result = { error: "Bilinmeyen işlem: " + action };
    }
    return jsonResponse(result);
  } catch (err) {
    logError(err);
    return jsonResponse({ error: err.message });
  }
}

// ── VERİ FONKSİYONLARI ──

// Tüm cari hesapları, her birinin güncel bakiyesiyle birlikte döndürür.
// Bakiye = toplam BORÇ - toplam ALACAK (pozitifse cari bize borçlu, negatifse biz ona borçluyuz).
function getCariListesi() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const hSheet = getOrCreateSheet(ss, SHEETS.cariHesaplar, ["ID","TIP","AD","TELEFON","ADRES","VERGI_NO","NOT","TARIH"]);
  const hkSheet = getOrCreateSheet(ss, SHEETS.cariHareketler, ["ID","CARI_ID","TARIH","TIP","TUTAR","ACIKLAMA","KAYIT_TARIHI"]);

  const hData = hSheet.getDataRange().getValues();
  const hkData = hkSheet.getDataRange().getValues();

  // Her cari için bakiyeyi tek geçişte hesapla
  const bakiyeMap = {};
  for (let i = 1; i < hkData.length; i++) {
    const row = hkData[i];
    const cariId = String(row[1] || "");
    if (!cariId) continue;
    const tip = String(row[3] || "");
    const tutar = parseFloat(row[4]) || 0;
    if (!bakiyeMap[cariId]) bakiyeMap[cariId] = 0;
    bakiyeMap[cariId] += (tip === "Borç") ? tutar : -tutar;
  }

  const sonuc = [];
  for (let i = 1; i < hData.length; i++) {
    const row = hData[i];
    const id = String(row[0] || "");
    if (!id) continue;
    sonuc.push({
      id: id,
      tip: String(row[1] || ""),
      ad: String(row[2] || ""),
      telefon: String(row[3] || ""),
      adres: String(row[4] || ""),
      vergiNo: String(row[5] || ""),
      not: String(row[6] || ""),
      tarih: String(row[7] || ""),
      bakiye: bakiyeMap[id] || 0,
    });
  }
  return { ok: true, cariler: sonuc };
}

// Tek bir cari hesabın bilgisini + tüm hareket geçmişini (tarihe göre sıralı, kümülatif bakiyeli) döndürür.
function getCariDetay(cariId) {
  if (!cariId) return { ok: false, hata: "cariId gerekli" };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const hSheet = getOrCreateSheet(ss, SHEETS.cariHesaplar, ["ID","TIP","AD","TELEFON","ADRES","VERGI_NO","NOT","TARIH"]);
  const hkSheet = getOrCreateSheet(ss, SHEETS.cariHareketler, ["ID","CARI_ID","TARIH","TIP","TUTAR","ACIKLAMA","KAYIT_TARIHI"]);

  const hData = hSheet.getDataRange().getValues();
  let cari = null;
  for (let i = 1; i < hData.length; i++) {
    if (String(hData[i][0]) === String(cariId)) {
      cari = {
        id: String(hData[i][0]), tip: String(hData[i][1] || ""), ad: String(hData[i][2] || ""),
        telefon: String(hData[i][3] || ""), adres: String(hData[i][4] || ""),
        vergiNo: String(hData[i][5] || ""), not: String(hData[i][6] || ""), tarih: String(hData[i][7] || ""),
      };
      break;
    }
  }
  if (!cari) return { ok: false, hata: "Cari bulunamadı" };

  const hkData = hkSheet.getDataRange().getValues();
  let hareketler = [];
  for (let i = 1; i < hkData.length; i++) {
    const row = hkData[i];
    if (String(row[1]) !== String(cariId)) continue;
    hareketler.push({
      id: String(row[0]), cariId: String(row[1]), tarih: String(row[2] || ""),
      tip: String(row[3] || ""), tutar: parseFloat(row[4]) || 0,
      aciklama: String(row[5] || ""), kayitTarihi: String(row[6] || ""),
    });
  }
  // Tarihe göre sırala (eskiden yeniye), kümülatif bakiyeyi hesapla
  hareketler.sort((a, b) => new Date(a.tarih) - new Date(b.tarih));
  let bakiye = 0;
  hareketler.forEach(h => {
    bakiye += (h.tip === "Borç") ? h.tutar : -h.tutar;
    h.bakiyeSonrasi = bakiye;
  });
  hareketler.reverse(); // en yeni en üstte gösterilsin

  return { ok: true, cari: cari, hareketler: hareketler, bakiye: bakiye };
}

// body: { id (varsa güncelleme), tip, ad, telefon, adres, vergiNo, not }
function saveCari(body) {
  const ad = String(body.ad || "").trim();
  if (!ad) return { ok: false, hata: "Cari adı gerekli" };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEETS.cariHesaplar, ["ID","TIP","AD","TELEFON","ADRES","VERGI_NO","NOT","TARIH"]);
  const data = sheet.getDataRange().getValues();

  let id = String(body.id || "").trim();
  let satirIdx = -1;
  if (id) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === id) { satirIdx = i + 1; break; }
    }
  }
  if (!id) id = "cr_" + Date.now();

  const satir = [
    id,
    String(body.tip || "Müşteri"),
    ad,
    String(body.telefon || ""),
    String(body.adres || ""),
    String(body.vergiNo || ""),
    String(body.not || ""),
    satirIdx > 0 ? data[satirIdx - 1][7] : Utilities.formatDate(new Date(), "Europe/Istanbul", "dd/MM/yyyy HH:mm"),
  ];
  if (satirIdx > 0) sheet.getRange(satirIdx, 1, 1, satir.length).setValues([satir]);
  else sheet.appendRow(satir);

  return { ok: true, id: id };
}

// body: { id } — sadece hiç hareketi olmayan cari silinebilir (güvenlik için)
function silCari(body) {
  const id = String(body.id || "").trim();
  if (!id) return { ok: false, hata: "id gerekli" };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const hkSheet = getOrCreateSheet(ss, SHEETS.cariHareketler, ["ID","CARI_ID","TARIH","TIP","TUTAR","ACIKLAMA","KAYIT_TARIHI"]);
  const hkData = hkSheet.getDataRange().getValues();
  for (let i = 1; i < hkData.length; i++) {
    if (String(hkData[i][1]) === id) {
      return { ok: false, hata: "Bu cariye ait hareketler var, önce onları silin veya cariyi silmeyin" };
    }
  }

  const sheet = getOrCreateSheet(ss, SHEETS.cariHesaplar, ["ID","TIP","AD","TELEFON","ADRES","VERGI_NO","NOT","TARIH"]);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === id) { sheet.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false, hata: "Cari bulunamadı" };
}

// body: { cariId, tarih, tip (Borç/Alacak), tutar, aciklama }
function cariHareketEkle(body) {
  const cariId = String(body.cariId || "").trim();
  const tip = String(body.tip || "").trim();
  const tutar = parseFloat(body.tutar) || 0;
  if (!cariId) return { ok: false, hata: "cariId gerekli" };
  if (tip !== "Borç" && tip !== "Alacak") return { ok: false, hata: "tip Borç veya Alacak olmalı" };
  if (tutar <= 0) return { ok: false, hata: "Tutar sıfırdan büyük olmalı" };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEETS.cariHareketler, ["ID","CARI_ID","TARIH","TIP","TUTAR","ACIKLAMA","KAYIT_TARIHI"]);
  const id = "hk_" + Date.now();
  const tarih = String(body.tarih || Utilities.formatDate(new Date(), "Europe/Istanbul", "yyyy-MM-dd"));
  sheet.appendRow([id, cariId, tarih, tip, tutar, String(body.aciklama || ""),
    Utilities.formatDate(new Date(), "Europe/Istanbul", "dd/MM/yyyy HH:mm")]);

  return { ok: true, id: id };
}

// body: { id }
function cariHareketSil(body) {
  const id = String(body.id || "").trim();
  if (!id) return { ok: false, hata: "id gerekli" };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEETS.cariHareketler, ["ID","CARI_ID","TARIH","TIP","TUTAR","ACIKLAMA","KAYIT_TARIHI"]);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === id) { sheet.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false, hata: "Hareket bulunamadı" };
}
