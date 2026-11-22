// Apps Script untuk Catatan Keuangan dengan login Google.
// Pengaturan: Project Settings > Script properties:
//   GOOGLE_CLIENT_ID = Client ID dari Google Cloud (OAuth client, tipe Web application)
//   OWNER_EMAIL      = email Google Anda (satu-satunya yang boleh menulis)

const SHEET_DATA = "Data";
const SHEET_MASTER = "Master";
const BATAS_BARIS = 5000;

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function prop(nama) {
  return PropertiesService.getScriptProperties().getProperty(nama) || "";
}

// Memeriksa ID token dari Google. Hanya lolos jika token ditujukan untuk aplikasi ini
// dan emailnya adalah email pemilik yang terverifikasi.
function cekLogin(idToken) {
  if (!idToken) return { ok: false, error: "login" };
  const r = UrlFetchApp.fetch(
    "https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(idToken),
    { muteHttpExceptions: true }
  );
  if (r.getResponseCode() !== 200) return { ok: false, error: "login" };
  const t = JSON.parse(r.getContentText());
  if (t.aud !== prop("GOOGLE_CLIENT_ID")) return { ok: false, error: "aud" };
  if (["accounts.google.com", "https://accounts.google.com"].indexOf(t.iss) < 0) return { ok: false, error: "iss" };
  if (String(t.email_verified) !== "true") return { ok: false, error: "email" };
  if (String(t.email || "").toLowerCase() !== prop("OWNER_EMAIL").toLowerCase()) return { ok: false, error: "email" };
  if (Number(t.exp) * 1000 < Date.now()) return { ok: false, error: "login" };
  return { ok: true };
}

function kolom(sh, col, n) {
  return sh.getRange(2, col, n, 1).getValues().map(r => String(r[0]).trim()).filter(x => x !== "");
}
// Membaca daftar dari sheet Master: kategori (E dan F), rekening (C), Keterangan (kolom Ket_...)
function daftar() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_MASTER);
  const kat = { Pengeluaran: kolom(sh, 5, 5), Pemasukan: kolom(sh, 6, 6) };
  const rek = kolom(sh, 3, 5);
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const ket = {};
  header.forEach((h, i) => {
    const m = String(h).match(/^Ket_(Pengeluaran|Pemasukan)_(.+)$/);
    if (!m) return;
    const nama = (kat[m[1]] || []).find(k => k.replace(/ /g, "_") === m[2]);
    if (!nama) return;
    ket[m[1] + "|" + nama] = sh.getRange(2, i + 1, 40, 1).getValues()
      .map(r => String(r[0]).trim()).filter(x => x !== "");
  });
  return { kategori: kat, rekening: rek, ket: ket };
}
function baris(sh) {
  const vals = sh.getRange(2, 1, BATAS_BARIS - 1, 1).getValues();
  let last = 1;
  for (let i = 0; i < vals.length; i++) if (vals[i][0] !== "") last = i + 2;
  return last;
}

function doGet(e) {
  const cek = cekLogin(e.parameter.idToken);
  if (!cek.ok) return json(cek);
  if (e.parameter.action === "lists") return json(Object.assign({ ok: true }, daftar()));
  return json({ ok: true, info: "Catatan keuangan API" });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    const cek = cekLogin(body.idToken);
    if (!cek.ok) return json(cek);

    const m = daftar();
    const tipe = String(body.tipe || "");
    const kat = String(body.kategori || "").trim();
    const jml = Number(body.jumlah);
    if (["Pemasukan", "Pengeluaran"].indexOf(tipe) < 0) return json({ ok: false, error: "tipe" });
    if ((m.kategori[tipe] || []).indexOf(kat) < 0) return json({ ok: false, error: "kategori" });
    if (!(jml > 0)) return json({ ok: false, error: "jumlah" });
    const p = String(body.tanggal || "").split("-");
    if (p.length !== 3) return json({ ok: false, error: "tanggal" });
    const tgl = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));

    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_DATA);
    const last = baris(sh);
    const tag = "PWA:" + String(body.uid || "");
    // Cegah dobel jika aplikasi mengirim ulang transaksi yang sama
    if (body.uid && last > 1) {
      const catatan = sh.getRange(2, 7, last - 1, 1).getValues();
      for (let i = 0; i < catatan.length; i++) {
        if (String(catatan[i][0]).indexOf(tag) === 0) return json({ ok: true, duplikat: true });
      }
    }
    const catatanTeks = tag + (body.catatan ? " | " + String(body.catatan).trim() : "");
    const row = last + 1;
    sh.getRange(row, 1, 1, 7).setValues([[
      tgl, tipe, kat,
      String(body.keterangan || "").replace(/\s+/g, " ").trim(),
      jml, String(body.rekening || ""), catatanTeks
    ]]);
    return json({ ok: true, baris: row });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
