import * as XLSX from "xlsx";
import React, { useState, useEffect, useMemo } from "react";
import {
  signInWithEmailAndPassword,
  signOut,
  onAuthStateChanged,
  createUserWithEmailAndPassword,
  updatePassword,
  reauthenticateWithCredential,
  EmailAuthProvider,
} from "firebase/auth";
import {
  collection,
  doc,
  setDoc,
  addDoc,
  updateDoc,
  deleteDoc,
  onSnapshot,
  serverTimestamp,
} from "firebase/firestore";
import { auth, secondaryAuth, db } from "./firebase";
import {
  LayoutDashboard,
  Users,
  CheckCircle2,
  XCircle,
  FileText,
  Wallet,
  LogOut,
  Copy,
  Check,
  MessageCircle,
  Plus,
  Edit2,
  Eye,
  Download,
  RefreshCw,
  Home,
  CreditCard,
  QrCode,
  AlertCircle,
  Trash2,
  Bell,
  BookOpen,
  Settings,
  HelpCircle,
  AlertTriangle,
  ArrowUpCircle,
  Clock,
  Activity,
  Lock,
  Menu,
  X,
  Search,
} from "lucide-react";

// ==========================================
// KONFIGURASI REKENING, IURAN & BANTUAN
// ==========================================
const CONFIG = {
  iuranPerBulan: 25000, // Total per bulan
  porsiPaguyuban: 20000, // Masuk Kas Paguyuban BSR
  porsiBlok9: 5000, // Masuk Kas Khusus BLOK 9
  waBantuan: "6282230280004",
  bca: { bank: "Bank BCA", norek: "3301144677", an: "Abdur Rouf" },
  bri: { bank: "Bank BRI", norek: "628801016232530", an: "Abdur Rouf" },
  qrisUrl: "/qris.jpeg",
};

const BULAN_LIST = [
  "Januari", "Februari", "Maret", "April", "Mei", "Juni",
  "Juli", "Agustus", "September", "Oktober", "November", "Desember",
];
const BULAN_SINGKAT = [
  "Jan", "Feb", "Mar", "Apr", "Mei", "Jun",
  "Jul", "Ags", "Sep", "Okt", "Nov", "Des",
];

// ==========================================
// HELPER FUNCTIONS
// ==========================================
const formatRupiah = (angka) =>
  new Intl.NumberFormat("id-ID", {
    style: "currency",
    currency: "IDR",
    minimumFractionDigits: 0,
  }).format(Number(angka) || 0);

const toFirebaseEmail = (username) =>
  `${username.trim().toLowerCase()}@blok9bsr.com`;
const toFirebasePassword = (password) => `${password.trim()}_bsr9`;

const formatWaNumber = (hp) => {
  if (!hp) return "";
  let clean = hp.replace(/\D/g, "");
  if (clean.startsWith("0")) clean = "62" + clean.slice(1);
  if (!clean.startsWith("62")) clean = "62" + clean;
  return clean;
};

const formatBulanLabel = (ym) => {
  const [y, m] = ym.split("-");
  return `${BULAN_SINGKAT[parseInt(m, 10) - 1] || m} ${y}`;
};

const formatTanggalIndo = (iso) => {
  if (!iso) return "-";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return iso.slice(0, 10);
  return d.toLocaleDateString("id-ID", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
};

// Hitung daftar bulan tunggakan dari Januari tahun ini s/d bulan berjalan
const getBulanWajibTahunIni = () => {
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth() + 1;
  const list = [];
  for (let i = 1; i <= m; i++) {
    list.push(`${y}-${String(i).padStart(2, "0")}`);
  }
  return list;
};

const compressImageToBase64 = (file, maxWidth = 800, quality = 0.6) =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = (e) => {
      const img = new Image();
      img.src = e.target.result;
      img.onload = () => {
        const canvas = document.createElement("canvas");
        let { width, height } = img;
        if (width > maxWidth) {
          height = Math.round((height * maxWidth) / width);
          width = maxWidth;
        }
        canvas.width = width;
        canvas.height = height;
        canvas.getContext("2d").drawImage(img, 0, 0, width, height);
        let dataUrl = canvas.toDataURL("image/jpeg", quality);
        if (dataUrl.length > 200000) dataUrl = canvas.toDataURL("image/jpeg", 0.35);
        resolve(dataUrl);
      };
      img.onerror = reject;
    };
    reader.onerror = reject;
  });
// Generator Kwitansi PNG otomatis sesuai desain BLOK 9 BSR
const downloadKwitansiPNG = (trx, wargaData = {}) => {
  const canvas = document.createElement("canvas");
  canvas.width = 800;
  canvas.height = 1120;
  const ctx = canvas.getContext("2d");

  const roundRect = (x, y, w, h, r) => {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.lineTo(x + w - r, y);
    ctx.quadraticCurveTo(x + w, y, x + w, y + r);
    ctx.lineTo(x + w, y + h - r);
    ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
    ctx.lineTo(x + r, y + h);
    ctx.quadraticCurveTo(x, y + h, x, y + h - r);
    ctx.lineTo(x, y + r);
    ctx.quadraticCurveTo(x, y, x + r, y);
    ctx.closePath();
  };

  const renderCanvas = (logoImg = null) => {
    // 1. Background putih & border biru
    ctx.fillStyle = "#FFFFFF";
    ctx.fillRect(0, 0, 800, 1120);

    // Gelombang kiri atas (biru & kuning)
    ctx.fillStyle = "#FFD100";
    ctx.beginPath();
    ctx.moveTo(0, 0);
    ctx.lineTo(280, 0);
    ctx.quadraticCurveTo(100, 40, 0, 160);
    ctx.closePath();
    ctx.fill();

    ctx.fillStyle = "#0B4DA2";
    ctx.beginPath();
    ctx.moveTo(0, 0);
    ctx.lineTo(240, 0);
    ctx.quadraticCurveTo(90, 25, 0, 130);
    ctx.closePath();
    ctx.fill();

    // Gelombang kanan bawah (biru & kuning)
    ctx.fillStyle = "#FFD100";
    ctx.beginPath();
    ctx.moveTo(0, 1120);
    ctx.lineTo(0, 1045);
    ctx.quadraticCurveTo(400, 1015, 800, 980);
    ctx.lineTo(800, 1120);
    ctx.closePath();
    ctx.fill();

    ctx.fillStyle = "#0B4DA2";
    ctx.beginPath();
    ctx.moveTo(0, 1120);
    ctx.lineTo(0, 1070);
    ctx.quadraticCurveTo(420, 1045, 800, 1010);
    ctx.lineTo(800, 1120);
    ctx.closePath();
    ctx.fill();

    // Border luar biru
    ctx.strokeStyle = "#0B4DA2";
    ctx.lineWidth = 6;
    roundRect(12, 12, 776, 1096, 24);
    ctx.stroke();

    // 2. Logo BLOK 9 BSR
    if (logoImg) {
      ctx.drawImage(logoImg, 310, 35, 180, 150);
    } else {
      ctx.fillStyle = "#0B2E59";
      ctx.font = "900 44px sans-serif";
      ctx.textAlign = "center";
      ctx.fillText("BLOK 9 BSR", 400, 125);
    }

    // 3. Judul KWITANSI & Badge
    ctx.fillStyle = "#0B2447";
    ctx.font = "900 64px sans-serif";
    ctx.textAlign = "center";
    ctx.fillText("KWITANSI", 400, 255);

    // Garis kuning kiri & kanan badge
    ctx.fillStyle = "#F5B800";
    ctx.fillRect(55, 293, 75, 5);
    ctx.fillRect(670, 293, 75, 5);

    // Pill badge biru border kuning
    ctx.fillStyle = "#0B57A4";
    roundRect(145, 270, 510, 50, 25);
    ctx.fill();
    ctx.strokeStyle = "#F5B800";
    ctx.lineWidth = 4;
    ctx.stroke();

    ctx.fillStyle = "#FFFFFF";
    ctx.font = "bold 26px sans-serif";
    ctx.fillText("Pembayaran Iuran Bulanan", 400, 304);

    // 4. Kotak Info Warga
    ctx.fillStyle = "#F4F9FF";
    roundRect(45, 345, 710, 235, 18);
    ctx.fill();
    ctx.strokeStyle = "#2E75B6";
    ctx.lineWidth = 2;
    ctx.stroke();

    const nama = trx.nama || wargaData.nama || "-";
    const blok = trx.blok || wargaData.blok || "-";
    const hp = trx.hp || wargaData.hp || "-";
    const bulanArr = trx.bulan_dibayar || [];
    const jmlBulan = bulanArr.length || 1;
    const bulanText = bulanArr.length > 0 ? bulanArr.join(", ") : "-";
    const tglBayar = new Date().toLocaleDateString("id-ID", {
      day: "numeric",
      month: "long",
      year: "numeric",
    });
    const metode = trx.metode || "Transfer / QRIS";

    const infoRows = [
      ["Nama", nama],
      ["Blok", `(${blok})`],
      ["HP", hp],
      ["Bulan Dibayar", bulanText],
      ["Tanggal Bayar", tglBayar],
      ["Metode Bayar", metode],
    ];

    ctx.textAlign = "left";
    infoRows.forEach(([label, val], idx) => {
      const y = 385 + idx * 35;
      ctx.fillStyle = "#1E293B";
      ctx.font = "bold 22px sans-serif";
      ctx.fillText(label, 85, y);
      ctx.fillText(":", 315, y);
      ctx.fillStyle = "#0F172A";
      ctx.fillText(String(val).slice(0, 34), 340, y);
    });

    // 5. Tabel Rincian Nominal
    const kasPaguyuban = 20000 * jmlBulan;
    const kasBlok9 = 5000 * jmlBulan;
    const total = Number(trx.nominal || kasPaguyuban + kasBlok9);
    const fmt = (n) => "Rp " + Number(n).toLocaleString("id-ID");

    // Header tabel
    ctx.fillStyle = "#0B4DA2";
    roundRect(45, 605, 710, 50, 12);
    ctx.fill();
    ctx.fillRect(45, 635, 710, 20);

    ctx.fillStyle = "#FFFFFF";
    ctx.font = "bold 24px sans-serif";
    ctx.textAlign = "center";
    ctx.fillText("Uraian", 265, 638);
    ctx.fillText("Nominal", 620, 638);

    // Baris 1 & 2
    ctx.fillStyle = "#FFFFFF";
    ctx.fillRect(45, 655, 710, 100);

    ctx.fillStyle = "#0F172A";
    ctx.font = "bold 23px sans-serif";
    ctx.textAlign = "left";
    ctx.fillText(`Kas Paguyuban BSR (20k x ${jmlBulan})`, 70, 690);
    ctx.fillText(`Kas BLOK 9 (5k x ${jmlBulan})`, 70, 740);

    ctx.textAlign = "right";
    ctx.fillText(fmt(kasPaguyuban), 725, 690);
    ctx.fillText(fmt(kasBlok9), 725, 740);

    // Baris Total (Kuning)
    ctx.fillStyle = "#FFD100";
    roundRect(45, 755, 710, 60, 12);
    ctx.fill();
    ctx.fillRect(45, 755, 710, 20);

    ctx.fillStyle = "#0B2447";
    ctx.font = "900 28px sans-serif";
    ctx.textAlign = "center";
    ctx.fillText("Total Tagihan", 265, 795);
    ctx.textAlign = "right";
    ctx.fillText(fmt(total), 725, 795);

    // Garis tabel
    ctx.strokeStyle = "#0B4DA2";
    ctx.lineWidth = 2;
    roundRect(45, 605, 710, 210, 12);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(490, 605);
    ctx.lineTo(490, 815);
    ctx.moveTo(45, 705);
    ctx.lineTo(755, 705);
    ctx.moveTo(45, 755);
    ctx.lineTo(755, 755);
    ctx.stroke();

    // 6. Stempel LUNAS
    ctx.save();
    ctx.translate(400, 945);
    ctx.rotate(-0.08);

    ctx.strokeStyle = "#0B4DA2";
    ctx.lineWidth = 6;
    ctx.beginPath();
    ctx.arc(0, 0, 95, 0, Math.PI * 2);
    ctx.stroke();

    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.arc(0, 0, 82, 0, Math.PI * 2);
    ctx.stroke();

    ctx.fillStyle = "#0B4DA2";
    roundRect(-125, -30, 250, 60, 14);
    ctx.fill();

    ctx.fillStyle = "#FFFFFF";
    ctx.font = "900 38px sans-serif";
    ctx.textAlign = "center";
    ctx.fillText("✓ LUNAS", 0, 13);

    ctx.fillStyle = "#0B2447";
    ctx.font = "bold 18px sans-serif";
    ctx.fillText("BLOK 9 BSR", 0, -48);
    ctx.fillText("PEMBAYARAN", 0, 62);
    ctx.restore();

    // Download file PNG
    const safeName = String(nama).replace(/[^a-zA-Z0-9]/g, "_");
    const link = document.createElement("a");
    link.download = `Kwitansi_Blok9_${safeName}.png`;
    link.href = canvas.toDataURL("image/png");
    link.click();
  };

  // Muat logo /icon.svg terlebih dahulu
  const img = new Image();
  img.crossOrigin = "anonymous";
  img.onload = () => renderCanvas(img);
  img.onerror = () => renderCanvas(null);
  img.src = "/icon.svg";
};

// ==========================================
// ADMIN: DENAH INTERAKTIF BLOK 9
// ==========================================
function AdminDenahBlok({ wargaOnly = [], pembayaranList = [], showToast, isAdmin = true }) {
  const [selectedWarga, setSelectedWarga] = useState(null);
  const [filterStatus, setFilterStatus] = useState("SEMUA");
  const [cariDenah, setCariDenah] = useState("");
  const [showFormCash, setShowFormCash] = useState(false);
  const [bulanDipilih, setBulanDipilih] = useState([]);
  const [ketCash, setKetCash] = useState("");
  const [savingCash, setSavingCash] = useState(false);

  const toggleBulanCash = (ym) => {
    setBulanDipilih((prev) =>
      prev.includes(ym) ? prev.filter((b) => b !== ym) : [...prev, ym]
    );
  };

  const handleBayarCash = async () => {
    if (!selectedWarga || bulanDipilih.length === 0) return;
    setSavingCash(true);
    try {
      const totalNominal = bulanDipilih.length * iuran;
      await addDoc(collection(db, "pembayaran"), {
        warga_uid: selectedWarga.uid || selectedWarga.id,
        wargaId: selectedWarga.id,
        nama: selectedWarga.nama || "",
        blok: selectedWarga.blok || "",
        bulan_dibayar: bulanDipilih,
        jumlahBulan: bulanDipilih.length,
        nominal: totalNominal,
        metode: "Cash",
        status: "Lunas",
        keterangan: ketCash.trim() || "Pembayaran Cash via Denah Admin",
        tanggalBayar: new Date().toISOString(),
        createdAt: serverTimestamp(),
      });
      if (showToast) showToast(`Pembayaran Cash ${selectedWarga.nama} berhasil disimpan!`);
      setShowFormCash(false);
      setBulanDipilih([]);
      setKetCash("");
      setSelectedWarga(null);
    } catch (err) {
      console.error(err);
      if (showToast) showToast("Gagal menyimpan pembayaran cash!", "error");
    } finally {
      setSavingCash(false);
    }
  };

  const now = new Date();
  const tahunIni = now.getFullYear();
  const bulanIni = now.getMonth() + 1; // 1 - 12
  const iuran = (typeof CONFIG !== "undefined" && CONFIG.iuranPerBulan) || 25000;

  const NAMA_BULAN = [
    "Jan", "Feb", "Mar", "Apr", "Mei", "Jun",
    "Jul", "Agu", "Sep", "Okt", "Nov", "Des",
  ];
  const fmtYm = (ym) => {
    const [y, m] = String(ym).split("-");
    return `${NAMA_BULAN[Number(m) - 1] || m} ${y}`;
  };
  const fmtRp = (n) => "Rp " + Number(n || 0).toLocaleString("id-ID");

  // Daftar bulan dari Januari sampai bulan saat ini
  const bulanWajib = [];
  for (let m = 1; m <= bulanIni; m++) {
    bulanWajib.push(`${tahunIni}-${String(m).padStart(2, "0")}`);
  }

  // Ambil angka nomor rumah paling belakang untuk pengurutan (misal "Blok 9 / 12" -> 12)
  const getNoRumah = (blok) => {
    const match = String(blok || "").match(/(\d+)\s*$/);
    return match ? parseInt(match[1], 10) : 9999;
  };

  // Hitung status & tunggakan tiap rumah
  const dataDenah = wargaOnly
    .map((w) => {
      const uid = w.uid || w.id;
      const trxWarga = pembayaranList.filter((p) => p.warga_uid === uid);

      const bulanLunas = new Set();
      const bulanProses = new Set();

      trxWarga.forEach((t) => {
        (t.bulan_dibayar || []).forEach((ym) => {
          if (t.status === "Lunas") bulanLunas.add(ym);
          else if (t.status === "Menunggu Verifikasi") bulanProses.add(ym);
        });
      });

      const bulanNunggak = bulanWajib.filter((ym) => !bulanLunas.has(ym));
      const isTidakAktif = w.status === "Tidak Aktif";

      return {
        ...w,
        noRumah: getNoRumah(w.blok),
        bulanNunggak: isTidakAktif ? [] : bulanNunggak,
        bulanProses: Array.from(bulanProses),
        totalNunggak: isTidakAktif ? 0 : bulanNunggak.length * iuran,
        isTidakAktif,
      };
    })
    .sort((a, b) => a.noRumah - b.noRumah);

  // Filter pencarian & status
  const filteredDenah = dataDenah.filter((w) => {
    if (filterStatus === "LUNAS" && (w.bulanNunggak.length > 0 || w.isTidakAktif)) return false;
    if (filterStatus === "NUNGGAK" && (w.bulanNunggak.length === 0 || w.isTidakAktif)) return false;
    if (cariDenah.trim()) {
      const q = cariDenah.trim().toLowerCase();
      const cocok =
        (w.nama || "").toLowerCase().includes(q) ||
        (w.blok || "").toLowerCase().includes(q);
      if (!cocok) return false;
    }
    return true;
  });

  const jmlLunas = dataDenah.filter((w) => !w.isTidakAktif && w.bulanNunggak.length === 0).length;
  const jmlNunggak = dataDenah.filter((w) => !w.isTidakAktif && w.bulanNunggak.length > 0).length;

  const getWarnaKartu = (w) => {
    if (w.isTidakAktif) return "bg-gray-100 border-gray-300 text-gray-500";
    if (w.bulanNunggak.length === 0) return "bg-emerald-50 border-emerald-400 text-emerald-900 hover:bg-emerald-100";
    if (w.bulanNunggak.length <= 2) return "bg-amber-50 border-amber-400 text-amber-900 hover:bg-amber-100";
    return "bg-red-50 border-red-400 text-red-900 hover:bg-red-100";
  };

  const getBadgeStatus = (w) => {
    if (w.isTidakAktif) return { text: "Nonaktif", cls: "bg-gray-300 text-gray-700" };
    if (w.bulanNunggak.length === 0) return { text: "LUNAS", cls: "bg-emerald-600 text-white" };
    if (w.bulanNunggak.length <= 2) return { text: `-${w.bulanNunggak.length} Bln`, cls: "bg-amber-500 text-white" };
    return { text: `-${w.bulanNunggak.length} Bln`, cls: "bg-red-600 text-white" };
  };

  const bukaWA = (w) => {
    const teksPesan = isAdmin
      ? `Assalamu'alaikum Bapak/Ibu *${w.nama}* (${w.blok}),\n\n` +
        `Mengingatkan iuran kas BLOK 9 BSR sampai bulan *${NAMA_BULAN[bulanIni - 1]} ${tahunIni}* masih ada tunggakan *${w.bulanNunggak.length} bulan* sebesar *${fmtRp(w.totalNunggak)}*.\n` +
        `Rincian: ${w.bulanNunggak.map(fmtYm).join(", ")}.\n\nTerima kasih 🙏`
      : `Assalamu'alaikum Bapak/Ibu *${w.nama}* (${w.blok}),\n\n` +
        `Apa kabar? Salam silaturahmi sesama warga BLOK 9 BSR 🙏`;

    const pesan = encodeURIComponent(teksPesan);
    let hp = String(w.hp || "").replace(/[^0-9]/g, "");
    if (hp.startsWith("0")) hp = "62" + hp.slice(1);
    window.open(hp ? `https://wa.me/${hp}?text=${pesan}` : `https://wa.me/?text=${pesan}`, "_blank");
  };

  return (
    <div className="space-y-4">
      {/* Header & Ringkasan */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h3 className="text-lg sm:text-xl font-bold text-gray-900">Denah Interaktif BLOK 9</h3>
          <p className="text-xs sm:text-sm text-gray-500">
            Status iuran s/d <b>{NAMA_BULAN[bulanIni - 1]} {tahunIni}</b>. Klik rumah untuk cek detail tunggakan.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <input
            type="text"
            value={cariDenah}
            onChange={(e) => setCariDenah(e.target.value)}
            placeholder="Cari nama / no rumah..."
            className="w-full sm:w-48 px-3 py-2 rounded-xl border border-gray-300 text-xs bg-white focus:ring-2 focus:ring-orange-500 outline-none"
          />
          <div className="flex rounded-xl overflow-hidden border border-gray-300 text-xs font-bold">
            {[
              ["SEMUA", `Semua (${dataDenah.length})`],
              ["LUNAS", `Lunas (${jmlLunas})`],
              ["NUNGGAK", `Nunggak (${jmlNunggak})`],
            ].map(([val, label]) => (
              <button
                key={val}
                onClick={() => setFilterStatus(val)}
                className={`px-3 py-2 ${
                  filterStatus === val ? "bg-orange-500 text-white" : "bg-white text-gray-600 hover:bg-gray-50"
                }`}
              >
                {label}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* Legenda Warna */}
      <div className="flex flex-wrap items-center gap-3 text-[11px] sm:text-xs bg-white p-3 rounded-xl border border-gray-200">
        <span className="flex items-center gap-1.5"><span className="w-3 h-3 rounded bg-emerald-500 inline-block" /> Lunas</span>
        <span className="flex items-center gap-1.5"><span className="w-3 h-3 rounded bg-amber-400 inline-block" /> Nunggak 1–2 Bln</span>
        <span className="flex items-center gap-1.5"><span className="w-3 h-3 rounded bg-red-500 inline-block" /> Nunggak ≥3 Bln</span>
        <span className="flex items-center gap-1.5"><span className="w-3 h-3 rounded bg-gray-300 inline-block" /> Tidak Aktif</span>
      </div>

      {/* Denah 2 Baris Rapi + Kompas Utara & Perempatan Lurus */}
      <div className="bg-white p-3 sm:p-4 rounded-2xl border border-slate-200 shadow-sm">
        {filteredDenah.length === 0 ? (
          <div className="text-center py-8 text-sm text-gray-400">
            Data rumah tidak ditemukan.
          </div>
        ) : (
          <div className="overflow-x-auto pb-2">
            {(() => {
              const KOLOM_GANG = 17; // Posisi gang setelah No. 33 & 34

              const kolomMap = {};
              filteredDenah.forEach((w) => {
                const no = w.noRumah !== 9999 ? w.noRumah : 0;
                const col = Math.ceil(no / 2);
                if (!kolomMap[col]) kolomMap[col] = { atas: null, bawah: null };
                if (no % 2 !== 0) kolomMap[col].atas = w;
                else kolomMap[col].bawah = w;
              });

              const semuaKolom = Object.keys(kolomMap)
                .map(Number)
                .sort((a, b) => b - a);

              const kolomKiri = semuaKolom.filter((c) => c > KOLOM_GANG);
              const kolomKanan = semuaKolom.filter((c) => c <= KOLOM_GANG);

              const renderRumah = (w) => {
                if (!w) {
                  return <div className="w-24 sm:w-28 h-[86px] shrink-0" />;
                }
                const badge = getBadgeStatus(w);
                return (
                  <button
                    key={w.id}
                    onClick={() => setSelectedWarga(w)}
                    className={`w-24 sm:w-28 h-[86px] shrink-0 p-2 rounded-xl border text-left transition-all shadow-sm hover:shadow active:scale-95 flex flex-col justify-between ${getWarnaKartu(w)}`}
                  >
                    <div className="flex items-center justify-between gap-1 w-full">
                      <span className="text-xs sm:text-sm font-black">
                        {w.noRumah !== 9999 ? w.noRumah : "-"}
                      </span>
                      <span className={`text-[9px] px-1.5 py-0.5 rounded-md font-bold ${badge.cls}`}>
                        {badge.text}
                      </span>
                    </div>
                    <div className="mt-1">
                      <div className="text-[10px] sm:text-xs font-bold leading-tight line-clamp-2 uppercase">
                        {w.nama || "-"}
                      </div>
                    </div>
                  </button>
                );
              };

              return (
                <div className="inline-flex items-stretch min-w-full">
                  {/* KOMPAS UTARA */}
                  <div className="w-11 shrink-0 flex flex-col items-center justify-start pt-2 pr-2 select-none">
                    <span className="text-[11px] font-black text-slate-800 tracking-wider leading-none mb-1">
                      UTARA
                    </span>
                    <svg viewBox="0 0 24 28" className="w-6 h-7 text-red-600 fill-current">
                      <path d="M12 1 L22 25 L12 18 L2 25 Z" />
                    </svg>
                  </div>

                  {/* SISI KIRI GANG (TAMAN + RUMAH KIRI + AREA MASJID) */}
                  {kolomKiri.length > 0 && (
                    <div className="flex flex-col shrink-0">
                      {/* Baris Atas Kiri: Taman + Jalan Blok 10 + Ganjil */}
                      <div className="flex items-stretch">
                        <div className="w-14 shrink-0 mr-2 flex flex-col">
                          <div className="flex-1 rounded-xl bg-emerald-50 border-2 border-emerald-300 flex items-center justify-center p-1 shadow-sm">
                            <span className="text-[10px] font-black text-emerald-800 tracking-widest uppercase [writing-mode:vertical-rl] rotate-180">
                              🌳 TAMAN
                            </span>
                          </div>
                        </div>

                        <div className="flex flex-col flex-1">
                          <div className="h-7 bg-slate-200/80 px-6 flex items-center justify-around text-[10px] font-bold tracking-[0.25em] text-slate-500 uppercase rounded-tl-md">
                            <span>JALAN BLOK 10</span>
                            <span>JALAN BLOK 10</span>
                          </div>
                          <div className="flex items-stretch justify-end gap-2 py-2.5 pr-2.5 border-b border-dashed border-slate-300">
                            {kolomKiri.map((col) => (
                              <div key={`atas-${col}`}>{renderRumah(kolomMap[col].atas)}</div>
                            ))}
                          </div>
                        </div>
                      </div>

                      {/* Jalan Blok 9 (Sisi Kiri) */}
                      <div className="h-9 bg-slate-200/80 flex items-center">
                        <div className="w-16 shrink-0" />
                        <div className="flex-1 px-6 flex items-center justify-around text-[11px] font-extrabold tracking-[0.25em] text-slate-600 uppercase">
                          <span>JALAN BLOK 9</span>
                          <span>JALAN BLOK 9</span>
                        </div>
                      </div>

                      {/* Baris Bawah Kiri: Jalan Ujung + Genap + Area Masjid */}
                      <div className="flex items-stretch">
                        <div className="w-14 shrink-0 mr-2 bg-slate-200/80 rounded-b-md" />
                        <div className="flex flex-col flex-1">
                          <div className="flex items-stretch justify-end gap-2 py-2.5 pr-2.5 border-t border-dashed border-slate-300">
                            {kolomKiri.map((col) => (
                              <div key={`bawah-${col}`}>{renderRumah(kolomMap[col].bawah)}</div>
                            ))}
                            {kolomKanan.length > 0 && (
                              <div className="w-44 sm:w-52 h-[86px] shrink-0 rounded-xl bg-sky-50 border-2 border-sky-400 flex flex-col items-center justify-center px-2 text-center shadow-sm">
                                <span className="text-base">🕌</span>
                                <span className="text-[11px] sm:text-xs font-black text-sky-900 uppercase tracking-wider">
                                  AREA MASJID
                                </span>
                              </div>
                            )}
                          </div>
                          <div className="h-7" />
                        </div>
                      </div>
                    </div>
                  )}

                  {/* GANG TENGAH (SATU JALUR LURUS DARI ATAS KE BAWAH) */}
                  {kolomKiri.length > 0 && kolomKanan.length > 0 && (
                    <div className="w-10 shrink-0 bg-slate-200/80 rounded-bl-md" />
                  )}

                  {/* SISI KANAN GANG (NO. 33/34 SAMPAI NO. 1/2) */}
                  {kolomKanan.length > 0 && (
                    <div className="flex flex-col shrink-0">
                      {/* Jalan Blok 10 (Sisi Kanan) */}
                      <div className="h-7 bg-slate-200/80 px-6 flex items-center justify-around text-[10px] font-bold tracking-[0.25em] text-slate-500 uppercase">
                        <span>JALAN BLOK 10</span>
                        <span>JALAN BLOK 10</span>
                        <span>JALAN BLOK 10</span>
                      </div>

                      {/* Baris Atas Kanan (Ganjil) */}
                      <div className="flex items-stretch gap-2 py-2.5 px-2.5 border-b border-dashed border-slate-300">
                        {kolomKanan.map((col) => (
                          <div key={`atas-${col}`}>{renderRumah(kolomMap[col].atas)}</div>
                        ))}
                      </div>

                      {/* Jalan Blok 9 (Sisi Kanan) */}
                      <div className="h-9 bg-slate-200/80 px-6 flex items-center justify-around text-[11px] font-extrabold tracking-[0.25em] text-slate-600 uppercase">
                        <span>JALAN BLOK 9</span>
                        <span>JALAN BLOK 9</span>
                        <span>JALAN BLOK 9</span>
                      </div>

                      {/* Baris Bawah Kanan (Genap) */}
                      <div className="flex items-stretch gap-2 py-2.5 px-2.5 border-t border-dashed border-slate-300">
                        {kolomKanan.map((col) => (
                          <div key={`bawah-${col}`}>{renderRumah(kolomMap[col].bawah)}</div>
                        ))}
                      </div>

                      {/* Jalan Blok 8 (Sisi Kanan) */}
                      <div className="h-7 bg-slate-200/80 px-6 flex items-center justify-around text-[10px] font-bold tracking-[0.25em] text-slate-500 uppercase">
                        <span>JALAN BLOK 8</span>
                        <span>JALAN BLOK 8</span>
                        <span>JALAN BLOK 8</span>
                      </div>
                    </div>
                  )}

                  {/* JALAN UJUNG KANAN (DEPAN NO. 1 & 2) */}
                  <div className="w-8 shrink-0 bg-slate-200/80 rounded-r-md flex items-center justify-center">
                    <span className="text-[10px] font-extrabold text-slate-600 tracking-widest uppercase [writing-mode:vertical-rl] rotate-180">
                      JALAN
                    </span>
                  </div>
                </div>
              );
            })()}
          </div>
        )}
      </div>

      {/* Popup Detail Saat Rumah Diklik */}
      {selectedWarga && (
        <div
          className="fixed inset-0 z-50 bg-black/50 flex items-end sm:items-center justify-center p-0 sm:p-4"
          onClick={() => setSelectedWarga(null)}
        >
          <div
            className="bg-white w-full sm:max-w-md rounded-t-2xl sm:rounded-2xl p-5 space-y-4 shadow-xl"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-start justify-between gap-2 border-b pb-3">
              <div>
                <span className="text-xs font-bold text-orange-600 uppercase">{selectedWarga.blok}</span>
                <h4 className="text-lg font-extrabold text-gray-900">{selectedWarga.nama}</h4>
                <p className="text-xs text-gray-500">HP: {selectedWarga.hp || "-"}</p>
              </div>
              <button
                onClick={() => setSelectedWarga(null)}
                className="w-8 h-8 rounded-full bg-gray-100 hover:bg-gray-200 font-bold text-gray-600"
              >
                ✕
              </button>
            </div>

            {selectedWarga.isTidakAktif ? (
              <div className="p-4 rounded-xl bg-gray-100 text-center text-sm font-bold text-gray-600">
                Status Rumah: Tidak Aktif
              </div>
            ) : selectedWarga.bulanNunggak.length === 0 ? (
              <div className="p-4 rounded-xl bg-emerald-50 border border-emerald-200 text-center">
                <div className="text-emerald-700 font-extrabold text-base">✅ LUNAS</div>
                <div className="text-xs text-emerald-600 mt-1">
                  Sudah lunas sampai {NAMA_BULAN[bulanIni - 1]} {tahunIni}
                </div>
                {!isAdmin && (
  <button
    onClick={() => bukaWA(selectedWarga)}
    className="mt-3 w-full py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-sm"
  >
    Sapa Warga
  </button>
)}
              </div>
            ) : (
              <div className="space-y-3">
                <div className="p-3.5 rounded-xl bg-red-50 border border-red-200 flex items-center justify-between">
                  <div>
                    <div className="text-xs text-red-600 font-semibold">Total Tunggakan</div>
                    <div className="text-lg font-black text-red-700">
                      {selectedWarga.bulanNunggak.length} Bulan
                    </div>
                  </div>
                  <div className="text-right">
                    <div className="text-xs text-red-600 font-semibold">Nominal</div>
                    <div className="text-lg font-black text-red-700">
                      {fmtRp(selectedWarga.totalNunggak)}
                    </div>
                  </div>
                </div>

                <div>
                  <div className="text-xs font-bold text-gray-600 mb-1.5">
                    Rincian Bulan Belum Bayar (s/d {NAMA_BULAN[bulanIni - 1]} {tahunIni}):
                  </div>
                  <div className="flex flex-wrap gap-1.5">
                    {selectedWarga.bulanNunggak.map((ym) => (
                      <span
                        key={ym}
                        className="px-2.5 py-1 rounded-lg bg-red-100 text-red-700 text-xs font-bold"
                      >
                        {fmtYm(ym)}
                      </span>
                    ))}
                  </div>
                </div>
{isAdmin && (
                  <div className="pt-2 border-t border-gray-200 space-y-2.5">
                    {!showFormCash ? (
                      <button
                        onClick={() => {
                          setBulanDipilih([selectedWarga.bulanNunggak[0]]);
                          setKetCash("");
                          setShowFormCash(true);
                        }}
                        className="w-full py-2.5 rounded-xl bg-orange-500 hover:bg-orange-600 text-white font-bold text-sm"
                      >
                        💵 Bayar Cash
                      </button>
                    ) : (
                      <div className="p-3 rounded-xl bg-orange-50 border border-orange-200 space-y-2.5">
                        <div className="text-xs font-extrabold text-orange-800">
                          Pilih Bulan Bayar Cash:
                        </div>

                        {/* Pilihan Bulan Tunggakan */}
                        <div className="flex flex-wrap gap-1.5">
                          {selectedWarga.bulanNunggak.map((ym) => {
                            const aktif = bulanDipilih.includes(ym);
                            return (
                              <button
                                key={ym}
                                type="button"
                                onClick={() => toggleBulanCash(ym)}
                                className={`px-2.5 py-1 rounded-lg text-[11px] font-bold border transition ${
                                  aktif
                                    ? "bg-orange-500 text-white border-orange-600"
                                    : "bg-white text-gray-700 border-gray-300"
                                }`}
                              >
                                {aktif ? "✓ " : ""}{fmtYm(ym)}
                              </button>
                            );
                          })}
                        </div>

                        {/* Total Nominal Otomatis */}
                        <div className="text-xs font-bold text-gray-700 flex justify-between">
                          <span>Total ({bulanDipilih.length} bln):</span>
                          <span className="text-orange-700 font-black">
                            {fmtRp(bulanDipilih.length * iuran)}
                          </span>
                        </div>

                        {/* Input Keterangan */}
                        <input
                          type="text"
                          value={ketCash}
                          onChange={(e) => setKetCash(e.target.value)}
                          placeholder="Keterangan (misal: Titip Pak RT / Lunas Tunai)..."
                          className="w-full px-3 py-2 rounded-lg border border-orange-300 text-xs bg-white focus:outline-none focus:ring-2 focus:ring-orange-400"
                        />

                        {/* Tombol Simpan & Batal */}
                        <div className="flex gap-2">
                          <button
                            type="button"
                            onClick={() => setShowFormCash(false)}
                            className="flex-1 py-2 rounded-lg bg-gray-200 hover:bg-gray-300 text-gray-700 font-bold text-xs"
                          >
                            Batal
                          </button>
                          <button
                            type="button"
                            disabled={savingCash || bulanDipilih.length === 0}
                            onClick={handleBayarCash}
                            className="flex-1 py-2 rounded-lg bg-orange-600 hover:bg-orange-700 disabled:opacity-50 text-white font-bold text-xs"
                          >
                            {savingCash ? "Menyimpan..." : "Simpan Cash"}
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                )}
                <button
                  onClick={() => bukaWA(selectedWarga)}
                  className="w-full py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-sm"
                >
                  {isAdmin ? "💬 Ingatkan via WhatsApp" : "Sapa Warga"}
                </button>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
// ==========================================
// KOMPONEN UTAMA
// ==========================================
export default function App() {
  const [user, setUser] = useState(null);
  const [userProfile, setUserProfile] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [impersonatedWarga, setImpersonatedWarga] = useState(null);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState("");

  const cocokPencarian = (item) => {
    if (!searchQuery.trim()) return true;
    const q = searchQuery.toLowerCase();
    const nama = (item.nama || item.namaWarga || "").toLowerCase();
    const rumah = (item.noRumah || item.rumah || item.username || "").toLowerCase();
    return nama.includes(q) || rumah.includes(q);
  };
  // Login State
  const [loginUsername, setLoginUsername] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [loginBusy, setLoginBusy] = useState(false);

  // Live Clock
  const [nowTime, setNowTime] = useState(new Date());
  useEffect(() => {
    const t = setInterval(() => setNowTime(new Date()), 1000);
    return () => clearInterval(t);
  }, []);

  // Firestore Real-Time Collections
  const [wargaList, setWargaList] = useState([]);
  const [pembayaranList, setPembayaranList] = useState([]);
  const [setorKasList, setSetorKasList] = useState([]);
  const [pengumumanList, setPengumumanList] = useState([]);
  const [tataTertibList, setTataTertibList] = useState([]);

  // UI State
  const [activeTab, setActiveTab] = useState("dashboard");
  const [previewImg, setPreviewImg] = useState(null);
  const [toast, setToast] = useState(null);

  const showToast = (msg, type = "success") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3500);
  };

  // Auth Listener
  useEffect(() => {
    return onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser);
      if (!currentUser) {
        setUserProfile(null);
        setAuthLoading(false);
      }
    });
  }, []);

  // Real-Time Firestore Listeners
  useEffect(() => {
    if (!user) return;

    const u1 = onSnapshot(collection(db, "data_warga"), (snap) => {
      const data = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
      data.sort((a, b) => (a.nomor_rumah || 999) - (b.nomor_rumah || 999));
      setWargaList(data);

      const me = data.find((w) => w.uid === user.uid || w.id === user.uid);
      if (me) {
        setUserProfile(me);
      } else if (user.email === "admin@blok9bsr.com") {
        setUserProfile({
          uid: user.uid,
          nama: "Administrator",
          blok: "Pengurus",
          username: "admin",
          role: "admin",
        });
      }
      setAuthLoading(false);
    });

    const u2 = onSnapshot(collection(db, "pembayaran_iuran"), (snap) => {
      const data = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
      data.sort((a, b) => (b.waktu || "").localeCompare(a.waktu || ""));
      setPembayaranList(data);
    });

    const u3 = onSnapshot(collection(db, "setor_kas"), (snap) => {
      const data = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
      data.sort((a, b) => (b.tanggal || "").localeCompare(a.tanggal || ""));
      setSetorKasList(data);
    });

    const u4 = onSnapshot(collection(db, "pengumuman"), (snap) => {
      const data = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
      data.sort((a, b) => (b.tanggal || "").localeCompare(a.tanggal || ""));
      setPengumumanList(data);
    });

    const u5 = onSnapshot(collection(db, "tata_tertib"), (snap) => {
      const data = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
      data.sort((a, b) => (a.urutan || 99) - (b.urutan || 99));
      setTataTertibList(data);
    });

    return () => {
      u1(); u2(); u3(); u4(); u5();
    };
  }, [user]);

  // Kalkulasi Saldo Real-Time (Kas Paguyuban BSR & Kas Khusus BLOK 9)
  const saldoSummary = useMemo(() => {
    const lunas = pembayaranList.filter((p) => p.status === "Lunas");
    const masukPag = lunas.reduce((s, p) => s + Number(p.kas_paguyuban || 0), 0);
    const masukB9 = lunas.reduce((s, p) => s + Number(p.kas_blok_9 || 0), 0);

    const keluarPag = setorKasList
      .filter((s) => s.jenis_kas === "paguyuban" && s.tipe === "Pengeluaran")
      .reduce((s, x) => s + Number(x.nominal || 0), 0);
    const tambahPag = setorKasList
      .filter((s) => s.jenis_kas === "paguyuban" && s.tipe === "Pemasukan")
      .reduce((s, x) => s + Number(x.nominal || 0), 0);

    const keluarB9 = setorKasList
      .filter((s) => s.jenis_kas === "blok9" && s.tipe === "Pengeluaran")
      .reduce((s, x) => s + Number(x.nominal || 0), 0);
    const tambahB9 = setorKasList
      .filter((s) => s.jenis_kas === "blok9" && s.tipe === "Pemasukan")
      .reduce((s, x) => s + Number(x.nominal || 0), 0);

    return {
      saldoPaguyuban: masukPag + tambahPag - keluarPag,
      saldoBlok9: masukB9 + tambahB9 - keluarB9,
    };
  }, [pembayaranList, setorKasList]);

  const pendingCount = useMemo(
    () => pembayaranList.filter((p) => p.status === "Menunggu Verifikasi").length,
    [pembayaranList]
  );

  const handleLogin = async (e) => {
    e.preventDefault();
    setLoginError("");
    setLoginBusy(true);
    try {
      await signInWithEmailAndPassword(
        auth,
        toFirebaseEmail(loginUsername),
        toFirebasePassword(loginPassword)
      );
    } catch {
      setLoginError("Username atau password salah.");
    } finally {
      setLoginBusy(false);
    }
  };

  const handleLogout = async () => {
    await signOut(auth);
    setActiveTab("dashboard");
  };

  const clockLabel = `${nowTime.toLocaleDateString("id-ID", {
    weekday: "long",
    day: "numeric",
    month: "long",
    year: "numeric",
  })}  |  ${nowTime.toLocaleTimeString("id-ID", { hour12: false })} WIB`;

  if (authLoading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-50">
        <RefreshCw className="w-8 h-8 text-orange-500 animate-spin" />
      </div>
    );
  }

  // ==========================================
  // HALAMAN LOGIN
  // ==========================================
 if (!user || !userProfile) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center p-4">
        <div className="w-full max-w-md bg-white rounded-2xl shadow-lg border border-gray-100 p-8">
          <div className="text-center mb-6">
            <img
              src="/icon.svg"
              alt="Logo Blok 9 BSR"
              className="w-16 h-16 mx-auto mb-3 rounded-2xl object-contain"
            />
            <h1 className="text-2xl font-bold text-gray-900">
              BLOK <span className="text-orange-500">9</span> BSR
            </h1>
            <p className="text-sm text-gray-500 mt-1">
              Portal Warga & Sistem Pembayaran Kas
            </p>
          </div>
{loginError && (
            <div className="mb-4 p-3.5 bg-red-50 border border-red-200 text-red-700 text-sm rounded-xl space-y-2.5">
              <div className="flex items-center gap-2">
                <AlertCircle className="w-4 h-4 shrink-0" />
                <span>{loginError}</span>
              </div>
              <a
                href={`https://wa.me/${CONFIG.waBantuan}?text=${encodeURIComponent(
                  `Halo Admin BLOK 9 BSR, saya lupa password / gagal login portal dengan username: ${
                    loginUsername || "-"
                  }. Mohon bantuannya.`
                )}`}
                target="_blank"
                rel="noreferrer"
                className="w-full inline-flex items-center justify-center gap-2 px-3 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold transition-colors"
              >
                <MessageCircle className="w-4 h-4 shrink-0" />
                <span>Lupa Password? Tanya Admin via WhatsApp</span>
              </a>
            </div>
          )}

          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">
                Username
              </label>
              <input
                type="text"
                required
                placeholder="Contoh: bsr1 atau admin"
                value={loginUsername}
                onChange={(e) => setLoginUsername(e.target.value)}
                className="w-full px-4 py-2.5 rounded-xl border border-gray-300 focus:ring-2 focus:ring-orange-500 outline-none"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">
                Password
              </label>
              <input
                type="password"
                required
                placeholder="Masukkan password"
                value={loginPassword}
                onChange={(e) => setLoginPassword(e.target.value)}
                className="w-full px-4 py-2.5 rounded-xl border border-gray-300 focus:ring-2 focus:ring-orange-500 outline-none"
              />
            </div>
            <button
              type="submit"
              disabled={loginBusy}
              className="w-full py-3 bg-orange-500 hover:bg-orange-600 text-white font-semibold rounded-xl transition shadow-sm disabled:opacity-60"
            >
              {loginBusy ? "Memproses..." : "Masuk ke Portal"}
            </button>
          </form>
        </div>
      </div>
    );
  }

const activeProfile = impersonatedWarga || userProfile;
  const isAdmin = userProfile.role === "admin" && !impersonatedWarga;
  const wargaOnly = wargaList.filter((w) => w.role !== "admin");

  const adminMenus = [
    { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { id: "verifikasi", label: "Verifikasi Pembayaran", icon: CheckCircle2, badge: pendingCount },
    { id: "laporan", label: "Laporan Tagihan", icon: FileText },
    { id: "warga", label: "Data Warga", icon: Users },
    { id: "nunggak", label: "Warga Nunggak", icon: AlertTriangle },
    { id: "denah", label: "Denah Blok", icon: Users },
    { id: "setor", label: "Setor Kas", icon: ArrowUpCircle },
    { id: "pengumuman", label: "Pengumuman", icon: Bell },
    { id: "tatib", label: "Tata Tertib", icon: BookOpen },
  ];

  const wargaMenus = [
    { id: "dashboard", label: "Beranda", icon: Home },
    { id: "bayar", label: "Bayar Tagihan", icon: CreditCard },
    { id: "pengumuman", label: "Pengumuman", icon: Bell },
    { id: "tatib", label: "Tata Tertib", icon: BookOpen },
    { id: "denah", label: "Denah Blok", icon: Users },
    { id: "pengaturan", label: "Pengaturan", icon: Settings },
  ];

  const menus = isAdmin ? adminMenus : wargaMenus;
  const currentMenuObj = menus.find((m) => m.id === activeTab) || menus[0];

  return (
    <div className="min-h-screen bg-gray-50 text-gray-800 flex pb-20 md:pb-0">
      {/* Toast */}
      {toast && (
        <div
          className={`fixed top-5 right-5 z-50 px-4 py-3 rounded-xl shadow-lg text-sm font-medium text-white ${
            toast.type === "error" ? "bg-red-600" : "bg-emerald-600"
          }`}
        >
          {toast.msg}
        </div>
      )}

{/* Modal Preview Gambar */}
      {previewImg && (
        <div
          className="fixed inset-0 z-[999] bg-black/80 flex items-center justify-center p-4"
          onClick={() => setPreviewImg(null)}
        >
          <div
            className="bg-white rounded-2xl p-3 max-w-lg w-full"
            onClick={(e) => e.stopPropagation()}
          >
            <img
              src={previewImg}
              alt="Preview"
              className="w-full max-h-[75vh] object-contain rounded-xl"
            />
            <button
              onClick={() => setPreviewImg(null)}
              className="mt-3 w-full py-2.5 bg-gray-900 text-white rounded-xl text-sm font-semibold"
            >
              Tutup
            </button>
          </div>
        </div>
      )}

   {/* OVERLAY MENU HP/TABLET */}
      {mobileMenuOpen && (
        <div
          onClick={() => setMobileMenuOpen(false)}
          className="fixed inset-0 bg-black/50 z-40 lg:hidden"
        />
      )}

      {/* SIDEBAR RESPONSIF (HP/TABLET DRAWER & DESKTOP) */}
      <aside
        className={`fixed inset-y-0 left-0 z-50 lg:sticky lg:top-0 flex flex-col w-64 bg-white border-r border-gray-200 h-screen shrink-0 transform transition-transform duration-200 ease-in-out ${
          mobileMenuOpen ? "translate-x-0" : "-translate-x-full lg:translate-x-0"
        }`}
      >
     <div className="p-5 flex items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <img
              src="/icon.svg"
              alt="Logo Blok 9 BSR"
              className="w-9 h-9 rounded-xl object-contain shrink-0"
            />
            <div>
              <h1 className="font-extrabold text-gray-900 leading-none text-lg">
                BLOK <span className="text-orange-500">9</span>
              </h1>
              <p className="text-xs text-gray-400 mt-0.5">
                {isAdmin ? "Admin Panel" : "Portal Warga"}
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={() => setMobileMenuOpen(false)}
            className="lg:hidden p-1.5 rounded-lg text-gray-500 hover:bg-gray-100"
            aria-label="Tutup Menu"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        <nav className="flex-1 px-3 py-2 space-y-1 overflow-y-auto">
          {menus.map((m) => {
            const Icon = m.icon;
            const active = activeTab === m.id;
            return (
              <button
                key={m.id}
                onClick={() => setActiveTab(m.id)}
                className={`w-full flex items-center justify-between px-3.5 py-2.5 rounded-xl text-sm font-medium transition ${
                  active
                    ? "bg-orange-50 text-orange-600 font-semibold"
                    : "text-gray-600 hover:bg-gray-50"
                }`}
              >
                <div className="flex items-center gap-3">
                  <Icon className="w-4 h-4" />
                  <span>{m.label}</span>
                </div>
                {m.badge > 0 && (
                  <span className="px-2 py-0.5 text-xs rounded-full font-bold bg-red-500 text-white">
                    {m.badge}
                  </span>
                )}
              </button>
            );
          })}
        </nav>
      </aside>

      {/* RIGHT AREA (TOPBAR + MAIN) */}
      <div className="flex-1 flex flex-col min-w-0">
        {/* TOP HEADER */}
       {impersonatedWarga && (
          <div className="bg-orange-600 text-white px-4 sm:px-8 py-2.5 flex flex-wrap items-center justify-between gap-2 sticky top-0 z-50 shadow-md">
            <div className="text-xs sm:text-sm font-medium">
              Mode Akses Admin: Sedang masuk sebagai{" "}
              <b className="underline">
                {impersonatedWarga.nama} ({impersonatedWarga.blok})
              </b>
            </div>
            <button
              type="button"
              onClick={() => {
                setImpersonatedWarga(null);
                setActiveTab("warga");
              }}
              className="px-3 py-1.5 bg-white text-orange-700 hover:bg-orange-50 rounded-lg text-xs font-bold shadow-xs"
            >
              ← Kembali ke Admin Panel
            </button>
          </div>
        )}

        <header className="bg-white border-b border-gray-200 px-4 sm:px-8 py-3.5 flex items-center justify-between gap-2 sticky top-0 z-30">
          <div className="flex items-center gap-2.5 min-w-0">
            <button
              type="button"
              onClick={() => setMobileMenuOpen(true)}
              className="lg:hidden p-2 -ml-2 rounded-xl text-gray-700 hover:bg-gray-100 shrink-0"
              aria-label="Buka Menu"
            >
              <Menu className="w-5 h-5" />
            </button>
            <h2 className="text-base sm:text-lg font-bold text-gray-900 truncate">
              {currentMenuObj.label}
            </h2>
          </div>
          <div className="flex items-center gap-2 sm:gap-3 shrink-0">
            <a
              href={`https://wa.me/${CONFIG.waBantuan}?text=Halo%20Admin%20BLOK%209%20BSR,%20saya%20butuh%20bantuan%20terkait%20portal.`}
              target="_blank"
              rel="noreferrer"
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-emerald-50 text-emerald-700 border border-emerald-200 text-xs font-semibold hover:bg-emerald-100"
            >
              <HelpCircle className="w-4 h-4" />
              <span className="hidden sm:inline">Bantuan</span>
            </a>
            <button
              onClick={handleLogout}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-red-50 text-red-600 border border-red-200 text-xs font-semibold hover:bg-red-100"
            >
              <LogOut className="w-4 h-4" />
              <span className="hidden sm:inline">Keluar</span>
            </button>
            <div className="w-9 h-9 rounded-full bg-orange-100 text-orange-600 font-bold text-sm flex items-center justify-center">
              {(activeProfile.nama || "U").charAt(0).toUpperCase()}
            </div>
          </div>
        </header>

        {/* MAIN CONTENT */}
        <main className="flex-1 p-3 sm:p-6 lg:p-8 max-w-full overflow-x-hidden">
          {isAdmin ? (
            <>
              {activeTab === "dashboard" && (
                <AdminDashboard
                  clockLabel={clockLabel}
                  saldoSummary={saldoSummary}
                  wargaOnly={wargaOnly}
                  pembayaranList={pembayaranList}
                  setActiveTab={setActiveTab}
                />
              )}
              {activeTab === "verifikasi" && (
                <AdminVerifikasi
                  pembayaranList={pembayaranList}
                  setPreviewImg={setPreviewImg}
                  showToast={showToast}
                />
              )}
              {activeTab === "laporan" && (
                <AdminLaporanTagihan
                  pembayaranList={pembayaranList}
                  wargaOnly={wargaOnly}
                />
              )}
            {activeTab === "warga" && (
                <AdminDataWarga
                  wargaOnly={wargaOnly}
                  showToast={showToast}
                  onImpersonateWarga={(w) => {
                    setImpersonatedWarga(w);
                    setActiveTab("dashboard");
                  }}
                />
              )}
              {activeTab === "denah" && (
                <AdminDenahBlok
                wargaOnly={wargaOnly}
                pembayaranList={pembayaranList}
                showToast={showToast}
                isAdmin={true}
              />
            )}
              {activeTab === "nunggak" && (
                <AdminWargaNunggak
                  wargaOnly={wargaOnly}
                  pembayaranList={pembayaranList}
                />
              )}
              {activeTab === "setor" && (
                <AdminSetorKas
                  setorKasList={setorKasList}
                  saldoSummary={saldoSummary}
                  showToast={showToast}
                />
              )}
              {activeTab === "pengumuman" && (
                <PengumumanView
                  isAdmin={true}
                  list={pengumumanList}
                  showToast={showToast}
                />
              )}
              {activeTab === "tatib" && (
                <TataTertibView
                  isAdmin={true}
                  list={tataTertibList}
                  showToast={showToast}
                />
              )}
            </>
          ) : (
            <>
              {activeTab === "dashboard" && (
              <WargaBeranda
                userProfile={userProfile}
                clockLabel={clockLabel}
                pembayaranList={pembayaranList}
                setorKasList={setorKasList}
                setActiveTab={setActiveTab}
              />
            )}
              {activeTab === "bayar" && (
                <WargaBayarTagihan
                  userProfile={userProfile}
                  pembayaranList={pembayaranList}
                  setPreviewImg={setPreviewImg}
                  showToast={showToast}
                />
              )}
              {activeTab === "denah" && (
            <AdminDenahBlok
              wargaOnly={wargaOnly}
              pembayaranList={pembayaranList}
              isAdmin={false}
            />
          )}
              {activeTab === "pengumuman" && (
                <PengumumanView isAdmin={false} list={pengumumanList} />
              )}
              {activeTab === "tatib" && (
                <TataTertibView isAdmin={false} list={tataTertibList} />
              )}
              {activeTab === "pengaturan" && (
                <WargaPengaturan
                  userProfile={userProfile}
                  showToast={showToast}
                />
              )}
            </>
          )}
        </main>
      </div>

     {/* BOTTOM NAV MOBILE */}
      <nav className="lg:hidden fixed bottom-0 inset-x-0 bg-white border-t border-gray-200 flex justify-around py-2 z-30 overflow-x-auto">
        {menus.slice(0, 5).map((m) => {
          const Icon = m.icon;
          const active = activeTab === m.id;
          return (
            <button
              key={m.id}
              onClick={() => {
                setActiveTab(m.id);
                setMobileMenuOpen(false);
              }}
              className={`relative flex flex-col items-center px-2 py-1 text-[10px] font-medium shrink-0 ${
                active ? "text-orange-600" : "text-gray-500"
              }`}
            >
              <Icon className="w-5 h-5 mb-0.5" />
              <span className="truncate max-w-[64px]">{m.label}</span>
              {m.badge > 0 && (
                <span className="absolute -top-0.5 right-1 w-4 h-4 bg-red-500 text-white text-[9px] rounded-full flex items-center justify-center font-bold">
                  {m.badge}
                </span>
              )}
            </button>
          );
        })}
      </nav>
    </div>
  );
}

// ==========================================
// WARGA 1: BERANDA
// ==========================================
function WargaBeranda({
  userProfile,
  clockLabel,
  pembayaranList,
  setorKasList,
  setActiveTab,
}) {
  const myTrx = pembayaranList.filter((p) => p.warga_uid === userProfile.uid);
  const bulanLunasSet = new Set(
    myTrx
      .filter((p) => p.status === "Lunas")
      .flatMap((p) => p.bulan_dibayar || [])
  );
  const wajib = getBulanWajibTahunIni();
  const nunggakCount = wajib.filter((b) => !bulanLunasSet.has(b)).length;

  return (
    <div className="space-y-5">
      {nunggakCount > 0 && (
        <div className="bg-red-600 text-white px-4 py-3 rounded-xl flex items-center gap-3 text-sm font-medium shadow-sm">
          <AlertCircle className="w-5 h-5 shrink-0" />
          <span>
            Anda memiliki {nunggakCount} bulan tunggakan (hingga bulan ini). Mohon
            segera melakukan pembayaran iuran Paguyuban & BLOK 9 demi kelancaran
            operasional lingkungan kita.
          </span>
        </div>
      )}

      {/* Banner Oranye */}
      <div className="bg-gradient-to-r from-orange-500 to-orange-400 text-white rounded-2xl p-6 sm:p-8 shadow-sm flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <p className="text-orange-100 text-sm">Selamat datang,</p>
          <h3 className="text-2xl sm:text-3xl font-extrabold mt-1">
            {userProfile.nama}{" "}
            <span className="text-lg font-normal text-orange-100">
              ({userProfile.blok})
            </span>
          </h3>
          <p className="text-sm text-orange-50 mt-2 max-w-xl">
            Kelola pembayaran iuran lingkungan dan dapatkan informasi terbaru dari
            Pengurus Paguyuban BSR dan BLOK 9 BSR dengan mudah melalui portal ini.
          </p>
        </div>
        <div className="bg-white/20 backdrop-blur-sm px-4 py-2 rounded-full text-xs font-semibold flex items-center gap-2 self-start md:self-center">
          <Clock className="w-4 h-4" />
          <span>{clockLabel}</span>
        </div>
      </div>

      {/* 2 Kartu Bawah */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
        <div className="bg-white rounded-2xl border border-gray-200 p-6 flex flex-col justify-between">
          <div>
            <div className="flex items-center gap-2 text-gray-800 font-bold mb-4">
              <AlertCircle className="w-5 h-5 text-orange-500" />
              <span>Status Tagihan</span>
            </div>
            <p className="text-sm text-gray-500">Tunggakan (Hingga Bulan Ini)</p>
            <p className="text-4xl font-extrabold text-gray-900 mt-1">
              {nunggakCount}{" "}
              <span className="text-base font-normal text-gray-500">Bulan</span>
            </p>
          </div>
          <button
            onClick={() => setActiveTab("bayar")}
            className="mt-6 w-full py-3 bg-orange-50 hover:bg-orange-100 text-orange-600 font-bold rounded-xl text-sm transition"
          >
            Bayar Sekarang →
          </button>
        </div>

        <div className="bg-white rounded-2xl border border-gray-200 p-6">
          <div className="flex items-center gap-2 text-gray-800 font-bold mb-4">
            <Clock className="w-5 h-5 text-blue-600" />
            <span>Riwayat Pembayaran Anda</span>
          </div>
          {myTrx.length === 0 ? (
            <p className="text-sm text-gray-400 py-6 text-center">
              Belum ada riwayat pembayaran.
            </p>
          ) : (
            <div className="divide-y divide-gray-100 max-h-64 overflow-y-auto">
              {myTrx.slice(0, 6).map((t) => (
                <div
                  key={t.id}
                  className="py-3 flex items-center justify-between gap-2"
                >
                  <div>
                    <p className="font-bold text-sm text-gray-900">
                      {(t.bulan_dibayar || []).map(formatBulanLabel).join(", ")}
                    </p>
                    <p className="text-xs text-gray-400">
                      {formatTanggalIndo(t.waktu)} • {t.metode_pembayaran}
                    </p>
                  </div>
                  <div className="text-right">
                    <span
                      className={`px-2 py-0.5 rounded-md text-[10px] font-bold uppercase ${
                        t.status === "Lunas"
                          ? "bg-emerald-100 text-emerald-700"
                          : t.status === "Ditolak"
                          ? "bg-red-100 text-red-700"
                          : "bg-amber-100 text-amber-700"
                      }`}
                    >
                      {t.status}
                    </span>
                    <p className="text-sm font-bold text-gray-900 mt-0.5">
                      {formatRupiah(t.nominal)}
                    </p>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
{/* Riwayat Setoran ke Pengurus */}
      <div className="bg-white rounded-2xl border border-gray-200 p-6 space-y-4">
        <div className="flex items-start justify-between gap-3">
          <div>
            <div className="flex items-center gap-2 text-gray-900 font-bold">
              <ArrowUpCircle className="w-5 h-5 text-blue-600" />
              <span>Riwayat Setoran ke Pengurus</span>
            </div>
            <p className="text-xs text-gray-500 mt-1">
              Uang iuran yang telah disetorkan oleh RT ke Bendahara Utama
              Paguyuban / Blok.
            </p>
          </div>
          <span className="px-3 py-1 rounded-lg bg-gray-100 text-gray-600 text-xs font-semibold shrink-0">
            Transparansi
          </span>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-sm text-left">
            <thead className="bg-gray-50 text-xs text-gray-500 uppercase border-y border-gray-200">
              <tr>
                <th className="px-4 py-3">Tanggal</th>
                <th className="px-4 py-3">Jenis Kas</th>
                <th className="px-4 py-3">Keterangan</th>
                <th className="px-4 py-3 text-right">Nominal (Rp)</th>
                <th className="px-4 py-3 text-center">Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {setorKasList.filter((s) => s.tipe === "Pengeluaran").length ===
              0 ? (
                <tr>
                  <td
                    colSpan={5}
                    className="px-4 py-6 text-center text-gray-400 text-xs"
                  >
                    Belum ada riwayat setoran ke pengurus.
                  </td>
                </tr>
              ) : (
                setorKasList
                  .filter((s) => s.tipe === "Pengeluaran")
                  .map((s) => (
                    <tr key={s.id} className="hover:bg-gray-50">
                      <td className="px-4 py-3.5 text-gray-600 whitespace-nowrap">
                        {formatTanggalIndo(s.tanggal)}
                      </td>
                      <td className="px-4 py-3.5 font-bold text-gray-900 whitespace-nowrap">
                        {s.jenis_kas === "paguyuban"
                          ? "Kas Paguyuban BSR"
                          : "Kas BLOK 9"}
                      </td>
                      <td className="px-4 py-3.5 text-gray-600">
                        {s.keterangan}
                      </td>
                      <td className="px-4 py-3.5 text-right font-extrabold text-gray-900 whitespace-nowrap">
                        {Number(s.nominal || 0).toLocaleString("id-ID")}
                      </td>
                      <td className="px-4 py-3.5 text-center">
                        <span className="px-2.5 py-1 rounded-lg bg-blue-50 text-blue-600 border border-blue-200 text-[11px] font-bold">
                          DISETOR
                        </span>
                      </td>
                    </tr>
                  ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>

  );
}

// ==========================================
// WARGA 2: BAYAR TAGIHAN (FLOATING BAR + MODAL DETAIL)
// ==========================================
function WargaBayarTagihan({
  userProfile,
  pembayaranList,
  setPreviewImg,
  showToast,
}) {
  const tahunSekarang = new Date().getFullYear();
  const [tahun, setTahun] = useState(tahunSekarang);
  const [selectedMonths, setSelectedMonths] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [metode, setMetode] = useState("");
  const [buktiBase64, setBuktiBase64] = useState("");
  const [busy, setBusy] = useState(false);
  const [copied, setCopied] = useState("");

  const myTrx = pembayaranList.filter((p) => p.warga_uid === userProfile.uid);

  const getMonthStatus = (ym) => {
    for (const t of myTrx) {
      if ((t.bulan_dibayar || []).includes(ym)) {
        if (t.status === "Lunas") return "LUNAS";
        if (t.status === "Menunggu Verifikasi") return "PROSES";
      }
    }
    return null;
  };

  const wajib = getBulanWajibTahunIni();
  const nunggakCount = wajib.filter((b) => getMonthStatus(b) !== "LUNAS").length;

  const toggleMonth = (ym) => {
    if (getMonthStatus(ym)) return;
    setSelectedMonths((prev) =>
      prev.includes(ym) ? prev.filter((x) => x !== ym) : [...prev, ym].sort()
    );
  };

  const availableMonthsInYear = BULAN_LIST.map(
    (_, idx) => `${tahun}-${String(idx + 1).padStart(2, "0")}`
  ).filter((ym) => !getMonthStatus(ym));

  const isAllSelectedInYear =
    availableMonthsInYear.length > 0 &&
    availableMonthsInYear.every((ym) => selectedMonths.includes(ym));

  const handleToggleSelectAll = () => {
    if (isAllSelectedInYear) {
      setSelectedMonths((prev) =>
        prev.filter((ym) => !availableMonthsInYear.includes(ym))
      );
    } else {
      setSelectedMonths((prev) =>
        Array.from(new Set([...prev, ...availableMonthsInYear])).sort()
      );
    }
  };

  const jmlBulan = selectedMonths.length;
  const totalPaguyuban = jmlBulan * CONFIG.porsiPaguyuban;
  const totalBlok9 = jmlBulan * CONFIG.porsiBlok9;
  const totalBayar = jmlBulan * CONFIG.iuranPerBulan;

  const handleCopy = (val, key) => {
    navigator.clipboard.writeText(val);
    setCopied(key);
    setTimeout(() => setCopied(""), 2000);
  };

  const handleFile = async (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    const b64 = await compressImageToBase64(f);
    setBuktiBase64(b64);
  };

  const handleConfirm = async () => {
    if (!metode) {
      showToast("Pilih metode pembayaran terlebih dahulu.", "error");
      return;
    }
    if (metode !== "Tunai" && !buktiBase64) {
      showToast("Mohon unggah bukti transfer/QRIS terlebih dahulu.", "error");
      return;
    }
    setBusy(true);
    try {
      await addDoc(collection(db, "pembayaran_iuran"), {
        warga_uid: userProfile.uid,
        warga_nama: userProfile.nama,
        blok: userProfile.blok,
        waktu: new Date().toISOString(),
        bulan_dibayar: selectedMonths,
        jumlah_bulan: jmlBulan,
        kas_paguyuban: totalPaguyuban,
        kas_blok_9: totalBlok9,
        nominal: totalBayar,
        metode_pembayaran: metode,
        bukti_transfer_base64: metode === "Tunai" ? "" : buktiBase64,
        status: "Menunggu Verifikasi",
        catatan_admin: "",
      });
      setSelectedMonths([]);
      setMetode("");
      setBuktiBase64("");
      setShowModal(false);
      showToast("Konfirmasi pembayaran berhasil dikirim! Menunggu verifikasi Admin.");
    } catch (err) {
      showToast("Gagal mengirim: " + err.message, "error");
    } finally {
      setBusy(false);
    }
  };

  const metodeOptions = [
    { id: "Tunai", label: "Uang Tunai", icon: Wallet, color: "text-gray-600" },
    { id: "QRIS", label: "QRIS", icon: QrCode, color: "text-gray-700" },
    { id: "BCA", label: "Trf. BCA", icon: CreditCard, color: "text-blue-600" },
    { id: "BRI", label: "Trf. BRI", icon: CreditCard, color: "text-blue-800" },
  ];

  return (
    <div className="space-y-5 pb-28">
      {nunggakCount > 0 && (
        <div className="bg-red-600 text-white p-5 rounded-2xl shadow-sm">
          <div className="flex items-center gap-2 font-bold">
            <AlertCircle className="w-5 h-5" /> Pengingat Tunggakan
          </div>
          <p className="text-sm text-red-50 mt-1">
            Anda memiliki <b>{nunggakCount} bulan</b> tunggakan iuran (terhitung
            hingga bulan ini). Silakan klik bulan yang belum dilunasi pada daftar
            di bawah ini untuk melakukan pembayaran.
          </p>
        </div>
      )}

      {/* Pilih Bulan Tagihan */}
      <div className="bg-white rounded-2xl border border-gray-200 p-6 space-y-5">
        <div>
          <h3 className="text-lg font-bold text-gray-900">Pilih Bulan Tagihan</h3>
          <p className="text-sm text-gray-500">
            Klik pada bulan yang ingin dibayarkan. Anda dapat memilih lebih dari
            satu bulan lintas tahun.
          </p>
        </div>

        <div className="bg-gray-50 p-3.5 rounded-xl flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <span className="text-sm font-semibold text-gray-700">
              Tahun Tagihan:
            </span>
            <select
              value={tahun}
              onChange={(e) => setTahun(Number(e.target.value))}
              className="px-3 py-1.5 bg-white border border-gray-300 rounded-lg text-sm font-semibold"
            >
              {[
                tahunSekarang,
                tahunSekarang + 1,
                tahunSekarang + 2,
                tahunSekarang + 3,
              ].map((y) => (
                <option key={y} value={y}>
                  Tahun {y}
                </option>
              ))}
            </select>
          </div>

          {availableMonthsInYear.length > 0 && (
            <button
              type="button"
              onClick={handleToggleSelectAll}
              className={`px-3.5 py-1.5 rounded-lg text-xs font-bold border transition ${
                isAllSelectedInYear
                  ? "bg-orange-500 text-white border-orange-500"
                  : "bg-white text-orange-600 border-orange-300 hover:bg-orange-50"
              }`}
            >
              {isAllSelectedInYear
                ? `Batal Pilih Semua (${tahun})`
                : `Pilih Semua Bulan (${availableMonthsInYear.length})`}
            </button>
          )}
        </div>

        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          {BULAN_LIST.map((nama, idx) => {
            const ym = `${tahun}-${String(idx + 1).padStart(2, "0")}`;
            const st = getMonthStatus(ym);
            const selected = selectedMonths.includes(ym);
            const locked = Boolean(st);

            return (
              <button
                key={ym}
                type="button"
                disabled={locked}
                onClick={() => toggleMonth(ym)}
                className={`relative p-4 rounded-xl border text-left transition ${
                  st === "LUNAS"
                    ? "bg-gray-50 border-gray-200 cursor-not-allowed opacity-75"
                    : st === "PROSES"
                    ? "bg-amber-50 border-amber-200 cursor-not-allowed"
                    : selected
                    ? "bg-orange-50/70 border-gray-900 ring-1 ring-gray-900"
                    : "bg-white border-gray-200 hover:border-orange-400"
                }`}
              >
                <div className="flex items-center justify-between">
                  <p
                    className={`font-bold text-sm ${
                      locked
                        ? "text-gray-400"
                        : selected
                        ? "text-orange-600"
                        : "text-gray-900"
                    }`}
                  >
                    {nama}
                  </p>
                  {selected && (
                    <CheckCircle2 className="w-4 h-4 text-orange-500" />
                  )}
                </div>

                {st === "LUNAS" ? (
                  <span className="inline-block mt-2 px-2 py-0.5 rounded bg-emerald-100 text-emerald-700 text-[10px] font-bold">
                    LUNAS
                  </span>
                ) : st === "PROSES" ? (
                  <span className="inline-block mt-2 px-2 py-0.5 rounded bg-amber-100 text-amber-700 text-[10px] font-bold">
                    VERIFIKASI
                  </span>
                ) : (
                  <p
                    className={`text-xs mt-1.5 ${
                      selected ? "text-orange-600 font-semibold" : "text-gray-500"
                    }`}
                  >
                    {formatRupiah(CONFIG.iuranPerBulan)}
                  </p>
                )}
              </button>
            );
          })}
        </div>
      </div>

      {/* FLOATING BAR HITAM (MUNCUL SAAT ADA BULAN DIPILIH) */}
      {jmlBulan > 0 && (
        <div className="fixed bottom-16 md:bottom-6 left-4 right-4 md:left-72 md:right-8 z-30 max-w-5xl mx-auto">
          <div className="bg-slate-900 text-white rounded-2xl px-6 py-4 shadow-2xl flex items-center justify-between gap-4">
            <div>
              <p className="text-[11px] font-semibold tracking-wider text-slate-400 uppercase">
                {jmlBulan} BULAN TERPILIH
              </p>
              <p className="text-2xl font-extrabold mt-0.5">
                {formatRupiah(totalBayar)}
              </p>
            </div>
            <button
              type="button"
              onClick={() => setShowModal(true)}
              className="px-6 py-3 bg-orange-500 hover:bg-orange-600 text-white font-bold rounded-xl text-sm shadow-md transition"
            >
              Lanjut Bayar &gt;
            </button>
          </div>
        </div>
      )}

      {/* MODAL POPUP DETAIL PEMBAYARAN */}
      {showModal && (
        <div
          className="fixed inset-0 z-50 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4"
          onClick={() => setShowModal(false)}
        >
          <div
            className="bg-white rounded-3xl max-w-md w-full max-h-[90vh] overflow-y-auto shadow-2xl"
            onClick={(e) => e.stopPropagation()}
          >
            {/* Header Modal */}
            <div className="bg-orange-50/70 px-6 py-4 flex items-center justify-between border-b border-orange-100 sticky top-0">
              <h3 className="font-bold text-gray-900 text-lg">
                Detail Pembayaran
              </h3>
              <button
                type="button"
                onClick={() => setShowModal(false)}
                className="w-8 h-8 rounded-full bg-white text-gray-500 hover:text-gray-800 flex items-center justify-center font-bold shadow-xs"
              >
                ✕
              </button>
            </div>

            <div className="p-6 space-y-5">
              {/* Box Rincian Peruntukan */}
              <div className="border border-gray-200 rounded-2xl p-4 space-y-2.5 text-sm">
                <p className="font-bold text-gray-800">
                  Rincian peruntukan ({jmlBulan} bulan)
                </p>
                <div className="flex justify-between text-gray-600">
                  <span>
                    Kas Paguyuban BSR{" "}
                    <span className="text-gray-400">(20k x {jmlBulan})</span>
                  </span>
                  <span className="font-semibold text-gray-900">
                    {formatRupiah(totalPaguyuban)}
                  </span>
                </div>
                <div className="flex justify-between text-gray-600">
                  <span>
                    Kas BLOK 9{" "}
                    <span className="text-gray-400">(5k x {jmlBulan})</span>
                  </span>
                  <span className="font-semibold text-gray-900">
                    {formatRupiah(totalBlok9)}
                  </span>
                </div>
                <div className="pt-2.5 border-t border-gray-100 flex justify-between items-center">
                  <span className="font-bold text-gray-900">Total Tagihan</span>
                  <span className="text-lg font-extrabold text-orange-600">
                    {formatRupiah(totalBayar)}
                  </span>
                </div>
              </div>

              {/* Pilih Metode Pembayaran (Grid 2x2) */}
              <div>
                <p className="text-sm font-bold text-gray-800 mb-3">
                  Pilih Metode Pembayaran
                </p>
                <div className="grid grid-cols-2 gap-3">
                  {metodeOptions.map((m) => {
                    const Icon = m.icon;
                    const active = metode === m.id;
                    return (
                      <button
                        key={m.id}
                        type="button"
                        onClick={() => setMetode(m.id)}
                        className={`p-4 rounded-2xl border flex flex-col items-center justify-center gap-1.5 transition ${
                          active
                            ? "border-orange-500 bg-orange-50 ring-2 ring-orange-200"
                            : "border-gray-200 bg-white hover:border-gray-300"
                        }`}
                      >
                        <Icon className={`w-5 h-5 ${m.color}`} />
                        <span className="text-xs font-semibold text-gray-800">
                          {m.label}
                        </span>
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* Info Detail Sesuai Metode Dipilih */}
              {metode === "BCA" && (
                <div className="p-4 bg-blue-50 rounded-2xl text-sm flex items-center justify-between">
                  <div>
                    <p className="text-xs font-bold text-blue-700">{CONFIG.bca.bank}</p>
                    <p className="font-mono font-bold text-gray-900">{CONFIG.bca.norek}</p>
                    <p className="text-xs text-gray-500">{CONFIG.bca.an}</p>
                  </div>
                  <button
                    type="button"
                    onClick={() => handleCopy(CONFIG.bca.norek, "bca")}
                    className="px-3 py-1.5 bg-white rounded-lg text-xs font-bold text-blue-700 shadow-xs"
                  >
                    {copied === "bca" ? "Disalin" : "Salin"}
                  </button>
                </div>
              )}

              {metode === "BRI" && (
                <div className="p-4 bg-blue-50 rounded-2xl text-sm flex items-center justify-between">
                  <div>
                    <p className="text-xs font-bold text-blue-900">{CONFIG.bri.bank}</p>
                    <p className="font-mono font-bold text-gray-900">{CONFIG.bri.norek}</p>
                    <p className="text-xs text-gray-500">{CONFIG.bri.an}</p>
                  </div>
                  <button
                    type="button"
                    onClick={() => handleCopy(CONFIG.bri.norek, "bri")}
                    className="px-3 py-1.5 bg-white rounded-lg text-xs font-bold text-blue-900 shadow-xs"
                  >
                    {copied === "bri" ? "Disalin" : "Salin"}
                  </button>
                </div>
              )}

              {metode === "QRIS" && (
                <div className="p-4 bg-gray-50 rounded-2xl text-center space-y-2">
                  <img
                    src={CONFIG.qrisUrl}
                    alt="QRIS"
                    onClick={() => setPreviewImg(CONFIG.qrisUrl)}
                    className="w-48 mx-auto rounded-xl border border-gray-200 cursor-pointer"
                  />
                  <p className="text-xs text-gray-500">
                    Klik gambar QRIS untuk memperbesar
                  </p>
                </div>
              )}

              {metode === "Tunai" && (
                <div className="p-4 bg-emerald-50 rounded-2xl text-xs text-emerald-800">
                  Silakan serahkan uang tunai sebesar <b>{formatRupiah(totalBayar)}</b>{" "}
                  kepada Bendahara Blok 9, lalu klik tombol konfirmasi di bawah ini.
                </div>
              )}

              {/* Upload Bukti (untuk non-tunai) */}
              {metode && metode !== "Tunai" && (
                <div>
                  <label className="block text-xs font-bold text-gray-700 mb-1.5">
                    Upload Bukti Transfer / QRIS
                  </label>
                  <input
                    type="file"
                    accept="image/*"
                    onChange={handleFile}
                    className="w-full text-xs file:mr-3 file:py-2 file:px-3 file:rounded-xl file:border-0 file:bg-orange-50 file:text-orange-700 file:font-bold"
                  />
                  {buktiBase64 && (
                    <img
                      src={buktiBase64}
                      alt="Bukti"
                      className="mt-2 h-32 rounded-xl border border-gray-200 object-contain"
                    />
                  )}
                </div>
              )}

              {/* Tombol Konfirmasi */}
              <button
                type="button"
                disabled={busy || !metode || (metode !== "Tunai" && !buktiBase64)}
                onClick={handleConfirm}
                className="w-full py-3.5 rounded-2xl font-bold text-sm transition bg-orange-500 hover:bg-orange-600 text-white disabled:bg-gray-200 disabled:text-gray-400"
              >
                {busy ? "Mengirim..." : "Konfirmasi Pembayaran"}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
// ==========================================
// WARGA 5: PENGATURAN AKUN (GANTI PASSWORD)
// ==========================================
function WargaPengaturan({ userProfile, showToast }) {
  const [oldPass, setOldPass] = useState("");
  const [newPass, setNewPass] = useState("");
  const [confirmPass, setConfirmPass] = useState("");
  const [busy, setBusy] = useState(false);

  const handleChangePass = async (e) => {
    e.preventDefault();
    if (newPass !== confirmPass) {
      showToast("Konfirmasi password baru tidak cocok.", "error");
      return;
    }
    setBusy(true);
    try {
      const cred = EmailAuthProvider.credential(
        auth.currentUser.email,
        toFirebasePassword(oldPass)
      );
      await reauthenticateWithCredential(auth.currentUser, cred);
      await updatePassword(auth.currentUser, toFirebasePassword(newPass));
      setOldPass("");
      setNewPass("");
      setConfirmPass("");
      showToast("Password berhasil diperbarui!");
    } catch {
      showToast("Password lama salah atau gagal memperbarui.", "error");
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="bg-white rounded-2xl border border-gray-200 p-6 sm:p-8 max-w-2xl space-y-6">
      <div className="flex items-center gap-3 pb-4 border-b border-gray-100">
        <div className="w-11 h-11 bg-blue-50 text-blue-600 rounded-xl flex items-center justify-center">
          <Settings className="w-6 h-6" />
        </div>
        <div>
          <h3 className="text-lg font-bold text-gray-900">Pengaturan Akun</h3>
          <p className="text-xs text-gray-500">Ganti sandi dan kelola profil Anda.</p>
        </div>
      </div>

      <div>
        <h4 className="text-sm font-bold text-gray-800 mb-2">Informasi Profil</h4>
        <div className="bg-gray-50 rounded-xl p-4 grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div>
            <p className="text-[11px] font-semibold text-gray-400 uppercase">
              Nama Lengkap
            </p>
            <p className="font-bold text-gray-900 mt-0.5">{userProfile.nama}</p>
          </div>
          <div>
            <p className="text-[11px] font-semibold text-gray-400 uppercase">
              ID Warga / Username
            </p>
            <p className="font-bold text-orange-600 mt-0.5">
              {userProfile.username}
            </p>
          </div>
        </div>
      </div>

      <form onSubmit={handleChangePass} className="space-y-4">
        <h4 className="text-sm font-bold text-gray-800">Ganti Password</h4>
        <div>
          <label className="block text-xs font-semibold text-gray-600 mb-1">
            Password Lama
          </label>
          <input
            type="password"
            required
            value={oldPass}
            onChange={(e) => setOldPass(e.target.value)}
            className="w-full px-3.5 py-2.5 rounded-xl border border-gray-300 text-sm"
          />
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div>
            <label className="block text-xs font-semibold text-gray-600 mb-1">
              Password Baru
            </label>
            <input
              type="password"
              required
              value={newPass}
              onChange={(e) => setNewPass(e.target.value)}
              className="w-full px-3.5 py-2.5 rounded-xl border border-gray-300 text-sm"
            />
          </div>
          <div>
            <label className="block text-xs font-semibold text-gray-600 mb-1">
              Ulangi Password Baru
            </label>
            <input
              type="password"
              required
              value={confirmPass}
              onChange={(e) => setConfirmPass(e.target.value)}
              className="w-full px-3.5 py-2.5 rounded-xl border border-gray-300 text-sm"
            />
          </div>
        </div>
        <button
          type="submit"
          disabled={busy}
          className="inline-flex items-center gap-2 px-5 py-2.5 bg-orange-500 hover:bg-orange-600 text-white rounded-xl text-sm font-bold"
        >
          <Lock className="w-4 h-4" />
          {busy ? "Menyimpan..." : "Simpan Password"}
        </button>
      </form>
    </div>
  );
}

// ==========================================
// ADMIN 1: DASHBOARD
// ==========================================
function AdminDashboard({
  clockLabel,
  saldoSummary,
  wargaOnly,
  pembayaranList,
  setActiveTab,
}) {
  const wargaAktif = wargaOnly.filter((w) => w.status !== "Tidak Aktif").length;
  const recentLunas = pembayaranList.filter((p) => p.status === "Lunas").slice(0, 5);

  return (
    <div className="space-y-6">
      <div className="inline-flex items-center gap-2 bg-white border border-gray-200 px-4 py-2 rounded-full text-xs font-semibold text-gray-700 shadow-xs">
        <Clock className="w-4 h-4 text-orange-500" />
        <span>{clockLabel}</span>
      </div>

      {/* 3 Kartu Ringkasan */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <div className="bg-white p-5 rounded-2xl border-2 border-blue-100 flex items-center justify-between">
          <div>
            <p className="text-xs font-semibold text-blue-600">Kas Paguyuban BSR</p>
            <p className="text-2xl font-extrabold text-gray-900 mt-1">
              {formatRupiah(saldoSummary.saldoPaguyuban)}
            </p>
          </div>
          <div className="w-11 h-11 rounded-xl bg-blue-50 text-blue-600 flex items-center justify-center">
            <Wallet className="w-5 h-5" />
          </div>
        </div>

        <div className="bg-white p-5 rounded-2xl border border-gray-200 flex items-center justify-between">
          <div>
            <p className="text-xs font-semibold text-gray-500">Kas Khusus BLOK 9</p>
            <p className="text-2xl font-extrabold text-orange-600 mt-1">
              {formatRupiah(saldoSummary.saldoBlok9)}
            </p>
          </div>
          <div className="w-11 h-11 rounded-xl bg-orange-50 text-orange-500 flex items-center justify-center">
            <Wallet className="w-5 h-5" />
          </div>
        </div>

        <div className="bg-white p-5 rounded-2xl border border-gray-200 flex items-center justify-between">
          <div>
            <p className="text-xs font-semibold text-gray-500">Warga Aktif</p>
            <p className="text-2xl font-extrabold text-emerald-600 mt-1">
              {wargaAktif}{" "}
              <span className="text-sm font-normal text-gray-400">KK</span>
            </p>
          </div>
          <div className="w-11 h-11 rounded-xl bg-emerald-50 text-emerald-600 flex items-center justify-center">
            <Users className="w-5 h-5" />
          </div>
        </div>
      </div>

      {/* Transaksi Terakhir */}
      <div className="bg-white rounded-2xl border border-gray-200 p-6">
        <div className="flex items-center justify-between mb-4">
          <div className="flex items-center gap-2 font-bold text-gray-900">
            <Activity className="w-5 h-5 text-orange-500" />
            <span>Transaksi Terakhir</span>
          </div>
          <button
            onClick={() => setActiveTab("laporan")}
            className="text-xs font-bold text-orange-600 hover:underline"
          >
            Lihat Semua
          </button>
        </div>
        {recentLunas.length === 0 ? (
          <p className="text-sm text-gray-400 py-6 text-center">
            Belum ada transaksi lunas.
          </p>
        ) : (
          <div className="divide-y divide-gray-100">
            {recentLunas.map((t) => (
              <div
                key={t.id}
                className="py-3.5 flex items-center justify-between gap-2"
              >
                <div>
                  <p className="font-bold text-sm text-gray-900">
                    {t.warga_nama}{" "}
                    <span className="font-normal text-gray-400">({t.blok})</span>
                  </p>
                  <p className="text-xs text-gray-500">
                    Membayar bulan:{" "}
                    {(t.bulan_dibayar || []).map(formatBulanLabel).join(", ")}
                  </p>
                </div>
                <div className="text-right">
                  <p className="font-bold text-sm text-emerald-600">
                    + {formatRupiah(t.nominal)}
                  </p>
                  <p className="text-[11px] text-gray-400">
                    {formatTanggalIndo(t.waktu)} via {t.metode_pembayaran}
                  </p>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// ==========================================
// ADMIN 2: VERIFIKASI PEMBAYARAN
// ==========================================
function AdminVerifikasi({ pembayaranList, setPreviewImg, showToast }) {
  const pending = pembayaranList.filter((p) => p.status === "Menunggu Verifikasi");

  const handleAction = async (item, status) => {
    let catatan = "";
    if (status === "Ditolak") {
      catatan = prompt("Alasan penolakan:");
      if (catatan === null) return;
    }
    await updateDoc(doc(db, "pembayaran_iuran", item.id), {
      status,
      catatan_admin: catatan,
    });
    showToast(`Pembayaran ${item.warga_nama} diubah menjadi ${status}`);

    // Jika disetujui (Lunas): otomatis unduh kwitansi PNG + buka WA warga
    if (status === "Lunas") {
      const nama = item.warga_nama || item.nama || "-";
      const blok = item.warga_blok || item.blok || "-";
      const hpRaw = item.warga_hp || item.hp || "";
      const bulanArr = item.bulan_dibayar || [];
      const nominal = Number(item.nominal || item.total || bulanArr.length * 25000);

      // 1. Unduh gambar kwitansi PNG otomatis
      downloadKwitansiPNG({
        ...item,
        nama,
        blok,
        hp: hpRaw || "-",
        nominal,
      });

      // 2. Siapkan pesan WA resmi
      const tgl = new Date().toLocaleDateString("id-ID", {
        day: "numeric",
        month: "long",
        year: "numeric",
      });
      const pesan = encodeURIComponent(
        `Assalamu'alaikum Bapak/Ibu *${nama}* (${blok}),\n\n` +
        `Pembayaran iuran kas BLOK 9 BSR telah kami terima & diverifikasi *LUNAS* ✅\n\n` +
        `• *Bulan:* ${bulanArr.join(", ") || "-"}\n` +
        `• *Total:* Rp ${nominal.toLocaleString("id-ID")}\n` +
        `• *Tanggal:* ${tgl}\n\n` +
        `Berikut kami lampirkan bukti kwitansinya. Terima kasih 🙏`
      );

      // Format nomor HP (08xxx -> 628xxx)
      let waNum = String(hpRaw).replace(/[^0-9]/g, "");
      if (waNum.startsWith("0")) waNum = "62" + waNum.slice(1);

      const waUrl = waNum
        ? `https://wa.me/${waNum}?text=${pesan}`
        : `https://wa.me/?text=${pesan}`;

      window.open(waUrl, "_blank");
    }
  };

  return (
    <div className="space-y-4">
      <div>
        <h3 className="text-xl font-bold text-gray-900">Verifikasi Pembayaran</h3>
        <p className="text-sm text-gray-500">
          Cek dan konfirmasi bukti transfer warga sebelum masuk ke laporan.
        </p>
      </div>

      <div className="bg-white rounded-2xl border border-gray-200 overflow-x-auto">
        <table className="w-full text-sm text-left">
          <thead className="bg-gray-50 text-xs text-gray-500 uppercase border-b border-gray-200">
            <tr>
              <th className="px-4 py-3">Tanggal</th>
              <th className="px-4 py-3">Warga (Blok)</th>
              <th className="px-4 py-3">Detail Tagihan</th>
              <th className="px-4 py-3">Metode</th>
              <th className="px-4 py-3">Bukti</th>
              <th className="px-4 py-3 text-right">Aksi</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {pending.length === 0 ? (
              <tr>
                <td colSpan={6} className="px-4 py-8 text-center text-gray-400">
                  Tidak ada bukti transfer yang menunggu verifikasi.
                </td>
              </tr>
            ) : (
              pending.map((p) => (
                <tr key={p.id} className="hover:bg-gray-50">
                  <td className="px-4 py-3.5 text-gray-500 whitespace-nowrap">
                    {formatTanggalIndo(p.waktu)}
                  </td>
                  <td className="px-4 py-3.5">
                    <span className="font-bold text-gray-900">{p.warga_nama}</span>{" "}
                    <span className="text-gray-400">({p.blok})</span>
                  </td>
                  <td className="px-4 py-3.5">
                    <div className="text-xs text-gray-500">
                      {(p.bulan_dibayar || []).map(formatBulanLabel).join(", ")}
                    </div>
                    <div className="font-bold text-gray-900">
                      {formatRupiah(p.nominal)}
                    </div>
                  </td>
                  <td className="px-4 py-3.5 text-gray-600">{p.metode_pembayaran}</td>
                  <td className="px-4 py-3.5">
                    <button
                      onClick={() => setPreviewImg(p.bukti_transfer_base64)}
                      className="px-3 py-1 bg-blue-50 text-blue-600 rounded-lg text-xs font-semibold hover:bg-blue-100"
                    >
                      Lihat Bukti
                    </button>
                  </td>
                  <td className="px-4 py-3.5 text-right whitespace-nowrap space-x-2">
                    <button
                      onClick={() => handleAction(p, "Lunas")}
                      className="inline-flex items-center gap-1 px-3 py-1.5 bg-emerald-500 hover:bg-emerald-600 text-white rounded-lg text-xs font-bold"
                    >
                      <CheckCircle2 className="w-3.5 h-3.5" /> Terima
                    </button>
                    <button
                      onClick={() => handleAction(p, "Ditolak")}
                      className="inline-flex items-center gap-1 px-3 py-1.5 bg-red-50 hover:bg-red-100 text-red-600 rounded-lg text-xs font-bold"
                    >
                      <XCircle className="w-3.5 h-3.5" /> Tolak
                    </button>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ==========================================
// ADMIN 3: LAPORAN TAGIHAN (IMPORT + EXPORT EXCEL & PDF)
// ==========================================
function AdminLaporanTagihan({ pembayaranList, wargaOnly, showToast }) {
  const tahunSekarang = new Date().getFullYear();
  const [tahun, setTahun] = useState(tahunSekarang);
  const [importing, setImporting] = useState(false);
  const [searchLaporan, setSearchLaporan] = useState("");

  const fmtRp = (n) => "Rp " + Number(n || 0).toLocaleString("id-ID");

  const notif = (msg, type = "success") => {
    if (showToast) showToast(msg, type);
    else alert(msg);
  };

  const getStatus = (uid, ym) => {
    const trx = pembayaranList.filter((p) => p.warga_uid === uid);
    for (const t of trx) {
      if ((t.bulan_dibayar || []).includes(ym)) {
        if (t.status === "Lunas") return "LUNAS";
        if (t.status === "Menunggu Verifikasi") return "PROSES";
      }
    }
    return "-";
  };

  // Import Laporan Tagihan dari Excel (.xlsx / .xls / .csv)
  const handleImportExcel = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = "";

    if (
      !confirm(
        `Import data tagihan lunas untuk Tahun ${tahun} dari file "${file.name}"?`
      )
    ) {
      return;
    }

    setImporting(true);
    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      let berhasilWarga = 0;
      let totalBulanMasuk = 0;
      let dilewati = 0;

      const isLunasCell = (val) => {
        const s = String(val ?? "").trim().toLowerCase();
        if (!s || s === "-" || s === "0" || s === "belum" || s === "proses" || s === "tidak") {
          return false;
        }
        return true;
      };

      for (const row of rows) {
        const nama = String(
          row["Nama Lengkap"] || row["Nama"] || row["Warga"] || row["nama"] || ""
        ).trim();
        const blok = String(
          row["Blok/No."] || row["Blok/No"] || row["Blok"] || row["blok"] || ""
        ).trim();

        if (!nama && !blok) continue;

        const numMatch = blok.match(/(\d+)\s*$/);
        const nomorRumah = numMatch ? parseInt(numMatch[1], 10) : null;

        // Cocokkan warga berdasarkan Blok, Nomor Rumah, atau Nama
        const warga = wargaOnly.find(
          (w) =>
            (blok &&
              (w.blok || "").toLowerCase().replace(/\s+/g, "") ===
                blok.toLowerCase().replace(/\s+/g, "")) ||
            (nomorRumah && w.nomor_rumah === nomorRumah) ||
            (nama && (w.nama || "").toLowerCase() === nama.toLowerCase())
        );

        if (!warga) {
          dilewati++;
          continue;
        }

        const uid = warga.uid || warga.id;
        const bulanBaru = [];

        BULAN_LIST.forEach((namaBulan, idx) => {
          const short3 = namaBulan.slice(0, 3); // Jan, Feb, Mar, ...
          // Cari nilai sel berdasarkan nama bulan pendek/lengkap (case-insensitive)
          const matchedKey = Object.keys(row).find((k) => {
            const kl = k.trim().toLowerCase();
            return (
              kl === short3.toLowerCase() ||
              kl === namaBulan.toLowerCase() ||
              (idx === 7 && (kl === "agt" || kl === "aug")) ||
              (idx === 9 && kl === "oct") ||
              (idx === 11 && kl === "dec") ||
              (idx === 4 && kl === "may")
            );
          });

          if (matchedKey && isLunasCell(row[matchedKey])) {
            const ym = `${tahun}-${String(idx + 1).padStart(2, "0")}`;
            // Hanya masukkan jika belum LUNAS di database (anti-duplikat)
            if (getStatus(uid, ym) !== "LUNAS") {
              bulanBaru.push(ym);
            }
          }
        });

        if (bulanBaru.length > 0) {
          await addDoc(collection(db, "pembayaran_iuran"), {
            warga_uid: uid,
            warga_nama: warga.nama,
            warga_blok: warga.blok,
            bulan_dibayar: bulanBaru,
            jumlah_bulan: bulanBaru.length,
            nominal_paguyuban: bulanBaru.length * CONFIG.porsiPaguyuban,
            nominal_blok9: bulanBaru.length * CONFIG.porsiBlok9,
            total_bayar: bulanBaru.length * CONFIG.iuranPerBulan,
            metode: "Import Excel",
            bukti_url: "",
            status: "Lunas",
            createdAt: serverTimestamp(),
            verifiedAt: serverTimestamp(),
          });
          berhasilWarga++;
          totalBulanMasuk += bulanBaru.length;
        }
      }

      notif(
        `Import Laporan ${tahun} selesai! ${berhasilWarga} warga diperbarui (${totalBulanMasuk} bulan lunas). Dilewati: ${dilewati}.`
      );
    } catch (err) {
      notif("Gagal import Excel: " + err.message, "error");
    } finally {
      setImporting(false);
    }
  };

  // Export Laporan Tagihan ke Excel (.xlsx)
  const handleExportExcel = () => {
    const rows = wargaOnly.map((w, idx) => {
      const rowData = {
        No: idx + 1,
        "Nama Lengkap": w.nama || "-",
        "Blok/No.": w.blok || "-",
      };
      let lunasCount = 0;
      BULAN_LIST.forEach((b, i) => {
        const ym = `${tahun}-${String(i + 1).padStart(2, "0")}`;
        const st = getStatus(w.uid || w.id, ym);
        if (st === "LUNAS") lunasCount++;
        rowData[b.slice(0, 3)] =
          st === "LUNAS" ? "Lunas" : st === "PROSES" ? "Proses" : "-";
      });
      rowData["Total Lunas"] = `${lunasCount} Bln`;
      rowData["Total Nominal"] = lunasCount * CONFIG.iuranPerBulan;
      return rowData;
    });

    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [
      { wch: 5 },
      { wch: 24 },
      { wch: 13 },
      ...Array(12).fill({ wch: 8 }),
      { wch: 11 },
      { wch: 15 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `Laporan ${tahun}`);
    XLSX.writeFile(wb, `Laporan_Tagihan_Blok9_BSR_${tahun}.xlsx`);
  };

  // Export Laporan Tagihan ke PDF (A4 Landscape)
  const handleExportPDF = () => {
    const tgl = new Date().toLocaleDateString("id-ID", {
      day: "2-digit",
      month: "long",
      year: "numeric",
    });

    const rowsHtml = wargaOnly
      .map((w, idx) => {
        let lunasCount = 0;
        const bulanCells = BULAN_LIST.map((_, i) => {
          const ym = `${tahun}-${String(i + 1).padStart(2, "0")}`;
          const st = getStatus(w.uid || w.id, ym);
          if (st === "LUNAS") lunasCount++;
          const label = st === "LUNAS" ? "✓" : st === "PROSES" ? "P" : "-";
          const color =
            st === "LUNAS" ? "#059669" : st === "PROSES" ? "#d97706" : "#9ca3af";
          return `<td style="text-align:center;font-weight:bold;color:${color}">${label}</td>`;
        }).join("");

        return `
          <tr>
            <td style="text-align:center">${idx + 1}</td>
            <td><b>${w.nama || "-"}</b></td>
            <td>${w.blok || "-"}</td>
            ${bulanCells}
            <td style="text-align:right;font-weight:bold">${fmtRp(
              lunasCount * CONFIG.iuranPerBulan
            )}</td>
          </tr>`;
      })
      .join("");

    const win = window.open("", "_blank");
    if (!win) {
      alert("Pop-up diblokir browser. Izinkan pop-up untuk export PDF.");
      return;
    }
    win.document.write(`
      <html>
        <head>
          <title>Laporan Tagihan BLOK 9 BSR - ${tahun}</title>
          <style>
            @page { size: A4 landscape; margin: 10mm; }
            body { font-family: Arial, sans-serif; color: #111827; font-size: 10px; }
            h2 { margin: 0 0 4px 0; font-size: 15px; }
            .meta { color: #6b7280; margin-bottom: 10px; font-size: 10px; }
            table { width: 100%; border-collapse: collapse; }
            th, td { border: 1px solid #d1d5db; padding: 4px 5px; text-align: left; }
            th { background: #f3f4f6; font-size: 9px; text-transform: uppercase; }
            tr:nth-child(even) { background: #f9fafb; }
          </style>
        </head>
        <body>
          <h2>LAPORAN TAGIHAN IURAN WARGA BLOK 9 BSR - TAHUN ${tahun}</h2>
          <div class="meta">Dicetak pada: ${tgl} &bull; Keterangan: ✓ = Lunas, P = Menunggu Verifikasi, - = Belum Bayar</div>
          <table>
            <thead>
              <tr>
                <th style="text-align:center">No</th>
                <th>Nama Lengkap</th>
                <th>Blok/No</th>
                ${BULAN_LIST.map(
                  (b) => `<th style="text-align:center">${b.slice(0, 3)}</th>`
                ).join("")}
                <th style="text-align:right">Total</th>
              </tr>
            </thead>
            <tbody>${rowsHtml}</tbody>
          </table>
          <script>window.onload = () => { window.print(); };</script>
        </body>
      </html>
    `);
    win.document.close();
  };

  return (
    <div className="space-y-4">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h3 className="text-xl font-bold text-gray-900">Laporan Tagihan</h3>
          <p className="text-sm text-gray-500">
            Rekap status pembayaran iuran warga per bulan.
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <div className="relative w-full sm:w-56">
            <Search className="w-4 h-4 text-gray-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchLaporan}
              onChange={(e) => setSearchLaporan(e.target.value)}
              placeholder="Cari nama / no rumah..."
              className="w-full pl-9 pr-3 py-2 rounded-xl border border-gray-300 text-xs focus:ring-2 focus:ring-orange-500 outline-none bg-white"
            />
          </div>
          <select
            value={tahun}
            onChange={(e) => setTahun(Number(e.target.value))}
            className="px-3 py-2 bg-white border border-gray-300 rounded-xl text-xs font-bold"
          >
            {[
              tahunSekarang - 1,
              tahunSekarang,
              tahunSekarang + 1,
              tahunSekarang + 2,
              tahunSekarang + 3,
            ].map((y) => (
              <option key={y} value={y}>
                Tahun {y}
              </option>
            ))}
          </select>

          <label
            className={`inline-flex items-center gap-1.5 px-3.5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold shadow-sm cursor-pointer ${
              importing ? "opacity-60 pointer-events-none" : ""
            }`}
          >
            <ArrowUpCircle className="w-4 h-4" />
            <span>{importing ? "Mengimpor..." : "Import Excel"}</span>
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleImportExcel}
              className="hidden"
            />
          </label>

          <button
            type="button"
            onClick={handleExportExcel}
            className="inline-flex items-center gap-1.5 px-3.5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold shadow-xs"
          >
            <Download className="w-4 h-4" /> Export Excel
          </button>

          <button
            type="button"
            onClick={handleExportPDF}
            className="inline-flex items-center gap-1.5 px-3.5 py-2.5 bg-red-600 hover:bg-red-700 text-white rounded-xl text-xs font-bold shadow-xs"
          >
            <Download className="w-4 h-4" /> Export PDF
          </button>
        </div>
      </div>

      <div className="bg-white rounded-2xl border border-gray-200 overflow-x-auto">
        <table className="w-full text-xs text-left">
          <thead className="bg-gray-50 text-gray-500 uppercase border-b border-gray-200">
            <tr>
              <th className="px-3 py-3 sticky left-0 bg-gray-50 z-10">Warga</th>
              {BULAN_LIST.map((b) => (
                <th key={b} className="px-2 py-3 text-center">
                  {b.slice(0, 3)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {wargaOnly
              .filter((w) => {
                if (!searchLaporan.trim()) return true;
                const q = searchLaporan.trim().toLowerCase();
                return (
                  (w.nama || "").toLowerCase().includes(q) ||
                  (w.blok || "").toLowerCase().includes(q) ||
                  (w.username || "").toLowerCase().includes(q)
                );
              })
              .map((w) => (
              <tr key={w.id} className="hover:bg-gray-50">
                <td className="px-3 py-3 sticky left-0 bg-white z-10 whitespace-nowrap">
                  <div className="font-bold text-gray-900">{w.nama}</div>
                  <div className="text-[11px] text-gray-500">{w.blok}</div>
                </td>
                {BULAN_LIST.map((_, i) => {
                  const ym = `${tahun}-${String(i + 1).padStart(2, "0")}`;
                  const st = getStatus(w.uid || w.id, ym);
                  return (
                    <td key={ym} className="px-2 py-3 text-center">
                      {st === "LUNAS" ? (
                        <span className="inline-block px-2 py-0.5 rounded-md bg-emerald-50 text-emerald-700 font-bold">
                          ✓
                        </span>
                      ) : st === "PROSES" ? (
                        <span className="inline-block px-2 py-0.5 rounded-md bg-amber-50 text-amber-700 font-bold">
                          P
                        </span>
                      ) : (
                        <span className="text-gray-300">-</span>
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
// ==========================================
// ADMIN 4: DATA WARGA (IMPORT & EXPORT EXCEL / PDF)
// ==========================================
function AdminDataWarga({ wargaOnly, showToast, onImpersonateWarga }) {
  const [importing, setImporting] = useState(false);
  const [editing, setEditing] = useState(null);

  // Export Data Warga ke Excel (.xlsx)
  const handleExportExcel = () => {
    const rows = wargaOnly.map((w, idx) => ({
      No: idx + 1,
      "Nama Lengkap": w.nama || "-",
      "Blok/No.": w.blok || "-",
      Username: w.username || "-",
      "No. HP": w.hp || "-",
      "Status Rumah": w.status_rumah || "Milik Sendiri",
      WiFi: w.wifi || "-",
      Asal: w.asal || "-",
      "Isi Rumah (Keluarga)": w.isi_rumah || "-",
      Status: w.status || "Aktif",
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [
      { wch: 5 },
      { wch: 26 },
      { wch: 13 },
      { wch: 12 },
      { wch: 16 },
      { wch: 15 },
      { wch: 14 },
      { wch: 16 },
      { wch: 28 },
      { wch: 12 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Data Warga");
    const tgl = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `Data_Warga_Blok9_BSR_${tgl}.xlsx`);
  };

  // Export Data Warga ke PDF (Cetak / Simpan sebagai PDF A4 Landscape)
  const handleExportPDF = () => {
    const tgl = new Date().toLocaleDateString("id-ID", {
      day: "2-digit",
      month: "long",
      year: "numeric",
    });
    const rowsHtml = wargaOnly
      .map(
        (w, idx) => `
      <tr>
        <td style="text-align:center">${idx + 1}</td>
        <td><b>${w.nama || "-"}</b></td>
        <td>${w.blok || "-"}</td>
        <td>${w.username || "-"}</td>
        <td>${w.hp || "-"}</td>
        <td>${w.status_rumah || "Milik Sendiri"}</td>
        <td>${w.wifi || "-"}</td>
        <td>${w.asal || "-"}</td>
        <td>${w.isi_rumah || "-"}</td>
        <td style="text-align:center">${w.status || "Aktif"}</td>
      </tr>`
      )
      .join("");

    const win = window.open("", "_blank");
    if (!win) {
      showToast("Pop-up diblokir browser. Izinkan pop-up untuk export PDF.", "error");
      return;
    }
    win.document.write(`
      <html>
        <head>
          <title>Data Warga BLOK 9 BSR</title>
          <style>
            @page { size: A4 landscape; margin: 12mm; }
            body { font-family: Arial, sans-serif; color: #111827; font-size: 11px; }
            h2 { margin: 0 0 4px 0; font-size: 16px; }
            .meta { color: #6b7280; margin-bottom: 12px; font-size: 11px; }
            table { width: 100%; border-collapse: collapse; }
            th, td { border: 1px solid #d1d5db; padding: 6px 8px; text-align: left; }
            th { background: #f3f4f6; font-size: 10px; text-transform: uppercase; }
            tr:nth-child(even) { background: #f9fafb; }
          </style>
        </head>
        <body>
          <h2>DAFTAR DATA WARGA BLOK 9 BUMI SAFIR RAYA</h2>
          <div class="meta">Dicetak pada: ${tgl} &bull; Total Warga: ${wargaOnly.length} Rumah</div>
          <table>
            <thead>
              <tr>
                <th style="text-align:center">No</th>
                <th>Nama Lengkap</th>
                <th>Blok/No</th>
                <th>Username</th>
                <th>No. HP</th>
                <th>Status Rumah</th>
                <th>WiFi</th>
                <th>Asal</th>
                <th>Isi Rumah (Keluarga)</th>
                <th style="text-align:center">Status</th>
              </tr>
            </thead>
            <tbody>${rowsHtml}</tbody>
          </table>
          <script>window.onload = () => { window.print(); };</script>
        </body>
      </html>
    `);
    win.document.close();
  };

  // Import Data & Akun Warga dari file Excel (.xlsx / .xls / .csv)
  const handleImportExcel = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = "";

    setImporting(true);
    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      let updated = 0;
      let created = 0;
      let skipped = 0;

      for (const row of rows) {
        const nama = String(
          row["Nama Lengkap"] || row["Nama"] || row["nama"] || ""
        ).trim();
        const blok = String(
          row["Blok/No."] || row["Blok/No"] || row["Blok"] || row["blok"] || ""
        ).trim();
        let hp = String(
          row["No. HP"] || row["No HP"] || row["HP"] || row["hp"] || ""
        )
          .replace(/\s+/g, "")
          .trim();

        const statusRumahRaw = String(
          row["Status Rumah"] ||
            row["Kepemilikan"] ||
            row["status_rumah"] ||
            "Milik Sendiri"
        ).trim();
        const status_rumah = statusRumahRaw.toLowerCase().includes("kontrak")
          ? "Kontrak"
          : "Milik Sendiri";

        const wifi = String(row["WiFi"] || row["Wifi"] || row["wifi"] || "").trim();
        const asal = String(row["Asal"] || row["asal"] || "").trim();
        const isi_rumah = String(
          row["Isi Rumah (Keluarga)"] ||
            row["Isi Rumah"] ||
            row["Keluarga"] ||
            row["isi_rumah"] ||
            ""
        ).trim();

        const statusInput = String(row["Status"] || row["status"] || "Aktif").trim();
        const status =
          statusInput.toLowerCase() === "tidak aktif" ? "Tidak Aktif" : "Aktif";

        if (hp && hp.startsWith("8")) hp = "0" + hp;

        if (!nama || nama.toLowerCase().startsWith("contoh:") || !blok) {
          skipped++;
          continue;
        }

        const numMatch = blok.match(/(\d+)\s*$/);
        const nomorRumah = numMatch ? parseInt(numMatch[1], 10) : null;

        const uname = String(
          row["Username"] || row["username"] || (nomorRumah ? `bsr${nomorRumah}` : "")
        )
          .trim()
          .toLowerCase();
        const pass = String(row["Password"] || row["password"] || uname).trim();

        const existing = wargaOnly.find(
          (w) =>
            (uname && (w.username || "").toLowerCase() === uname) ||
            (w.blok || "").toLowerCase().replace(/\s+/g, "") ===
              blok.toLowerCase().replace(/\s+/g, "") ||
            (nomorRumah && w.nomor_rumah === nomorRumah)
        );

        if (existing) {
          await updateDoc(doc(db, "data_warga", existing.id), {
            nama,
            blok,
            hp,
            status_rumah,
            wifi,
            asal,
            isi_rumah,
            status,
          });
          updated++;
        } else if (uname) {
          try {
            const cred = await createUserWithEmailAndPassword(
              secondaryAuth,
              toFirebaseEmail(uname),
              toFirebasePassword(pass)
            );
            await setDoc(doc(db, "data_warga", cred.user.uid), {
              uid: cred.user.uid,
              nama,
              blok,
              nomor_rumah: nomorRumah || 999,
              hp,
              status_rumah,
              wifi,
              asal,
              isi_rumah,
              username: uname,
              status,
              role: "warga",
              createdAt: serverTimestamp(),
            });
            created++;
          } catch {
            skipped++;
          }
        } else {
          skipped++;
        }
      }

      showToast(
        `Import selesai! Diperbarui: ${updated}, Baru: ${created}, Dilewati: ${skipped}`
      );
    } catch (err) {
      showToast("Gagal membaca file Excel: " + err.message, "error");
    } finally {
      setImporting(false);
    }
  };

  const handleSaveEdit = async (e) => {
    e.preventDefault();
    const payload = {
      nama: editing.nama,
      blok: editing.blok,
      username: editing.username || "", // <-- tambahkan baris ini
      hp: editing.hp || "",
      status_rumah: editing.status_rumah || "Milik Sendiri",
      wifi: editing.wifi || "",
      asal: editing.asal || "",
      isi_rumah: editing.isi_rumah || "",
      status: editing.status || "Aktif",
    };

    if (editing.id) {
      await updateDoc(doc(db, "data_warga", editing.id), payload);
      showToast("Data warga diperbarui.");
    } else {
      await addDoc(collection(db, "data_warga"), {
        ...payload,
        role: "warga",
        createdAt: serverTimestamp(),
      });
      showToast("Warga baru berhasil ditambahkan.");
    }
    setEditing(null);
  };

  const handleDeleteWarga = async (w) => {
    if (!window.confirm(`Yakin ingin menghapus data warga "${w.nama}" (${w.blok})?`)) return;
    try {
      await deleteDoc(doc(db, "data_warga", w.id));
      setEditing(null);
      showToast(`Data ${w.nama} berhasil dihapus.`);
    } catch (err) {
      console.error(err);
      showToast("Gagal menghapus data warga!", "error");
    }
  };

const [searchWarga, setSearchWarga] = useState("");

  return (
    <div className="space-y-4">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h3 className="text-xl font-bold text-gray-900">Data Warga</h3>
          <p className="text-sm text-gray-500">Kelola daftar warga Blok 9.</p>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <div className="relative w-full sm:w-56">
            <Search className="w-4 h-4 text-gray-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchWarga}
              onChange={(e) => setSearchWarga(e.target.value)}
              placeholder="Cari nama / no rumah..."
              className="w-full pl-9 pr-3 py-2 rounded-xl border border-gray-300 text-xs focus:ring-2 focus:ring-orange-500 outline-none bg-white"
            />
          </div>

          <button
            type="button"
            onClick={handleExportExcel}
            className="inline-flex items-center gap-1.5 px-3.5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold shadow-xs"
          >
            <Download className="w-4 h-4" /> Export Excel
          </button>

          <button
            type="button"
            onClick={handleExportPDF}
            className="inline-flex items-center gap-1.5 px-3.5 py-2.5 bg-red-600 hover:bg-red-700 text-white rounded-xl text-xs font-bold shadow-xs"
          >
            <Download className="w-4 h-4" /> Export PDF
          </button>
          <button
         type="button"
         onClick={() =>
           setEditing({
             nama: "",
             blok: "Blok 9 / ",
             username: "",
             hp: "",
             status_rumah: "Milik Sendiri",
             wifi: "",
             asal: "",
             isi_rumah: "",
             status: "Aktif",
           })
         }
         className="inline-flex items-center gap-1.5 px-3.5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold shadow-sm"
       >
         + Tambah Warga
       </button>

          <label
            className={`inline-flex items-center gap-1.5 px-3.5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold shadow-sm cursor-pointer ${
              importing ? "opacity-60 pointer-events-none" : ""
            }`}
          >
            <ArrowUpCircle className="w-4 h-4" />
            <span>{importing ? "Mengimpor..." : "Import Excel"}</span>
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleImportExcel}
              className="hidden"
            />
          </label>
        </div>
      </div>

      {/* Modal Edit Data Warga */}
      {editing && (
        <div className="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
          <form
            onSubmit={handleSaveEdit}
            className="bg-white rounded-2xl p-6 max-w-md w-full max-h-[90vh] overflow-y-auto space-y-3.5"
          >
            <h3 className="font-bold text-lg text-gray-900">Edit Data Warga</h3>

            <div>
              <label className="block text-xs font-semibold text-gray-600 mb-1">
                Nama Lengkap
              </label>
              <input
                type="text"
                required
                placeholder="Nama Lengkap"
                value={editing.nama || ""}
                onChange={(e) => setEditing({ ...editing, nama: e.target.value })}
                className="w-full px-3.5 py-2 rounded-xl border border-gray-300 text-sm"
              />
            </div>
<div>
             <label className="block text-xs font-semibold text-gray-600 mb-1">
               Username
             </label>
             <input
               type="text"
               placeholder="Contoh: warga56"
               value={editing.username || ""}
               onChange={(e) => setEditing({ ...editing, username: e.target.value })}
               className="w-full px-3.5 py-2 rounded-xl border border-gray-300 text-sm"
             />
           </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-semibold text-gray-600 mb-1">
                  Blok / No.
                </label>
                <input
                  type="text"
                  required
                  placeholder="Blok 9 / 1"
                  value={editing.blok || ""}
                  onChange={(e) => setEditing({ ...editing, blok: e.target.value })}
                  className="w-full px-3.5 py-2 rounded-xl border border-gray-300 text-sm"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-gray-600 mb-1">
                  No. HP
                </label>
                <input
                  type="text"
                  placeholder="08..."
                  value={editing.hp || ""}
                  onChange={(e) => setEditing({ ...editing, hp: e.target.value })}
                  className="w-full px-3.5 py-2 rounded-xl border border-gray-300 text-sm"
                />
              </div>
            </div>

            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="block text-xs font-semibold text-gray-600 mb-1">
                  Status Rumah
                </label>
                <select
                  value={editing.status_rumah || "Milik Sendiri"}
                  onChange={(e) =>
                    setEditing({ ...editing, status_rumah: e.target.value })
                  }
                  className="w-full px-3.5 py-2 rounded-xl border border-gray-300 text-sm bg-white"
                >
                  <option value="Milik Sendiri">Milik Sendiri</option>
                  <option value="Kontrak">Kontrak</option>
                </select>
              </div>
              <div>
                <label className="block text-xs font-semibold text-gray-600 mb-1">
                  WiFi yang Dipakai
                </label>
                <input
                  type="text"
                  placeholder="IndiHome / Biznet / -"
                  value={editing.wifi || ""}
                  onChange={(e) => setEditing({ ...editing, wifi: e.target.value })}
                  className="w-full px-3.5 py-2 rounded-xl border border-gray-300 text-sm"
                />
              </div>
            </div>

            <div>
              <label className="block text-xs font-semibold text-gray-600 mb-1">
                Asal
              </label>
              <input
                type="text"
                placeholder="Contoh: Lamongan / Surabaya"
                value={editing.asal || ""}
                onChange={(e) => setEditing({ ...editing, asal: e.target.value })}
                className="w-full px-3.5 py-2 rounded-xl border border-gray-300 text-sm"
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-gray-600 mb-1">
                Isi Rumah (Keluarga)
              </label>
              <input
                type="text"
                placeholder="Contoh: 4 Orang (Suami, Istri, 2 Anak)"
                value={editing.isi_rumah || ""}
                onChange={(e) =>
                  setEditing({ ...editing, isi_rumah: e.target.value })
                }
                className="w-full px-3.5 py-2 rounded-xl border border-gray-300 text-sm"
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-gray-600 mb-1">
                Status Akun
              </label>
              <select
                value={editing.status || "Aktif"}
                onChange={(e) =>
                  setEditing({ ...editing, status: e.target.value })
                }
                className="w-full px-3.5 py-2 rounded-xl border border-gray-300 text-sm bg-white"
              >
                <option value="Aktif">Aktif</option>
                <option value="Tidak Aktif">Tidak Aktif</option>
              </select>
            </div>

            <div className="flex gap-2 pt-2">
              <button
                type="button"
                onClick={() => setEditing(null)}
                className="flex-1 py-2.5 bg-gray-100 rounded-xl text-sm font-semibold"
              >
                Batal
              </button>
              <button
                type="submit"
                className="flex-1 py-2.5 bg-orange-500 text-white rounded-xl text-sm font-bold"
              >
                Simpan
              </button>
            </div>
          </form>
        </div>
      )}

      {/* Tabel Data Warga */}
      <div className="bg-white rounded-2xl border border-gray-200 overflow-x-auto">
        <table className="w-full text-sm text-left">
          <thead className="bg-gray-50 text-xs text-gray-500 uppercase border-b border-gray-200">
            <tr>
              <th className="px-4 py-3 whitespace-nowrap">Nama Lengkap</th>
              <th className="px-4 py-3 whitespace-nowrap">Blok/No</th>
              <th className="px-4 py-3 whitespace-nowrap">Username</th>
              <th className="px-4 py-3 whitespace-nowrap">No. HP</th>
              <th className="px-4 py-3 whitespace-nowrap">Status Rumah</th>
              <th className="px-4 py-3 whitespace-nowrap">WiFi</th>
              <th className="px-4 py-3 whitespace-nowrap">Asal</th>
              <th className="px-4 py-3 whitespace-nowrap">Isi Rumah (Keluarga)</th>
              <th className="px-4 py-3 whitespace-nowrap">Status</th>
              <th className="px-4 py-3 text-right whitespace-nowrap">Aksi</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {wargaOnly
              .filter((w) => {
                if (!searchWarga.trim()) return true;
                const q = searchWarga.trim().toLowerCase();
                return (
                  (w.nama || "").toLowerCase().includes(q) ||
                  (w.blok || "").toLowerCase().includes(q) ||
                  (w.username || "").toLowerCase().includes(q)
                );
              })
              .map((w) => (
              <tr key={w.id} className="hover:bg-gray-50">
                <td className="px-4 py-3.5 font-bold text-gray-900 whitespace-nowrap">
                  {w.nama}
                </td>
                <td className="px-4 py-3.5 text-gray-600 whitespace-nowrap">
                  {w.blok}
                </td>
                <td className="px-4 py-3.5 font-mono text-xs text-orange-600 font-semibold whitespace-nowrap">
                  {w.username}
                </td>
                <td className="px-4 py-3.5 text-gray-600 whitespace-nowrap">
                  {w.hp || "-"}
                </td>
                <td className="px-4 py-3.5 whitespace-nowrap">
                  <span
                    className={`px-2.5 py-1 rounded-lg text-xs font-bold ${
                      w.status_rumah === "Kontrak"
                        ? "bg-amber-50 text-amber-700"
                        : "bg-blue-50 text-blue-700"
                    }`}
                  >
                    {w.status_rumah || "Milik Sendiri"}
                  </span>
                </td>
                <td className="px-4 py-3.5 text-gray-600 whitespace-nowrap">
                  {w.wifi || "-"}
                </td>
                <td className="px-4 py-3.5 text-gray-600 whitespace-nowrap">
                  {w.asal || "-"}
                </td>
                <td className="px-4 py-3.5 text-gray-600 min-w-[180px]">
                  {w.isi_rumah || "-"}
                </td>
                <td className="px-4 py-3.5 whitespace-nowrap">
                  <span
                    className={`px-2.5 py-1 rounded-lg text-xs font-bold ${
                      w.status === "Tidak Aktif"
                        ? "bg-red-50 text-red-600"
                        : "bg-emerald-50 text-emerald-700"
                    }`}
                  >
                    {w.status || "Aktif"}
                  </span>
                </td>
                <td className="px-4 py-3.5 text-right space-x-2 whitespace-nowrap">
                  {w.hp && (
                    <a
                      href={`https://wa.me/${formatWaNumber(w.hp)}`}
                      target="_blank"
                      rel="noreferrer"
                      className="inline-flex p-1.5 text-emerald-600 hover:bg-emerald-50 rounded-lg"
                    >
                      <MessageCircle className="w-4 h-4" />
                    </a>
                  )}
          {onImpersonateWarga && (
                    <button
                      type="button"
                      onClick={() => onImpersonateWarga(w)}
                      title="Masuk ke akun warga ini"
                      className="inline-flex items-center px-2.5 py-1 bg-orange-500 hover:bg-orange-600 text-white rounded-lg text-xs font-bold"
                    >
                      Masuk →
                    </button>
                  )}
                  <button
                    onClick={() => setEditing(w)}
                    className="inline-flex p-1.5 text-blue-600 hover:bg-blue-50 rounded-lg"
                  >
                    <Edit2 className="w-4 h-4" />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
// ==========================================
// ADMIN 5: WARGA NUNGGAK (EXPORT EXCEL & PDF)
// ==========================================
function AdminWargaNunggak({ wargaOnly, pembayaranList }) {
  const wajib = getBulanWajibTahunIni();
  const fmtRp = (n) => "Rp " + Number(n || 0).toLocaleString("id-ID");

  const formatYmLabel = (ym) => {
    const [y, m] = ym.split("-");
    return `${BULAN_LIST[Number(m) - 1]?.slice(0, 3) || m} ${y}`;
  };

  const listNunggak = wargaOnly
    .map((w) => {
      const trxLunas = pembayaranList.filter(
        (p) => p.warga_uid === (w.uid || w.id) && p.status === "Lunas"
      );
      const lunasSet = new Set(trxLunas.flatMap((t) => t.bulan_dibayar || []));
      const bulanNunggak = wajib.filter((b) => !lunasSet.has(b));
      return {
        ...w,
        bulanNunggak,
        totalNunggak: bulanNunggak.length * CONFIG.iuranPerBulan,
      };
    })
    .filter((w) => w.bulanNunggak.length > 0 && w.status !== "Tidak Aktif");

  // Export Warga Nunggak ke Excel (.xlsx)
  const handleExportExcel = () => {
    const rows = listNunggak.map((w, idx) => ({
      No: idx + 1,
      "Nama Lengkap": w.nama || "-",
      "Blok/No.": w.blok || "-",
      "No. HP": w.hp || "-",
      "Jumlah Bulan": w.bulanNunggak.length,
      "Rincian Bulan": w.bulanNunggak.map(formatYmLabel).join(", "),
      "Total Tunggakan": w.totalNunggak,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [
      { wch: 5 },
      { wch: 26 },
      { wch: 13 },
      { wch: 16 },
      { wch: 14 },
      { wch: 36 },
      { wch: 18 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Warga Nunggak");
    const tgl = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `Warga_Nunggak_Blok9_BSR_${tgl}.xlsx`);
  };

  // Export Warga Nunggak ke PDF (A4 Landscape)
  const handleExportPDF = () => {
    const tgl = new Date().toLocaleDateString("id-ID", {
      day: "2-digit",
      month: "long",
      year: "numeric",
    });
    const grandTotal = listNunggak.reduce((acc, w) => acc + w.totalNunggak, 0);

    const rowsHtml = listNunggak
      .map(
        (w, idx) => `
      <tr>
        <td style="text-align:center">${idx + 1}</td>
        <td><b>${w.nama || "-"}</b></td>
        <td>${w.blok || "-"}</td>
        <td>${w.hp || "-"}</td>
        <td style="text-align:center">${w.bulanNunggak.length} Bln</td>
        <td>${w.bulanNunggak.map(formatYmLabel).join(", ")}</td>
        <td style="text-align:right;font-weight:bold;color:#dc2626">${fmtRp(
          w.totalNunggak
        )}</td>
      </tr>`
      )
      .join("");

    const win = window.open("", "_blank");
    if (!win) {
      alert("Pop-up diblokir browser. Izinkan pop-up untuk export PDF.");
      return;
    }
    win.document.write(`
      <html>
        <head>
          <title>Daftar Warga Nunggak BLOK 9 BSR</title>
          <style>
            @page { size: A4 landscape; margin: 12mm; }
            body { font-family: Arial, sans-serif; color: #111827; font-size: 11px; }
            h2 { margin: 0 0 4px 0; font-size: 16px; }
            .meta { color: #6b7280; margin-bottom: 12px; font-size: 11px; }
            table { width: 100%; border-collapse: collapse; }
            th, td { border: 1px solid #d1d5db; padding: 6px 8px; text-align: left; }
            th { background: #f3f4f6; font-size: 10px; text-transform: uppercase; }
            tr:nth-child(even) { background: #f9fafb; }
            .total-row td { background: #fef2f2; font-weight: bold; }
          </style>
        </head>
        <body>
          <h2>DAFTAR WARGA NUNGGAK IURAN BLOK 9 BSR</h2>
          <div class="meta">Dicetak pada: ${tgl} &bull; Total Warga Nunggak: ${
            listNunggak.length
          } Rumah &bull; Total Tunggakan: ${fmtRp(grandTotal)}</div>
          <table>
            <thead>
              <tr>
                <th style="text-align:center">No</th>
                <th>Nama Lengkap</th>
                <th>Blok/No</th>
                <th>No. HP</th>
                <th style="text-align:center">Jml Bulan</th>
                <th>Rincian Bulan Tunggakan</th>
                <th style="text-align:right">Total Tunggakan</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml}
              <tr class="total-row">
                <td colspan="6" style="text-align:right">GRAND TOTAL TUNGGAKAN</td>
                <td style="text-align:right;color:#dc2626">${fmtRp(grandTotal)}</td>
              </tr>
            </tbody>
          </table>
          <script>window.onload = () => { window.print(); };</script>
        </body>
      </html>
    `);
    win.document.close();
  };
const [searchNunggak, setSearchNunggak] = useState("");
  const filteredNunggak = listNunggak.filter((item) => {
    if (!searchNunggak.trim()) return true;
    const q = searchNunggak.toLowerCase();
    const nama = (item.nama || item.namaWarga || "").toLowerCase();
    const rumah = (item.noRumah || item.rumah || item.username || "").toLowerCase();
    return nama.includes(q) || rumah.includes(q);
  });
  return (
    <div className="space-y-4">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div>
          <h3 className="text-xl font-bold text-gray-900">Warga Nunggak</h3>
          <p className="text-sm text-gray-500">
            Daftar warga yang belum melunasi iuran sampai bulan ini ({listNunggak.length} rumah).
          </p>
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <div className="relative w-full sm:w-56">
            <Search className="w-4 h-4 text-gray-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchNunggak}
              onChange={(e) => setSearchNunggak(e.target.value)}
              placeholder="Cari nama / no rumah..."
              className="w-full pl-9 pr-3 py-2 rounded-xl border border-gray-300 text-xs focus:ring-2 focus:ring-orange-500 outline-none bg-white"
            />
          </div>
          <button
            type="button"
            onClick={handleExportExcel}
            className="inline-flex items-center gap-1.5 px-3.5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold shadow-xs"
          >
            <Download className="w-4 h-4" /> Export Excel
          </button>

          <button
            type="button"
            onClick={handleExportPDF}
            className="inline-flex items-center gap-1.5 px-3.5 py-2.5 bg-red-600 hover:bg-red-700 text-white rounded-xl text-xs font-bold shadow-xs"
          >
            <Download className="w-4 h-4" /> Export PDF
          </button>
        </div>
      </div>

      <div className="bg-white rounded-2xl border border-gray-200 overflow-x-auto">
      <table className="w-full text-sm text-left">
          <thead className="bg-gray-50 text-xs text-gray-500 uppercase border-b border-gray-200">
            <tr>
              <th className="px-4 py-3">Nama Lengkap</th>
              <th className="px-4 py-3">Blok/No</th>
              <th className="px-4 py-3">Tunggakan</th>
              <th className="px-4 py-3">Nominal</th>
              <th className="px-4 py-3 text-right">Ingatkan WA</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {filteredNunggak.length === 0 ? (
              <tr>
                <td colSpan={5} className="px-4 py-8 text-center text-gray-400">
                  {searchNunggak
                    ? "Data warga tidak ditemukan."
                    : "Semua warga aktif sudah lunas sampai bulan ini!"}
                </td>
              </tr>
            ) : (
              filteredNunggak.map((w) => {
                const pesan = encodeURIComponent(
                  `Assalamu'alaikum ${w.nama} (${w.blok}), mengingatkan iuran kas Blok 9 BSR sebesar ${fmtRp(
                    w.totalNunggak
                  )} untuk ${w.bulanNunggak.length} bulan (${w.bulanNunggak
                    .map(formatYmLabel)
                    .join(", ")}). Terima kasih.`
                );
                return (
                  <tr key={w.id} className="hover:bg-gray-50">
                    <td className="px-4 py-3.5 font-bold text-gray-900">
                      {w.nama}
                    </td>
                    <td className="px-4 py-3.5 text-gray-600">{w.blok}</td>
                    <td className="px-4 py-3.5">
                      <span className="px-2.5 py-1 rounded-lg text-xs font-bold bg-red-50 text-red-600">
                        {w.bulanNunggak.length} Bulan
                      </span>
                      <div className="text-xs text-gray-500 mt-1">
                        {w.bulanNunggak.map(formatYmLabel).join(", ")}
                      </div>
                    </td>
                    <td className="px-4 py-3.5 font-bold text-red-600">
                      {fmtRp(w.totalNunggak)}
                    </td>
                    <td className="px-4 py-3.5 text-right">
                      {w.hp ? (
                        <a
                          href={`https://wa.me/${formatWaNumber(w.hp)}?text=${pesan}`}
                          target="_blank"
                          rel="noreferrer"
                          className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-bold"
                        >
                          <MessageCircle className="w-3.5 h-3.5" /> Kirim WA
                        </a>
                      ) : (
                        <span className="text-xs text-gray-400">No HP kosong</span>
                      )}
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
// ==========================================
// ADMIN 6: SETOR KAS (PEMASUKAN / PENGELUARAN / SETORAN)
// ==========================================
function AdminSetorKas({ setorKasList, saldoSummary, showToast }) {
  const [jenisKas, setJenisKas] = useState("paguyuban");
  const [tipe, setTipe] = useState("Pengeluaran");
  const [nominal, setNominal] = useState("");
  const [keterangan, setKeterangan] = useState("");
  const [tanggal, setTanggal] = useState(new Date().toISOString().slice(0, 10));

  const handleSubmit = async (e) => {
    e.preventDefault();
    await addDoc(collection(db, "setor_kas"), {
      jenis_kas: jenisKas,
      tipe,
      nominal: Number(nominal),
      keterangan,
      tanggal,
      createdAt: serverTimestamp(),
    });
    setNominal("");
    setKeterangan("");
    showToast("Catatan Setor Kas berhasil disimpan!");
  };

  return (
    <div className="space-y-5">
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <div className="bg-blue-600 text-white p-5 rounded-2xl">
          <p className="text-xs text-blue-100">Saldo Kas Paguyuban BSR</p>
          <p className="text-2xl font-extrabold mt-1">
            {formatRupiah(saldoSummary.saldoPaguyuban)}
          </p>
        </div>
        <div className="bg-orange-500 text-white p-5 rounded-2xl">
          <p className="text-xs text-orange-100">Saldo Kas Khusus BLOK 9</p>
          <p className="text-2xl font-extrabold mt-1">
            {formatRupiah(saldoSummary.saldoBlok9)}
          </p>
        </div>
      </div>

      <form
        onSubmit={handleSubmit}
        className="bg-white p-5 rounded-2xl border border-gray-200 grid grid-cols-1 sm:grid-cols-5 gap-3"
      >
        <select
          value={jenisKas}
          onChange={(e) => setJenisKas(e.target.value)}
          className="px-3 py-2 rounded-xl border border-gray-300 text-sm"
        >
          <option value="paguyuban">Kas Paguyuban BSR</option>
          <option value="blok9">Kas Khusus BLOK 9</option>
        </select>
        <select
          value={tipe}
          onChange={(e) => setTipe(e.target.value)}
          className="px-3 py-2 rounded-xl border border-gray-300 text-sm"
        >
          <option value="Pengeluaran">Setor / Pengeluaran</option>
          <option value="Pemasukan">Pemasukan Tambahan</option>
        </select>
        <input
          type="date"
          required
          value={tanggal}
          onChange={(e) => setTanggal(e.target.value)}
          className="px-3 py-2 rounded-xl border border-gray-300 text-sm"
        />
        <input
          type="number"
          required
          placeholder="Nominal (Rp)"
          value={nominal}
          onChange={(e) => setNominal(e.target.value)}
          className="px-3 py-2 rounded-xl border border-gray-300 text-sm"
        />
        <input
          type="text"
          required
          placeholder="Keterangan (mis: Setor ke Paguyuban)"
          value={keterangan}
          onChange={(e) => setKeterangan(e.target.value)}
          className="px-3 py-2 rounded-xl border border-gray-300 text-sm sm:col-span-4"
        />
        <button
          type="submit"
          className="py-2 bg-orange-500 hover:bg-orange-600 text-white rounded-xl text-sm font-bold"
        >
          Simpan
        </button>
      </form>

      <div className="bg-white rounded-2xl border border-gray-200 overflow-x-auto">
        <table className="w-full text-sm text-left">
          <thead className="bg-gray-50 text-xs text-gray-500 uppercase border-b border-gray-200">
            <tr>
              <th className="px-4 py-3">Tanggal</th>
              <th className="px-4 py-3">Dompet Kas</th>
              <th className="px-4 py-3">Keterangan</th>
              <th className="px-4 py-3 text-right">Nominal</th>
              <th className="px-4 py-3 text-right">Hapus</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-100">
            {setorKasList.map((s) => (
              <tr key={s.id}>
                <td className="px-4 py-3 text-gray-500">{s.tanggal}</td>
                <td className="px-4 py-3 font-semibold">
                  {s.jenis_kas === "paguyuban" ? "Kas Paguyuban" : "Kas Blok 9"}
                </td>
                <td className="px-4 py-3">{s.keterangan}</td>
                <td
                  className={`px-4 py-3 text-right font-bold ${
                    s.tipe === "Pemasukan" ? "text-emerald-600" : "text-red-600"
                  }`}
                >
                  {s.tipe === "Pemasukan" ? "+" : "-"} {formatRupiah(s.nominal)}
                </td>
                <td className="px-4 py-3 text-right">
                  <button
                    onClick={() => deleteDoc(doc(db, "setor_kas", s.id))}
                    className="text-red-500 hover:text-red-700"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ==========================================
// ADMIN 7 & WARGA 3: PENGUMUMAN
// ==========================================
function PengumumanView({ isAdmin, list, showToast }) {
  const [judul, setJudul] = useState("");
  const [isi, setIsi] = useState("");

  const handleAdd = async (e) => {
    e.preventDefault();
    await addDoc(collection(db, "pengumuman"), {
      judul,
      isi,
      tanggal: new Date().toISOString().slice(0, 10),
      createdAt: serverTimestamp(),
    });
    setJudul("");
    setIsi("");
    showToast("Pengumuman berhasil dipublikasikan!");
  };

  return (
    <div className="space-y-5 max-w-3xl">
      <h3 className="text-xl font-bold text-gray-900">Pengumuman BLOK 9 BSR</h3>

      {isAdmin && (
        <form
          onSubmit={handleAdd}
          className="bg-white p-5 rounded-2xl border border-gray-200 space-y-3"
        >
          <input
            type="text"
            required
            placeholder="Judul Pengumuman (mis: Kerja Bakti Rutin)"
            value={judul}
            onChange={(e) => setJudul(e.target.value)}
            className="w-full px-3.5 py-2.5 rounded-xl border border-gray-300 text-sm"
          />
          <textarea
            required
            rows={3}
            placeholder="Isi pengumuman untuk seluruh warga..."
            value={isi}
            onChange={(e) => setIsi(e.target.value)}
            className="w-full px-3.5 py-2.5 rounded-xl border border-gray-300 text-sm"
          />
          <button
            type="submit"
            className="px-5 py-2.5 bg-orange-500 hover:bg-orange-600 text-white rounded-xl text-sm font-bold"
          >
            + Posting Pengumuman
          </button>
        </form>
      )}

      <div className="space-y-3">
        {list.length === 0 ? (
          <p className="text-sm text-gray-400">Belum ada pengumuman.</p>
        ) : (
          list.map((item) => (
            <div
              key={item.id}
              className="bg-white rounded-2xl border border-gray-200 border-l-4 border-l-orange-500 p-5 shadow-xs flex justify-between gap-4"
            >
              <div>
                <h4 className="font-bold text-gray-900">{item.judul}</h4>
                <p className="text-sm text-gray-600 mt-1">{item.isi}</p>
              </div>
              <div className="flex flex-col items-end gap-2 shrink-0">
                <span className="px-2.5 py-1 rounded-lg bg-orange-50 text-orange-600 text-xs font-semibold">
                  {formatTanggalIndo(item.tanggal)}
                </span>
                {isAdmin && (
                  <button
                    onClick={() => deleteDoc(doc(db, "pengumuman", item.id))}
                    className="text-red-500 hover:text-red-700"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                )}
              </div>
            </div>
          ))
        )}
      </div>
    </div>
  );
}

// ==========================================
// ADMIN 8 & WARGA 4: TATA TERTIB
// ==========================================
function TataTertibView({ isAdmin, list, showToast }) {
  const [isi, setIsi] = useState("");

  const handleAdd = async (e) => {
    e.preventDefault();
    await addDoc(collection(db, "tata_tertib"), {
      urutan: list.length + 1,
      isi,
      createdAt: serverTimestamp(),
    });
    setIsi("");
    showToast("Poin tata tertib ditambahkan!");
  };

  return (
    <div className="bg-white rounded-2xl border border-gray-200 p-6 sm:p-8 max-w-3xl space-y-6">
      <div className="flex items-center gap-3 pb-4 border-b border-gray-100">
        <div className="w-11 h-11 bg-blue-50 text-blue-600 rounded-xl flex items-center justify-center">
          <BookOpen className="w-6 h-6" />
        </div>
        <div>
          <h3 className="text-lg font-bold text-gray-900">Tata Tertib Warga</h3>
          <p className="text-xs text-gray-500">Peraturan lingkungan BLOK 9 BSR</p>
        </div>
      </div>

      {isAdmin && (
        <form onSubmit={handleAdd} className="flex gap-2">
          <input
            type="text"
            required
            placeholder="Tambah poin tata tertib baru..."
            value={isi}
            onChange={(e) => setIsi(e.target.value)}
            className="flex-1 px-3.5 py-2.5 rounded-xl border border-gray-300 text-sm"
          />
          <button
            type="submit"
            className="px-4 py-2.5 bg-orange-500 hover:bg-orange-600 text-white rounded-xl text-sm font-bold"
          >
            Tambah
          </button>
        </form>
      )}

      <div className="space-y-4">
        {list.length === 0 ? (
          <p className="text-sm text-gray-400">Belum ada tata tertib yang ditambahkan.</p>
        ) : (
          list.map((item, idx) => (
            <div key={item.id} className="flex items-start justify-between gap-3">
              <div className="flex items-start gap-3">
                <span className="w-7 h-7 rounded-full bg-orange-100 text-orange-600 font-bold text-xs flex items-center justify-center shrink-0 mt-0.5">
                  {idx + 1}
                </span>
                <p className="text-sm text-gray-700 leading-relaxed">{item.isi}</p>
              </div>
              {isAdmin && (
  <button
    onClick={() => bukaWA(selectedWarga)}
    className="w-full py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-sm"
  >
    💬 Ingatkan via WhatsApp
  </button>
)}
            </div>
          ))
        )}
      </div>
    </div>
  );
}
