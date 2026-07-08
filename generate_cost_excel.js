// Kalimera Kitchen — Maliyet Hesap Tablosu Üreticisi
// Çalıştırma: node generate_cost_excel.js
// Çıktı: Kalimera_Maliyet_Hesabi.xlsx

const ExcelJS = require('exceljs');
const path = require('path');

// ============================================================
// 1) MALZEME ANA LİSTESİ — TEK FİYAT KAYNAĞI
// Birim: kg, lt, adet, demet
// Fiyat: TL (Türkiye 2026 toptan tahmini — güncellemen yeterli)
// ============================================================
const INGREDIENTS = [
  // KATEGORİ: ET
  { id: 'M001', kategori: 'Et',           ad: 'Kuzu But',                    birim: 'kg',    fiyat: 850 },
  { id: 'M002', kategori: 'Et',           ad: 'Kuzu Eti (kuşbaşı)',          birim: 'kg',    fiyat: 800 },
  { id: 'M003', kategori: 'Et',           ad: 'Dana Eti (kuşbaşı/sote)',     birim: 'kg',    fiyat: 700 },
  { id: 'M004', kategori: 'Et',           ad: 'Dana Kıyma',                  birim: 'kg',    fiyat: 600 },
  { id: 'M005', kategori: 'Et',           ad: 'Tavuk But Fileto (kemiksiz)', birim: 'kg',    fiyat: 200 },
  { id: 'M006', kategori: 'Et',           ad: 'Tavuk Göğüs',                 birim: 'kg',    fiyat: 220 },
  { id: 'M007', kategori: 'Et',           ad: 'Tavuk Kanat',                 birim: 'kg',    fiyat: 180 },
  { id: 'M008', kategori: 'Et',           ad: 'Tavuk Pizola',                birim: 'kg',    fiyat: 200 },
  { id: 'M009', kategori: 'Et',           ad: 'Sucuk',                       birim: 'kg',    fiyat: 450 },
  // KATEGORİ: BALIK & DENİZ
  { id: 'M010', kategori: 'Deniz',        ad: 'Levrek Fileto',               birim: 'kg',    fiyat: 600 },
  { id: 'M011', kategori: 'Deniz',        ad: 'Somon Fileto',                birim: 'kg',    fiyat: 850 },
  { id: 'M012', kategori: 'Deniz',        ad: 'Karides (orta boy)',          birim: 'kg',    fiyat: 700 },
  { id: 'M013', kategori: 'Deniz',        ad: 'Kalamar',                     birim: 'kg',    fiyat: 600 },
  { id: 'M014', kategori: 'Deniz',        ad: 'Midye',                       birim: 'kg',    fiyat: 250 },
  { id: 'M015', kategori: 'Deniz',        ad: 'Füme Somon',                  birim: 'kg',    fiyat: 1200 },
  // KATEGORİ: SEBZE & MEYVE
  { id: 'M016', kategori: 'Sebze',        ad: 'Domates',                     birim: 'kg',    fiyat: 35 },
  { id: 'M017', kategori: 'Sebze',        ad: 'Kiraz Domates',               birim: 'kg',    fiyat: 90 },
  { id: 'M018', kategori: 'Sebze',        ad: 'Pembe Domates',               birim: 'kg',    fiyat: 60 },
  { id: 'M019', kategori: 'Sebze',        ad: 'Salatalık',                   birim: 'kg',    fiyat: 30 },
  { id: 'M020', kategori: 'Sebze',        ad: 'Soğan (kuru)',                birim: 'kg',    fiyat: 25 },
  { id: 'M021', kategori: 'Sebze',        ad: 'Sarımsak',                    birim: 'kg',    fiyat: 220 },
  { id: 'M022', kategori: 'Sebze',        ad: 'Patates',                     birim: 'kg',    fiyat: 30 },
  { id: 'M023', kategori: 'Sebze',        ad: 'Körpe Patates',               birim: 'kg',    fiyat: 60 },
  { id: 'M024', kategori: 'Sebze',        ad: 'Patlıcan',                    birim: 'kg',    fiyat: 35 },
  { id: 'M025', kategori: 'Sebze',        ad: 'Kabak',                       birim: 'kg',    fiyat: 30 },
  { id: 'M026', kategori: 'Sebze',        ad: 'Havuç',                       birim: 'kg',    fiyat: 25 },
  { id: 'M027', kategori: 'Sebze',        ad: 'Dolma Biber',                 birim: 'kg',    fiyat: 50 },
  { id: 'M028', kategori: 'Sebze',        ad: 'Renkli Biber (kırmızı/sarı/yeşil)', birim: 'kg', fiyat: 90 },
  { id: 'M029', kategori: 'Sebze',        ad: 'Kırmızı Lahana',              birim: 'kg',    fiyat: 35 },
  { id: 'M030', kategori: 'Sebze',        ad: 'Mantar (kültür)',             birim: 'kg',    fiyat: 120 },
  { id: 'M031', kategori: 'Sebze',        ad: 'Karnabahar',                  birim: 'kg',    fiyat: 50 },
  { id: 'M032', kategori: 'Sebze',        ad: 'Taze Fasulye',                birim: 'kg',    fiyat: 60 },
  { id: 'M033', kategori: 'Meyve',        ad: 'Yeşil Elma',                  birim: 'kg',    fiyat: 65 },
  { id: 'M034', kategori: 'Meyve',        ad: 'Avokado',                     birim: 'adet',  fiyat: 60 },
  { id: 'M035', kategori: 'Yeşillik',     ad: 'Roka',                        birim: 'kg',    fiyat: 90 },
  { id: 'M036', kategori: 'Yeşillik',     ad: 'Marul / Yeşillik Karışım',    birim: 'kg',    fiyat: 70 },
  { id: 'M037', kategori: 'Yeşillik',     ad: 'Semizotu',                    birim: 'kg',    fiyat: 50 },
  { id: 'M038', kategori: 'Ot',           ad: 'Maydanoz',                    birim: 'demet', fiyat: 10 },
  { id: 'M039', kategori: 'Ot',           ad: 'Dereotu',                     birim: 'demet', fiyat: 12 },
  { id: 'M040', kategori: 'Ot',           ad: 'Nane',                        birim: 'demet', fiyat: 12 },
  { id: 'M041', kategori: 'Ot',           ad: 'Adaçayı (taze)',              birim: 'demet', fiyat: 25 },
  { id: 'M042', kategori: 'Ot',           ad: 'Asma Yaprağı (salamura)',     birim: 'kg',    fiyat: 200 },
  { id: 'M043', kategori: 'Meyve',        ad: 'Limon',                       birim: 'kg',    fiyat: 50 },
  { id: 'M044', kategori: 'Meyve',        ad: 'Karpuz',                      birim: 'kg',    fiyat: 25 },
  { id: 'M045', kategori: 'Meyve',        ad: 'Kavun',                       birim: 'kg',    fiyat: 35 },
  { id: 'M046', kategori: 'Meyve',        ad: 'Şeftali',                     birim: 'kg',    fiyat: 60 },
  { id: 'M047', kategori: 'Meyve',        ad: 'Üzüm',                        birim: 'kg',    fiyat: 80 },
  { id: 'M048', kategori: 'Meyve',        ad: 'Çilek',                       birim: 'kg',    fiyat: 90 },
  { id: 'M049', kategori: 'Meyve',        ad: 'Portakal (sıkmalık)',         birim: 'kg',    fiyat: 35 },
  { id: 'M050', kategori: 'Meyve',        ad: 'Nar Tanesi',                  birim: 'kg',    fiyat: 90 },
  // KATEGORİ: BAKLAGİL / TAHIL
  { id: 'M051', kategori: 'Bakliyat',     ad: 'Pirinç (Baldo)',              birim: 'kg',    fiyat: 90 },
  { id: 'M052', kategori: 'Bakliyat',     ad: 'Şehriye',                     birim: 'kg',    fiyat: 70 },
  { id: 'M053', kategori: 'Bakliyat',     ad: 'Nohut (kuru)',                birim: 'kg',    fiyat: 80 },
  { id: 'M054', kategori: 'Bakliyat',     ad: 'Beyaz Fasulye (kuru)',        birim: 'kg',    fiyat: 90 },
  { id: 'M055', kategori: 'Tahıl',        ad: 'Un',                          birim: 'kg',    fiyat: 30 },
  // KATEGORİ: KURUYEMİŞ
  { id: 'M056', kategori: 'Kuruyemiş',    ad: 'Badem',                       birim: 'kg',    fiyat: 800 },
  { id: 'M057', kategori: 'Kuruyemiş',    ad: 'Ceviz İçi',                   birim: 'kg',    fiyat: 700 },
  { id: 'M058', kategori: 'Kuruyemiş',    ad: 'Balkabağı Çekirdeği',         birim: 'kg',    fiyat: 350 },
  { id: 'M059', kategori: 'Kuruyemiş',    ad: 'Kuş Üzümü',                   birim: 'kg',    fiyat: 350 },
  { id: 'M060', kategori: 'Kuruyemiş',    ad: 'Kuru İncir',                  birim: 'kg',    fiyat: 600 },
  // KATEGORİ: SÜT & PEYNİR
  { id: 'M061', kategori: 'Süt',          ad: 'Süt (tam yağlı)',             birim: 'lt',    fiyat: 35 },
  { id: 'M062', kategori: 'Süt',          ad: 'Yoğurt',                      birim: 'kg',    fiyat: 80 },
  { id: 'M063', kategori: 'Süt',          ad: 'Süzme Yoğurt',                birim: 'kg',    fiyat: 130 },
  { id: 'M064', kategori: 'Süt',          ad: 'Kaymak',                      birim: 'kg',    fiyat: 600 },
  { id: 'M065', kategori: 'Süt',          ad: 'Tereyağı',                    birim: 'kg',    fiyat: 450 },
  { id: 'M066', kategori: 'Süt',          ad: 'Krema (sıvı)',                birim: 'lt',    fiyat: 200 },
  { id: 'M067', kategori: 'Peynir',       ad: 'Beyaz Peynir',                birim: 'kg',    fiyat: 280 },
  { id: 'M068', kategori: 'Peynir',       ad: 'Ezine Peyniri',               birim: 'kg',    fiyat: 380 },
  { id: 'M069', kategori: 'Peynir',       ad: 'Eski Kaşar',                  birim: 'kg',    fiyat: 450 },
  { id: 'M070', kategori: 'Peynir',       ad: 'Kaşar',                       birim: 'kg',    fiyat: 350 },
  { id: 'M071', kategori: 'Peynir',       ad: 'Tulum Peyniri',               birim: 'kg',    fiyat: 500 },
  { id: 'M072', kategori: 'Peynir',       ad: 'Köy Peyniri',                 birim: 'kg',    fiyat: 320 },
  { id: 'M073', kategori: 'Süt',          ad: 'Vanilyalı Dondurma',          birim: 'lt',    fiyat: 200 },
  // KATEGORİ: YAĞ
  { id: 'M074', kategori: 'Yağ',          ad: 'Zeytinyağı (sızma)',          birim: 'lt',    fiyat: 280 },
  { id: 'M075', kategori: 'Yağ',          ad: 'Ayçiçek Yağı',                birim: 'lt',    fiyat: 80 },
  // KATEGORİ: SOS / TATLI MALZEMESİ
  { id: 'M076', kategori: 'Sos',          ad: 'Tahin',                       birim: 'kg',    fiyat: 250 },
  { id: 'M077', kategori: 'Sos',          ad: 'Bal (çiçek)',                 birim: 'kg',    fiyat: 600 },
  { id: 'M078', kategori: 'Sos',          ad: 'Üzüm Pekmezi',                birim: 'kg',    fiyat: 200 },
  { id: 'M079', kategori: 'Sos',          ad: 'Yeşil Zeytin',                birim: 'kg',    fiyat: 180 },
  { id: 'M080', kategori: 'Sos',          ad: 'Siyah Zeytin',                birim: 'kg',    fiyat: 200 },
  { id: 'M081', kategori: 'Sos',          ad: 'Çilek Reçeli',                birim: 'kg',    fiyat: 180 },
  { id: 'M082', kategori: 'Sos',          ad: 'Kayısı Reçeli',               birim: 'kg',    fiyat: 200 },
  { id: 'M083', kategori: 'Sos',          ad: 'İncir Reçeli',                birim: 'kg',    fiyat: 250 },
  { id: 'M084', kategori: 'Sos',          ad: 'Domates Salçası',             birim: 'kg',    fiyat: 90 },
  { id: 'M085', kategori: 'İçecek',       ad: 'Beyaz Şarap (mutfak)',        birim: 'lt',    fiyat: 250 },
  { id: 'M086', kategori: 'İçecek',       ad: 'Kırmızı Şarap (mutfak)',      birim: 'lt',    fiyat: 280 },
  { id: 'M087', kategori: 'İçecek',       ad: 'Baileys Likör',               birim: 'lt',    fiyat: 850 },
  { id: 'M088', kategori: 'Sos',          ad: 'Soya Sosu',                   birim: 'lt',    fiyat: 120 },
  { id: 'M089', kategori: 'Sos',          ad: 'Elma Sirkesi',                birim: 'lt',    fiyat: 60 },
  { id: 'M090', kategori: 'Sos',          ad: 'Balsamik Sirke',              birim: 'lt',    fiyat: 200 },
  // KATEGORİ: BAHARAT
  { id: 'M091', kategori: 'Baharat',      ad: 'Tuz',                         birim: 'kg',    fiyat: 15 },
  { id: 'M092', kategori: 'Baharat',      ad: 'Karabiber',                   birim: 'kg',    fiyat: 600 },
  { id: 'M093', kategori: 'Baharat',      ad: 'Kekik (kuru)',                birim: 'kg',    fiyat: 250 },
  { id: 'M094', kategori: 'Baharat',      ad: 'Biberiye',                    birim: 'kg',    fiyat: 300 },
  { id: 'M095', kategori: 'Baharat',      ad: 'Kimyon',                      birim: 'kg',    fiyat: 200 },
  { id: 'M096', kategori: 'Baharat',      ad: 'Tarçın',                      birim: 'kg',    fiyat: 350 },
  { id: 'M097', kategori: 'Baharat',      ad: 'Karanfil',                    birim: 'kg',    fiyat: 800 },
  { id: 'M098', kategori: 'Baharat',      ad: 'Pul Biber',                   birim: 'kg',    fiyat: 250 },
  { id: 'M099', kategori: 'Baharat',      ad: 'Paprika (tatlı)',             birim: 'kg',    fiyat: 250 },
  { id: 'M100', kategori: 'Baharat',      ad: 'Zencefil (taze)',             birim: 'kg',    fiyat: 200 },
  // KATEGORİ: EKMEK / UNLU
  { id: 'M101', kategori: 'Unlu',         ad: 'Yufka (sigara böreği)',       birim: 'kg',    fiyat: 120 },
  { id: 'M102', kategori: 'Unlu',         ad: 'Lavaş',                       birim: 'adet',  fiyat: 8 },
  { id: 'M103', kategori: 'Unlu',         ad: 'Bazlama',                     birim: 'adet',  fiyat: 10 },
  { id: 'M104', kategori: 'Unlu',         ad: 'Ekmek (somun)',               birim: 'adet',  fiyat: 15 },
  { id: 'M105', kategori: 'Unlu',         ad: 'Simit',                       birim: 'adet',  fiyat: 12 },
  { id: 'M106', kategori: 'Unlu',         ad: 'Ekşi Mayalı Ekmek',           birim: 'adet',  fiyat: 25 },
  // KATEGORİ: GIDA TAKVİYESİ
  { id: 'M107', kategori: 'Diğer',        ad: 'Yumurta',                     birim: 'adet',  fiyat: 6 },
  { id: 'M108', kategori: 'Diğer',        ad: 'Şeker',                       birim: 'kg',    fiyat: 35 },
  { id: 'M109', kategori: 'İçecek',       ad: 'Çay (siyah)',                 birim: 'kg',    fiyat: 350 },
  { id: 'M110', kategori: 'İçecek',       ad: 'Türk Kahvesi',                birim: 'kg',    fiyat: 800 },
  { id: 'M111', kategori: 'İçecek',       ad: 'Filtre Kahve',                birim: 'kg',    fiyat: 700 },
  { id: 'M112', kategori: 'Diğer',        ad: 'Tavuk Suyu (bulyon)',         birim: 'kg',    fiyat: 250 },
];

// Yardımcı: malzemenin bizim listemizdeki sıra numarası (1-tabanlı satır)
// Master listede başlık satırı + 1 boş + asıl başlık olduğundan veri 4. satırdan başlar.
const ING_MAP = {};
INGREDIENTS.forEach((ing, i) => { ING_MAP[ing.id] = i + 4; }); // veri başlangıç satırı = 4

// ============================================================
// 2) MENÜ TARİFLERİ — KIŞI BAŞI MİKTARLAR
// Birim her zaman master listedeki birim ile UYUMLU:
// kg→g cinsinden (gram), lt→ml cinsinden, adet/demet→ondalık adet
// Excel'de bu birimleri base unit'e çevirmek için her satırda "çevirme" yapılır.
// ============================================================

// Birim çevirim faktörü: gram→kg=1/1000, ml→lt=1/1000, adet→adet=1, demet→demet=1
function unitFactor(birim) {
  if (birim === 'kg' || birim === 'lt') return 0.001;
  return 1; // adet, demet
}

// Yardımcı: bir tarif satırı ekle
// kullanım: r('M021', 5, 'g')  → 5 gram sarımsak
function r(id, qty, displayUnit) {
  const ing = INGREDIENTS.find(x => x.id === id);
  if (!ing) throw new Error(`Bilinmeyen malzeme: ${id}`);
  return { id, ing, qty, displayUnit };
}

const MENUS = {
  'Turkish Night': {
    fiyat_gbp: 40,
    yemekler: [
      {
        ad: 'Özel Soslu Kuzu But',
        bolum: 'Ana Yemek',
        items: [
          r('M001', 200, 'g'),  r('M023', 150, 'g'),  r('M030', 80, 'g'),
          r('M065', 15, 'g'),   r('M074', 10, 'ml'),  r('M091', 1, 'g'),
          r('M092', 0.5, 'g'),  r('M093', 1, 'g'),    r('M094', 0.5, 'g'),
          r('M021', 5, 'g'),    r('M084', 10, 'g'),   r('M020', 30, 'g'),
          r('M086', 20, 'ml'),
        ],
      },
      {
        ad: 'İç Pilav',
        bolum: 'Ana Yemek',
        items: [
          r('M051', 80, 'g'),   r('M056', 15, 'g'),   r('M059', 10, 'g'),
          r('M065', 15, 'g'),   r('M096', 1, 'g'),    r('M097', 0.5, 'g'),
          r('M040', 0.1, 'demet'), r('M091', 1, 'g'), r('M092', 0.5, 'g'),
          r('M020', 20, 'g'),   r('M112', 5, 'g'),
        ],
      },
      {
        ad: 'Humuslu Karnabahar',
        bolum: 'Meze',
        items: [
          r('M031', 100, 'g'),  r('M053', 40, 'g'),   r('M076', 15, 'g'),
          r('M074', 10, 'ml'),  r('M043', 10, 'g'),   r('M021', 3, 'g'),
          r('M095', 0.5, 'g'),  r('M091', 1, 'g'),    r('M038', 0.1, 'demet'),
        ],
      },
      {
        ad: 'Pembe Domates Sarımsaklı Semizotu',
        bolum: 'Meze',
        items: [
          r('M037', 60, 'g'),   r('M018', 50, 'g'),   r('M021', 5, 'g'),
          r('M074', 10, 'ml'),  r('M043', 5, 'g'),    r('M091', 1, 'g'),
        ],
      },
      {
        ad: 'Yoğurtlu Fırın Sarımsak',
        bolum: 'Meze',
        items: [
          r('M021', 30, 'g'),   r('M062', 50, 'g'),   r('M074', 10, 'ml'),
          r('M091', 1, 'g'),    r('M038', 0.05, 'demet'),
        ],
      },
      {
        ad: 'Zeytinyağlı Yeşil Elmalı Taze Fasulye',
        bolum: 'Meze',
        items: [
          r('M032', 100, 'g'),  r('M033', 40, 'g'),   r('M074', 15, 'ml'),
          r('M020', 15, 'g'),   r('M016', 20, 'g'),   r('M091', 1, 'g'),
          r('M043', 5, 'g'),
        ],
      },
      {
        ad: 'Tulum Peynirli Karpuz Roka Salatası',
        bolum: 'Salata',
        items: [
          r('M044', 100, 'g'),  r('M035', 30, 'g'),   r('M071', 40, 'g'),
          r('M074', 10, 'ml'),  r('M043', 5, 'g'),    r('M091', 1, 'g'),
          r('M092', 0.3, 'g'),  r('M077', 5, 'g'),
        ],
      },
      {
        ad: 'Dondurmalı Fırın Şeftali Tatlısı',
        bolum: 'Tatlı',
        items: [
          r('M046', 150, 'g'),  r('M073', 60, 'g'),   r('M108', 15, 'g'),
          r('M065', 10, 'g'),   r('M096', 0.5, 'g'),
        ],
      },
    ],
  },

  'Mediterranean Night': {
    fiyat_gbp: 45,
    yemekler: [
      {
        ad: 'Akdeniz Levreği',
        bolum: 'Ana Yemek',
        items: [
          r('M010', 180, 'g'),  r('M017', 50, 'g'),   r('M080', 20, 'g'),
          r('M074', 15, 'ml'),  r('M043', 10, 'g'),   r('M021', 5, 'g'),
          r('M038', 0.1, 'demet'), r('M091', 1, 'g'), r('M092', 0.3, 'g'),
          r('M085', 15, 'ml'),
        ],
      },
      {
        ad: 'Hünkar Beğendi',
        bolum: 'Ana Yemek',
        items: [
          r('M003', 120, 'g'),  r('M024', 200, 'g'),  r('M016', 60, 'g'),
          r('M020', 30, 'g'),   r('M055', 10, 'g'),   r('M061', 50, 'ml'),
          r('M070', 20, 'g'),   r('M065', 15, 'g'),   r('M074', 10, 'ml'),
          r('M084', 10, 'g'),   r('M021', 3, 'g'),    r('M091', 1, 'g'),
          r('M092', 0.5, 'g'),
        ],
      },
      {
        ad: 'Adaçaylı Fırın Sebzeler',
        bolum: 'Ana Yemek',
        items: [
          r('M025', 60, 'g'),   r('M026', 50, 'g'),   r('M022', 60, 'g'),
          r('M020', 30, 'g'),   r('M028', 40, 'g'),   r('M074', 15, 'ml'),
          r('M041', 0.1, 'demet'), r('M091', 1, 'g'), r('M092', 0.3, 'g'),
          r('M093', 0.5, 'g'),
        ],
      },
      {
        ad: 'Saganaki (Deniz Mahsulleri)',
        bolum: 'Meze',
        items: [
          r('M012', 40, 'g'),   r('M013', 40, 'g'),   r('M014', 30, 'g'),
          r('M016', 30, 'g'),   r('M070', 30, 'g'),   r('M074', 10, 'ml'),
          r('M021', 3, 'g'),    r('M038', 0.05, 'demet'),
        ],
      },
      {
        ad: 'Humuslu Karnabahar',
        bolum: 'Meze',
        items: [
          r('M031', 100, 'g'),  r('M053', 40, 'g'),   r('M076', 15, 'g'),
          r('M074', 10, 'ml'),  r('M043', 10, 'g'),   r('M021', 3, 'g'),
          r('M095', 0.5, 'g'),  r('M091', 1, 'g'),
        ],
      },
      {
        ad: 'Zeytinyağlı Fırın Dolma Biber',
        bolum: 'Meze',
        items: [
          r('M027', 80, 'g'),   r('M051', 30, 'g'),   r('M020', 15, 'g'),
          r('M016', 15, 'g'),   r('M074', 15, 'ml'),  r('M040', 0.1, 'demet'),
          r('M038', 0.05, 'demet'),
        ],
      },
      {
        ad: 'Peynirli Çıtır Mücver + Füme Somon',
        bolum: 'Meze',
        items: [
          r('M025', 60, 'g'),   r('M055', 15, 'g'),   r('M107', 0.5, 'adet'),
          r('M067', 20, 'g'),   r('M038', 0.1, 'demet'), r('M039', 0.1, 'demet'),
          r('M075', 10, 'ml'),  r('M015', 30, 'g'),
        ],
      },
      {
        ad: 'Aşk Mezesi',
        bolum: 'Meze',
        items: [
          r('M033', 50, 'g'),   r('M048', 30, 'g'),   r('M062', 40, 'g'),
          r('M077', 5, 'g'),    r('M040', 0.05, 'demet'),
        ],
      },
      {
        ad: 'Izgara Sebzeler',
        bolum: 'Meze',
        items: [
          r('M025', 40, 'g'),   r('M026', 40, 'g'),   r('M028', 40, 'g'),
          r('M029', 30, 'g'),   r('M074', 10, 'ml'),  r('M091', 1, 'g'),
          r('M092', 0.3, 'g'),
        ],
      },
      {
        ad: 'Mevsim Yeşillikleri Salatası',
        bolum: 'Salata',
        items: [
          r('M036', 50, 'g'),   r('M058', 10, 'g'),   r('M050', 20, 'g'),
          r('M074', 10, 'ml'),  r('M090', 5, 'ml'),   r('M091', 1, 'g'),
        ],
      },
      {
        ad: 'İncir Çiçeği Tatlısı',
        bolum: 'Tatlı',
        items: [
          r('M083', 40, 'g'),   r('M060', 30, 'g'),   r('M087', 15, 'ml'),
          r('M056', 15, 'g'),   r('M108', 10, 'g'),   r('M066', 30, 'ml'),
        ],
      },
    ],
  },

  'BBQ Menu': {
    fiyat_gbp: 35,
    yemekler: [
      {
        ad: 'Özel Marineli Tavuk Pizola',
        bolum: 'Izgara',
        items: [
          r('M008', 180, 'g'),  r('M074', 10, 'ml'),  r('M091', 1, 'g'),
          r('M093', 0.5, 'g'),  r('M099', 0.5, 'g'),  r('M021', 5, 'g'),
          r('M043', 5, 'g'),
        ],
      },
      {
        ad: 'Kuzu Şiş',
        bolum: 'Izgara',
        items: [
          r('M002', 180, 'g'),  r('M020', 20, 'g'),   r('M074', 10, 'ml'),
          r('M093', 0.5, 'g'),  r('M094', 0.5, 'g'),  r('M021', 5, 'g'),
          r('M028', 10, 'g'),   r('M091', 1, 'g'),
        ],
      },
      {
        ad: 'Köfte (Izgara)',
        bolum: 'Izgara',
        items: [
          r('M004', 150, 'g'),  r('M020', 20, 'g'),   r('M038', 0.1, 'demet'),
          r('M104', 0.05, 'adet'), r('M107', 0.3, 'adet'),
          r('M091', 1, 'g'),    r('M092', 0.5, 'g'),  r('M095', 0.5, 'g'),
        ],
      },
      {
        ad: 'Yeşil Elmalı Yoğurtlu Semizotu',
        bolum: 'Meze',
        items: [
          r('M037', 60, 'g'),   r('M033', 30, 'g'),   r('M062', 40, 'g'),
          r('M074', 5, 'ml'),   r('M091', 1, 'g'),    r('M021', 3, 'g'),
        ],
      },
      {
        ad: 'Humuslu Karnabahar',
        bolum: 'Meze',
        items: [
          r('M031', 80, 'g'),   r('M053', 30, 'g'),   r('M076', 10, 'g'),
          r('M074', 8, 'ml'),   r('M043', 5, 'g'),    r('M021', 3, 'g'),
        ],
      },
      {
        ad: 'Havuç Tarator',
        bolum: 'Meze',
        items: [
          r('M026', 80, 'g'),   r('M057', 15, 'g'),   r('M021', 3, 'g'),
          r('M062', 30, 'g'),   r('M074', 10, 'ml'),  r('M091', 1, 'g'),
        ],
      },
      {
        ad: 'Mücver',
        bolum: 'Meze',
        items: [
          r('M025', 60, 'g'),   r('M026', 20, 'g'),   r('M055', 10, 'g'),
          r('M107', 0.3, 'adet'), r('M038', 0.1, 'demet'), r('M039', 0.05, 'demet'),
          r('M075', 10, 'ml'),
        ],
      },
      {
        ad: 'Piyaz',
        bolum: 'Meze',
        items: [
          r('M054', 40, 'g'),   r('M020', 20, 'g'),   r('M038', 0.1, 'demet'),
          r('M074', 10, 'ml'),  r('M043', 5, 'g'),    r('M091', 1, 'g'),
        ],
      },
      {
        ad: 'Tereyağlı Türk Pilavı',
        bolum: 'Yan Lezzet',
        items: [
          r('M051', 60, 'g'),   r('M052', 10, 'g'),   r('M065', 15, 'g'),
          r('M091', 1, 'g'),    r('M112', 5, 'g'),
        ],
      },
      {
        ad: 'Baharatlı Fırın Sebzeler / Sade Patates',
        bolum: 'Yan Lezzet',
        items: [
          r('M022', 120, 'g'),  r('M074', 10, 'ml'),  r('M091', 1, 'g'),
          r('M093', 0.5, 'g'),  r('M094', 0.5, 'g'),
        ],
      },
      {
        ad: 'Tereyağlı Karides',
        bolum: 'Yan Lezzet',
        items: [
          r('M012', 60, 'g'),   r('M065', 15, 'g'),   r('M021', 3, 'g'),
          r('M038', 0.05, 'demet'), r('M043', 5, 'g'),
        ],
      },
      {
        ad: 'Mevsim Yeşillikleri Salatası',
        bolum: 'Salata',
        items: [
          r('M036', 40, 'g'),   r('M074', 8, 'ml'),   r('M043', 5, 'g'),
          r('M091', 1, 'g'),
        ],
      },
      {
        ad: 'Şefin Özel Tatlısı (porsiyon bütçesi)',
        bolum: 'Tatlı',
        items: [
          r('M061', 80, 'ml'),  r('M051', 20, 'g'),   r('M108', 25, 'g'),
          r('M096', 0.5, 'g'),  r('M064', 5, 'g'),
        ],
      },
    ],
  },

  'Chicken BBQ': {
    fiyat_gbp: 26,
    yemekler: [
      {
        ad: 'Özel Marineli Tavuk But Fileto',
        bolum: 'Izgara Tavuk',
        items: [
          r('M005', 180, 'g'),  r('M074', 10, 'ml'),  r('M021', 5, 'g'),
          r('M093', 0.5, 'g'),  r('M099', 0.5, 'g'),  r('M043', 5, 'g'),
          r('M091', 1, 'g'),    r('M092', 0.3, 'g'),
        ],
      },
      {
        ad: 'Marineli Tavuk Kanatları',
        bolum: 'Izgara Tavuk',
        items: [
          r('M007', 150, 'g'),  r('M074', 8, 'ml'),   r('M021', 5, 'g'),
          r('M043', 5, 'g'),    r('M091', 1, 'g'),    r('M098', 0.3, 'g'),
          r('M099', 0.3, 'g'),
        ],
      },
      {
        ad: 'Renkli Biberli Tavuk Şiş',
        bolum: 'Izgara Tavuk',
        items: [
          r('M006', 150, 'g'),  r('M028', 90, 'g'),   r('M020', 30, 'g'),
          r('M074', 8, 'ml'),   r('M098', 0.5, 'g'),  r('M021', 3, 'g'),
        ],
      },
      {
        ad: 'Mücver',
        bolum: 'Meze',
        items: [
          r('M025', 60, 'g'),   r('M026', 20, 'g'),   r('M022', 20, 'g'),
          r('M055', 10, 'g'),   r('M107', 0.3, 'adet'),
          r('M038', 0.1, 'demet'), r('M039', 0.05, 'demet'), r('M075', 10, 'ml'),
        ],
      },
      {
        ad: 'Babagannuş',
        bolum: 'Meze',
        items: [
          r('M024', 120, 'g'),  r('M076', 10, 'g'),   r('M021', 3, 'g'),
          r('M043', 5, 'g'),    r('M074', 5, 'ml'),   r('M091', 1, 'g'),
        ],
      },
      {
        ad: 'Peri Peri Sos',
        bolum: 'Sos',
        items: [
          r('M028', 10, 'g'),   r('M021', 2, 'g'),    r('M043', 3, 'g'),
          r('M074', 5, 'ml'),   r('M098', 1, 'g'),    r('M089', 3, 'ml'),
          r('M091', 0.5, 'g'),
        ],
      },
      {
        ad: 'Taze Otlu Sarımsaklı Sos',
        bolum: 'Sos',
        items: [
          r('M062', 40, 'g'),   r('M038', 0.05, 'demet'), r('M039', 0.05, 'demet'),
          r('M021', 3, 'g'),    r('M091', 0.5, 'g'),
        ],
      },
      {
        ad: 'Barbekü Sosu (ev yapımı)',
        bolum: 'Sos',
        items: [
          r('M084', 15, 'g'),   r('M077', 5, 'g'),    r('M074', 3, 'ml'),
          r('M089', 3, 'ml'),   r('M098', 0.3, 'g'),  r('M091', 0.3, 'g'),
        ],
      },
      {
        ad: 'Avokado Soslu Yeşil Salata',
        bolum: 'Salata',
        items: [
          r('M036', 50, 'g'),   r('M034', 0.4, 'adet'), r('M074', 8, 'ml'),
          r('M043', 5, 'g'),    r('M091', 0.5, 'g'),
        ],
      },
      {
        ad: 'Türk Ekmekleri',
        bolum: 'Ekmek',
        items: [
          r('M102', 0.4, 'adet'), r('M103', 0.3, 'adet'),
        ],
      },
    ],
  },

  'Fish BBQ': {
    fiyat_gbp: 45,
    yemekler: [
      {
        ad: 'Asma Yaprağında Marineli Levrek',
        bolum: 'Izgara',
        items: [
          r('M010', 180, 'g'),  r('M042', 15, 'g'),   r('M074', 10, 'ml'),
          r('M043', 5, 'g'),    r('M021', 3, 'g'),    r('M091', 1, 'g'),
          r('M093', 0.5, 'g'),
        ],
      },
      {
        ad: 'Marineli Somon',
        bolum: 'Izgara',
        items: [
          r('M011', 150, 'g'),  r('M074', 10, 'ml'),  r('M043', 8, 'g'),
          r('M039', 0.1, 'demet'), r('M091', 1, 'g'), r('M092', 0.3, 'g'),
        ],
      },
      {
        ad: 'Adaçaylı Fırın Sebzeler',
        bolum: 'Izgara',
        items: [
          r('M025', 60, 'g'),   r('M026', 40, 'g'),   r('M022', 50, 'g'),
          r('M074', 10, 'ml'),  r('M041', 0.1, 'demet'), r('M091', 1, 'g'),
        ],
      },
      {
        ad: 'Izgara Zeytin',
        bolum: 'Meze',
        items: [
          r('M079', 40, 'g'),   r('M080', 40, 'g'),   r('M074', 5, 'ml'),
          r('M093', 0.3, 'g'),
        ],
      },
      {
        ad: 'Humuslu Taze Otlu Karnabahar',
        bolum: 'Meze',
        items: [
          r('M031', 80, 'g'),   r('M053', 30, 'g'),   r('M076', 10, 'g'),
          r('M074', 8, 'ml'),   r('M043', 5, 'g'),    r('M038', 0.05, 'demet'),
          r('M091', 0.5, 'g'),
        ],
      },
      {
        ad: 'Yeşil Elmalı Çilekli Süzme Yoğurt',
        bolum: 'Meze',
        items: [
          r('M063', 60, 'g'),   r('M033', 30, 'g'),   r('M048', 30, 'g'),
          r('M077', 5, 'g'),    r('M040', 0.05, 'demet'),
        ],
      },
      {
        ad: 'Taze Otlu Sarımsaklı Sos',
        bolum: 'Sos',
        items: [
          r('M062', 30, 'g'),   r('M038', 0.05, 'demet'), r('M039', 0.05, 'demet'),
          r('M021', 3, 'g'),    r('M091', 0.5, 'g'),
        ],
      },
      {
        ad: 'Limonlu Tereyağı Sosu',
        bolum: 'Sos',
        items: [
          r('M065', 20, 'g'),   r('M043', 8, 'g'),    r('M038', 0.05, 'demet'),
          r('M091', 0.5, 'g'),
        ],
      },
      {
        ad: 'Acı Tatlı Özel Sos',
        bolum: 'Sos',
        items: [
          r('M077', 10, 'g'),   r('M098', 0.5, 'g'),  r('M088', 5, 'ml'),
          r('M100', 2, 'g'),    r('M021', 2, 'g'),
        ],
      },
      {
        ad: 'Roka Tulum Salatası',
        bolum: 'Salata',
        items: [
          r('M035', 40, 'g'),   r('M071', 30, 'g'),   r('M074', 10, 'ml'),
          r('M043', 5, 'g'),    r('M091', 0.5, 'g'),
        ],
      },
      {
        ad: 'Türk Ekmekleri',
        bolum: 'Ekmek',
        items: [
          r('M102', 0.4, 'adet'), r('M103', 0.3, 'adet'), r('M106', 0.2, 'adet'),
        ],
      },
    ],
  },

  'Breakfast': {
    fiyat_gbp: 20,
    yemekler: [
      {
        ad: 'Sucuklu Yumurta',
        bolum: 'Sıcak',
        items: [
          r('M107', 2, 'adet'), r('M009', 40, 'g'),   r('M065', 5, 'g'),
        ],
      },
      {
        ad: 'Patates Kızartması',
        bolum: 'Sıcak',
        items: [
          r('M022', 100, 'g'),  r('M075', 15, 'ml'),  r('M091', 0.5, 'g'),
        ],
      },
      {
        ad: 'Beyaz Peynirli Sigara Böreği',
        bolum: 'Sıcak',
        items: [
          r('M101', 40, 'g'),   r('M067', 30, 'g'),   r('M038', 0.05, 'demet'),
          r('M107', 0.2, 'adet'), r('M075', 10, 'ml'),
        ],
      },
      {
        ad: 'Şarküteri Tabağı (sebze)',
        bolum: 'Sofra',
        items: [
          r('M016', 40, 'g'),   r('M019', 40, 'g'),   r('M035', 10, 'g'),
          r('M038', 0.05, 'demet'), r('M043', 10, 'g'),
        ],
      },
      {
        ad: 'Dört Çeşit Peynir',
        bolum: 'Sofra',
        items: [
          r('M067', 15, 'g'),   r('M068', 15, 'g'),   r('M069', 15, 'g'),
          r('M071', 15, 'g'),
        ],
      },
      {
        ad: 'Yeşil ve Siyah Zeytin',
        bolum: 'Sofra',
        items: [
          r('M079', 15, 'g'),   r('M080', 15, 'g'),   r('M074', 2, 'ml'),
          r('M093', 0.2, 'g'),
        ],
      },
      {
        ad: 'Reçel · Tahin · Pekmez · Bal',
        bolum: 'Sofra',
        items: [
          r('M081', 20, 'g'),   r('M082', 20, 'g'),   r('M083', 20, 'g'),
          r('M076', 15, 'g'),   r('M078', 15, 'g'),   r('M077', 15, 'g'),
        ],
      },
      {
        ad: 'Kaymak ve Yoğurt',
        bolum: 'Sofra',
        items: [
          r('M064', 30, 'g'),   r('M063', 50, 'g'),
        ],
      },
      {
        ad: 'Mevsim Meyveleri',
        bolum: 'Sofra',
        items: [
          r('M044', 60, 'g'),   r('M045', 50, 'g'),   r('M046', 50, 'g'),
          r('M047', 40, 'g'),
        ],
      },
      {
        ad: 'Mevsim Ekmeği ve Simit',
        bolum: 'Ekmek',
        items: [
          r('M105', 1, 'adet'), r('M104', 0.3, 'adet'),
        ],
      },
      {
        ad: 'Portakal Suyu (taze sıkma)',
        bolum: 'İçecek',
        items: [
          r('M049', 200, 'g'),
        ],
      },
      {
        ad: 'Kahve ve Çay',
        bolum: 'İçecek',
        items: [
          r('M109', 3, 'g'),    r('M110', 8, 'g'),    r('M108', 5, 'g'),
        ],
      },
    ],
  },
};

// ============================================================
// 3) STİLLER
// ============================================================
const STYLE = {
  brand: { argb: 'FF15342B' },          // forest
  brand2: { argb: 'FF22483D' },         // forest3
  gold: { argb: 'FFD9923B' },           // sun
  cream: { argb: 'FFF2EBDA' },
  sand: { argb: 'FFE5D9BE' },
  light: { argb: 'FFFAF6EE' },
  border: { argb: 'FFC0B294' },
  totalRow: { argb: 'FFFFF4D6' },
  grandTotal: { argb: 'FFE5D9BE' },
};

function setBorder(cell, color = STYLE.border) {
  cell.border = {
    top:    { style: 'thin', color },
    bottom: { style: 'thin', color },
    left:   { style: 'thin', color },
    right:  { style: 'thin', color },
  };
}

function styleHeader(cell) {
  cell.font = { name: 'Calibri', size: 11, bold: true, color: STYLE.cream };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: STYLE.brand };
  cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  setBorder(cell);
}

function styleSection(cell) {
  cell.font = { name: 'Calibri', size: 11, bold: true, color: STYLE.brand };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: STYLE.sand };
  cell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
  setBorder(cell);
}

function styleData(cell, opts = {}) {
  cell.font = { name: 'Calibri', size: 10 };
  cell.alignment = { vertical: 'middle', horizontal: opts.align || 'left' };
  setBorder(cell);
  if (opts.numFmt) cell.numFmt = opts.numFmt;
  if (opts.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: opts.fill };
}

function styleTitle(cell) {
  cell.font = { name: 'Calibri', size: 18, bold: true, color: STYLE.gold };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: STYLE.brand };
  cell.alignment = { vertical: 'middle', horizontal: 'center' };
}

// ============================================================
// 4) ANA SAYFA: MALZEME FİYAT LİSTESİ
// ============================================================
async function buildPriceSheet(ws) {
  ws.columns = [
    { width: 8 },   // A: ID
    { width: 18 },  // B: Kategori
    { width: 38 },  // C: Malzeme
    { width: 10 },  // D: Birim
    { width: 16 },  // E: Birim Fiyat (TL)
    { width: 22 },  // F: Tedarikçi
    { width: 16 },  // G: Son Güncelleme
    { width: 28 },  // H: Notlar
  ];

  // Başlık
  ws.mergeCells('A1:H1');
  styleTitle(ws.getCell('A1'));
  ws.getCell('A1').value = 'KALİMERA KITCHEN — MALZEME FİYAT LİSTESİ';
  ws.getRow(1).height = 36;

  // Açıklama
  ws.mergeCells('A2:H2');
  ws.getCell('A2').value = 'Tek Fiyat Kaynağı  ·  Bu sayfadaki E sütununu güncellersen, tüm menü maliyetleri OTOMATİK güncellenir.';
  ws.getCell('A2').font = { italic: true, color: STYLE.brand, size: 10 };
  ws.getCell('A2').alignment = { horizontal: 'center' };
  ws.getRow(2).height = 22;

  // Sütun başlıkları (3. satır)
  const headers = ['ID', 'Kategori', 'Malzeme Adı', 'Birim', 'Birim Fiyat (TL)', 'Tedarikçi', 'Son Güncelleme', 'Notlar'];
  headers.forEach((h, i) => {
    const cell = ws.getCell(3, i + 1);
    cell.value = h;
    styleHeader(cell);
  });
  ws.getRow(3).height = 24;

  // Veri (satır 4'ten başlar)
  INGREDIENTS.forEach((ing, i) => {
    const row = i + 4;
    ws.getCell(row, 1).value = ing.id;
    ws.getCell(row, 2).value = ing.kategori;
    ws.getCell(row, 3).value = ing.ad;
    ws.getCell(row, 4).value = ing.birim;
    ws.getCell(row, 5).value = ing.fiyat;
    ws.getCell(row, 6).value = '';
    ws.getCell(row, 7).value = '';
    ws.getCell(row, 8).value = '';

    styleData(ws.getCell(row, 1), { align: 'center' });
    styleData(ws.getCell(row, 2), { align: 'left' });
    styleData(ws.getCell(row, 3), { align: 'left' });
    styleData(ws.getCell(row, 4), { align: 'center' });
    styleData(ws.getCell(row, 5), { align: 'right', numFmt: '#,##0.00 ₺', fill: STYLE.totalRow });
    styleData(ws.getCell(row, 6), { align: 'left' });
    styleData(ws.getCell(row, 7), { align: 'center', numFmt: 'dd/mm/yyyy' });
    styleData(ws.getCell(row, 8), { align: 'left' });

    // Renk: kategoriye göre arkaplan rengi (3. sütun)
    const catRow = ws.getCell(row, 2);
    catRow.font = { size: 10, bold: true, color: STYLE.brand };
  });

  // Donmuş başlık
  ws.views = [{ state: 'frozen', ySplit: 3 }];

  // Otomatik filtre
  ws.autoFilter = {
    from: { row: 3, column: 1 },
    to: { row: 3 + INGREDIENTS.length, column: 8 },
  };
}

// ============================================================
// 5) MENÜ SAYFASI ÜRETİCİSİ
// ============================================================
async function buildMenuSheet(ws, menuName, menuData) {
  ws.columns = [
    { width: 28 },  // A: Yemek
    { width: 32 },  // B: Malzeme
    { width: 14 },  // C: Miktar (görüntü)
    { width: 8 },   // D: Birim (görüntü)
    { width: 10 },  // E: Master Birim
    { width: 16 },  // F: Birim Fiyat (TL)
    { width: 14 },  // G: 1 kişi (TL)
    { width: 14 },  // H: 6 kişi (TL)
    { width: 14 },  // I: 8 kişi (TL)
    { width: 14 },  // J: 10 kişi (TL)
  ];

  // Başlık
  ws.mergeCells('A1:J1');
  styleTitle(ws.getCell('A1'));
  ws.getCell('A1').value = `${menuName.toUpperCase()} — Maliyet Hesabı`;
  ws.getRow(1).height = 36;

  // Üst bilgi: Satış fiyatı + GBP→TL kuru
  ws.mergeCells('A2:F2');
  ws.getCell('A2').value = `Satış Fiyatı: ${menuData.fiyat_gbp} £ / kişi`;
  ws.getCell('A2').font = { italic: true, bold: true, color: STYLE.brand, size: 11 };
  ws.getCell('A2').alignment = { horizontal: 'left', indent: 1 };

  ws.getCell('G2').value = 'GBP→TL Kuru:';
  ws.getCell('G2').font = { italic: true, bold: true, color: STYLE.brand, size: 10 };
  ws.getCell('G2').alignment = { horizontal: 'right' };
  ws.getCell('H2').value = 50; // Varsayılan kur, değiştirilebilir
  ws.getCell('H2').font = { bold: true, color: { argb: 'FFB35900' }, size: 11 };
  ws.getCell('H2').numFmt = '#,##0.00 ₺';
  ws.getCell('H2').alignment = { horizontal: 'center' };
  ws.getCell('H2').fill = { type: 'pattern', pattern: 'solid', fgColor: STYLE.totalRow };
  setBorder(ws.getCell('H2'));
  ws.getRow(2).height = 22;

  // Sütun başlıkları
  const headers = ['Yemek / Bölüm', 'Malzeme', 'Miktar', 'Birim', 'Master Birim', 'Birim Fiyat (TL)', '1 Kişi (TL)', '6 Kişi (TL)', '8 Kişi (TL)', '10 Kişi (TL)'];
  headers.forEach((h, i) => {
    const cell = ws.getCell(3, i + 1);
    cell.value = h;
    styleHeader(cell);
  });
  ws.getRow(3).height = 36;

  // Veri satırları
  let row = 4;
  const dishTotalRows = []; // her yemeğin toplam satırını sakla
  const sectionMap = {}; // bölüm adı → satır listesi

  for (const yemek of menuData.yemekler) {
    // Yemek başlığı
    ws.mergeCells(row, 1, row, 10);
    const titleCell = ws.getCell(row, 1);
    titleCell.value = `${yemek.bolum} · ${yemek.ad}`;
    styleSection(titleCell);
    ws.getRow(row).height = 22;
    row++;

    const dishStartRow = row;
    for (const item of yemek.items) {
      const masterRow = ING_MAP[item.id];
      const factor = unitFactor(item.ing.birim);
      // Birim fiyat (Master sayfadan formülle)
      const priceFormula = `'Malzeme Fiyat Listesi'!E${masterRow}`;

      ws.getCell(row, 1).value = '';
      ws.getCell(row, 2).value = item.ing.ad;
      ws.getCell(row, 3).value = item.qty;
      ws.getCell(row, 4).value = item.displayUnit;
      ws.getCell(row, 5).value = item.ing.birim;
      // Birim fiyat formül
      ws.getCell(row, 6).value = { formula: priceFormula };
      // 1 kişi (TL) = qty * factor * fiyat
      // qty C{row}, factor sabit, fiyat F{row}
      ws.getCell(row, 7).value = { formula: `C${row}*${factor}*F${row}` };
      ws.getCell(row, 8).value = { formula: `G${row}*6` };
      ws.getCell(row, 9).value = { formula: `G${row}*8` };
      ws.getCell(row, 10).value = { formula: `G${row}*10` };

      styleData(ws.getCell(row, 1));
      styleData(ws.getCell(row, 2), { align: 'left' });
      styleData(ws.getCell(row, 3), { align: 'right', numFmt: '0.###' });
      styleData(ws.getCell(row, 4), { align: 'center' });
      styleData(ws.getCell(row, 5), { align: 'center' });
      styleData(ws.getCell(row, 6), { align: 'right', numFmt: '#,##0.00 ₺' });
      styleData(ws.getCell(row, 7), { align: 'right', numFmt: '#,##0.00 ₺' });
      styleData(ws.getCell(row, 8), { align: 'right', numFmt: '#,##0.00 ₺' });
      styleData(ws.getCell(row, 9), { align: 'right', numFmt: '#,##0.00 ₺' });
      styleData(ws.getCell(row, 10), { align: 'right', numFmt: '#,##0.00 ₺' });

      row++;
    }
    // Yemek toplamı satırı
    ws.mergeCells(row, 1, row, 6);
    const sumLabel = ws.getCell(row, 1);
    sumLabel.value = `${yemek.ad} — Yemek Toplamı`;
    sumLabel.font = { bold: true, italic: true, color: STYLE.brand, size: 10 };
    sumLabel.alignment = { horizontal: 'right', indent: 1 };
    sumLabel.fill = { type: 'pattern', pattern: 'solid', fgColor: STYLE.totalRow };
    setBorder(sumLabel);
    for (let c = 7; c <= 10; c++) {
      const col = String.fromCharCode(64 + c);
      const cell = ws.getCell(row, c);
      cell.value = { formula: `SUM(${col}${dishStartRow}:${col}${row - 1})` };
      cell.numFmt = '#,##0.00 ₺';
      cell.font = { bold: true, color: STYLE.brand, size: 10 };
      cell.alignment = { horizontal: 'right' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: STYLE.totalRow };
      setBorder(cell);
    }
    dishTotalRows.push(row);
    row++;
    // Boş ayraç
    row++;
  }

  // ============= GENEL TOPLAM BÖLÜMÜ =============
  row++;
  ws.mergeCells(row, 1, row, 10);
  const grandHeader = ws.getCell(row, 1);
  grandHeader.value = 'MENÜ GENEL TOPLAM';
  grandHeader.font = { bold: true, color: STYLE.gold, size: 13 };
  grandHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: STYLE.brand };
  grandHeader.alignment = { horizontal: 'center' };
  setBorder(grandHeader);
  ws.getRow(row).height = 28;
  row++;

  // Toplam maliyet satırı
  ws.mergeCells(row, 1, row, 6);
  const totalLabel = ws.getCell(row, 1);
  totalLabel.value = 'TOPLAM MALİYET (Malzeme)';
  totalLabel.font = { bold: true, color: STYLE.brand, size: 11 };
  totalLabel.fill = { type: 'pattern', pattern: 'solid', fgColor: STYLE.grandTotal };
  totalLabel.alignment = { horizontal: 'right', indent: 1 };
  setBorder(totalLabel);
  const totalRow = row;
  for (let c = 7; c <= 10; c++) {
    const col = String.fromCharCode(64 + c);
    const sumParts = dishTotalRows.map(r => `${col}${r}`).join(',');
    const cell = ws.getCell(row, c);
    cell.value = { formula: `${sumParts ? sumParts.split(',').join('+') : '0'}` };
    cell.numFmt = '#,##0.00 ₺';
    cell.font = { bold: true, color: STYLE.brand, size: 12 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: STYLE.grandTotal };
    cell.alignment = { horizontal: 'right' };
    setBorder(cell);
  }
  ws.getRow(row).height = 24;
  row++;

  // Satış geliri satırı (GBP × kur × kişi)
  ws.mergeCells(row, 1, row, 6);
  const salesLabel = ws.getCell(row, 1);
  salesLabel.value = `SATIŞ GELİRİ (${menuData.fiyat_gbp}£ × Kur × Kişi)`;
  salesLabel.font = { bold: true, color: STYLE.brand, size: 11 };
  salesLabel.fill = { type: 'pattern', pattern: 'solid', fgColor: STYLE.totalRow };
  salesLabel.alignment = { horizontal: 'right', indent: 1 };
  setBorder(salesLabel);
  const persons = [1, 6, 8, 10];
  for (let i = 0; i < 4; i++) {
    const c = 7 + i;
    const cell = ws.getCell(row, c);
    cell.value = { formula: `${menuData.fiyat_gbp}*$H$2*${persons[i]}` };
    cell.numFmt = '#,##0.00 ₺';
    cell.font = { bold: true, color: { argb: 'FF0F5132' }, size: 11 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: STYLE.totalRow };
    cell.alignment = { horizontal: 'right' };
    setBorder(cell);
  }
  const salesRow = row;
  row++;

  // Brüt kâr satırı
  ws.mergeCells(row, 1, row, 6);
  const profitLabel = ws.getCell(row, 1);
  profitLabel.value = 'BRÜT KÂR (Satış − Malzeme)';
  profitLabel.font = { bold: true, color: STYLE.brand, size: 11 };
  profitLabel.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD4F4DD' } };
  profitLabel.alignment = { horizontal: 'right', indent: 1 };
  setBorder(profitLabel);
  for (let c = 7; c <= 10; c++) {
    const col = String.fromCharCode(64 + c);
    const cell = ws.getCell(row, c);
    cell.value = { formula: `${col}${salesRow}-${col}${totalRow}` };
    cell.numFmt = '#,##0.00 ₺';
    cell.font = { bold: true, color: { argb: 'FF0F5132' }, size: 12 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD4F4DD' } };
    cell.alignment = { horizontal: 'right' };
    setBorder(cell);
  }
  row++;

  // Kâr marjı yüzdesi
  ws.mergeCells(row, 1, row, 6);
  const marginLabel = ws.getCell(row, 1);
  marginLabel.value = 'KÂR MARJI (%)';
  marginLabel.font = { bold: true, color: STYLE.brand, size: 11 };
  marginLabel.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD4F4DD' } };
  marginLabel.alignment = { horizontal: 'right', indent: 1 };
  setBorder(marginLabel);
  for (let c = 7; c <= 10; c++) {
    const col = String.fromCharCode(64 + c);
    const cell = ws.getCell(row, c);
    cell.value = { formula: `IFERROR((${col}${salesRow}-${col}${totalRow})/${col}${salesRow},0)` };
    cell.numFmt = '0.0%';
    cell.font = { bold: true, color: { argb: 'FF0F5132' }, size: 12 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD4F4DD' } };
    cell.alignment = { horizontal: 'right' };
    setBorder(cell);
  }
  row++;

  // Donmuş başlık
  ws.views = [{ state: 'frozen', ySplit: 3 }];
}

// ============================================================
// 6) ALIŞVERİŞ ÖZETİ — Menü × Kişi sayısı toplam malzeme
// ============================================================
async function buildShoppingSheet(ws) {
  ws.columns = [
    { width: 8 },   // A: ID
    { width: 38 },  // B: Malzeme
    { width: 10 },  // C: Birim
    { width: 14 },  // D: Birim Fiyat
    { width: 14 },  // E: Toplam (1 kişi-g/ml/adet)
    { width: 14 },  // F: 1 kişi miktar
    { width: 14 },  // G: 6 kişi
    { width: 14 },  // H: 8 kişi
    { width: 14 },  // I: 10 kişi
    { width: 14 },  // J: 10 kişi maliyet (TL)
  ];

  ws.mergeCells('A1:J1');
  styleTitle(ws.getCell('A1'));
  ws.getCell('A1').value = 'MENÜ SEÇ — TOPLAM MALZEME ALIŞVERİŞ LİSTESİ';
  ws.getRow(1).height = 36;

  // Menü seçici
  ws.mergeCells('A2:C2');
  ws.getCell('A2').value = 'Menü Seç:';
  ws.getCell('A2').font = { bold: true, color: STYLE.brand, size: 11 };
  ws.getCell('A2').alignment = { horizontal: 'right' };

  ws.mergeCells('D2:F2');
  ws.getCell('D2').value = 'Turkish Night';
  ws.getCell('D2').font = { bold: true, color: { argb: 'FFB35900' }, size: 12 };
  ws.getCell('D2').alignment = { horizontal: 'center' };
  ws.getCell('D2').fill = { type: 'pattern', pattern: 'solid', fgColor: STYLE.totalRow };
  setBorder(ws.getCell('D2'));

  // Veri doğrulaması: menü adı listesi
  const menuNames = Object.keys(MENUS);
  ws.dataValidations.add('D2:F2', {
    type: 'list',
    allowBlank: false,
    formulae: [`"${menuNames.join(',')}"`],
    showErrorMessage: true,
    errorTitle: 'Geçersiz Menü',
    error: 'Lütfen listeden bir menü seçin.',
  });

  ws.getCell('G2').value = 'Kişi:';
  ws.getCell('G2').font = { bold: true, color: STYLE.brand, size: 11 };
  ws.getCell('G2').alignment = { horizontal: 'right' };

  ws.getCell('H2').value = 8;
  ws.getCell('H2').font = { bold: true, color: { argb: 'FFB35900' }, size: 12 };
  ws.getCell('H2').alignment = { horizontal: 'center' };
  ws.getCell('H2').fill = { type: 'pattern', pattern: 'solid', fgColor: STYLE.totalRow };
  setBorder(ws.getCell('H2'));
  ws.dataValidations.add('H2', {
    type: 'list',
    allowBlank: false,
    formulae: ['"1,6,8,10,12,15,20,25,30,40,50"'],
  });

  ws.mergeCells('I2:J2');
  ws.getCell('I2').value = '← Menü ve kişiyi seçtiğinde liste otomatik dolar';
  ws.getCell('I2').font = { italic: true, color: STYLE.brand, size: 9 };
  ws.getCell('I2').alignment = { horizontal: 'left', indent: 1 };

  ws.getRow(2).height = 24;

  // Sütun başlıkları
  const headers = ['ID', 'Malzeme', 'Birim', 'Birim Fiyat', 'Tek Kişi (gram/ml/adet)', '1 Kişi', '6 Kişi', '8 Kişi', '10 Kişi', 'Seçili Toplam (TL)'];
  headers.forEach((h, i) => {
    const cell = ws.getCell(3, i + 1);
    cell.value = h;
    styleHeader(cell);
  });
  ws.getRow(3).height = 36;

  // Her menü için per-kişi toplam miktarları hesapla
  // Sayısal toplama burada tek-kişi base unit (g/ml/adet/demet) cinsinden yapılır
  const aggregates = {}; // { menuName: { ingId: totalAmount } }
  for (const [menuName, menuData] of Object.entries(MENUS)) {
    const map = {};
    for (const yemek of menuData.yemekler) {
      for (const item of yemek.items) {
        // Display unit: g, ml, adet, demet
        // Master birim: kg, lt, adet, demet
        // qty her zaman base unit cinsinden (g/ml/adet/demet)
        if (!map[item.id]) map[item.id] = 0;
        map[item.id] += item.qty;
      }
    }
    aggregates[menuName] = map;
  }

  // Şimdi listenin tamamı — kullanılan tüm malzemeler
  const usedIds = new Set();
  Object.values(aggregates).forEach(m => Object.keys(m).forEach(id => usedIds.add(id)));
  const sortedIds = INGREDIENTS.filter(i => usedIds.has(i.id)).map(i => i.id);

  // Veri başlangıç satırı
  let row = 4;
  for (const id of sortedIds) {
    const ing = INGREDIENTS.find(x => x.id === id);
    const masterRow = ING_MAP[id];
    const priceFormula = `'Malzeme Fiyat Listesi'!E${masterRow}`;

    // CHOOSE bazlı per-menu sabit dizisi
    // sıralı menüler: aggregates'in key sırası
    const menuKeys = Object.keys(MENUS);
    const perKisiVals = menuKeys.map(m => aggregates[m][id] || 0);

    // CHOOSE(MATCH(menü, {liste}, 0), val1, val2, ...)
    // Excel formülü:
    const matchFormula = `MATCH($D$2,{${menuKeys.map(m => `"${m}"`).join(',')}},0)`;
    const chooseArgs = perKisiVals.map(v => v.toString()).join(',');
    const tekKisiCell = `CHOOSE(${matchFormula},${chooseArgs})`;

    ws.getCell(row, 1).value = id;
    ws.getCell(row, 2).value = ing.ad;
    ws.getCell(row, 3).value = ing.birim === 'kg' ? 'g' : ing.birim === 'lt' ? 'ml' : ing.birim;
    ws.getCell(row, 4).value = { formula: priceFormula };
    ws.getCell(row, 5).value = { formula: tekKisiCell };
    ws.getCell(row, 6).value = { formula: `E${row}*1` };
    ws.getCell(row, 7).value = { formula: `E${row}*6` };
    ws.getCell(row, 8).value = { formula: `E${row}*8` };
    ws.getCell(row, 9).value = { formula: `E${row}*10` };

    // Maliyet: tek_kişi_miktar * H2_kişi * (factor) * birim_fiyat
    const factor = unitFactor(ing.birim);
    ws.getCell(row, 10).value = { formula: `E${row}*$H$2*${factor}*D${row}` };

    styleData(ws.getCell(row, 1), { align: 'center' });
    styleData(ws.getCell(row, 2));
    styleData(ws.getCell(row, 3), { align: 'center' });
    styleData(ws.getCell(row, 4), { align: 'right', numFmt: '#,##0.00 ₺' });
    styleData(ws.getCell(row, 5), { align: 'right', numFmt: '#,##0.###' });
    styleData(ws.getCell(row, 6), { align: 'right', numFmt: '#,##0.###' });
    styleData(ws.getCell(row, 7), { align: 'right', numFmt: '#,##0.###' });
    styleData(ws.getCell(row, 8), { align: 'right', numFmt: '#,##0.###' });
    styleData(ws.getCell(row, 9), { align: 'right', numFmt: '#,##0.###' });
    styleData(ws.getCell(row, 10), { align: 'right', numFmt: '#,##0.00 ₺', fill: STYLE.totalRow });

    row++;
  }

  // Genel toplam satırı
  ws.mergeCells(row, 1, row, 9);
  const totalLabel = ws.getCell(row, 1);
  totalLabel.value = 'SEÇİLİ MENÜ × KİŞİ — TOPLAM MALZEME MALİYETİ';
  totalLabel.font = { bold: true, color: STYLE.gold, size: 12 };
  totalLabel.fill = { type: 'pattern', pattern: 'solid', fgColor: STYLE.brand };
  totalLabel.alignment = { horizontal: 'right', indent: 1 };
  setBorder(totalLabel);
  const totalCell = ws.getCell(row, 10);
  totalCell.value = { formula: `SUM(J4:J${row - 1})` };
  totalCell.numFmt = '#,##0.00 ₺';
  totalCell.font = { bold: true, color: STYLE.gold, size: 13 };
  totalCell.fill = { type: 'pattern', pattern: 'solid', fgColor: STYLE.brand };
  totalCell.alignment = { horizontal: 'right' };
  setBorder(totalCell);
  ws.getRow(row).height = 28;

  ws.views = [{ state: 'frozen', ySplit: 3, xSplit: 2 }];
}

// ============================================================
// 7) GİRİŞ SAYFASI — Kullanım kılavuzu + özet
// ============================================================
async function buildIntroSheet(ws) {
  ws.columns = [{ width: 6 }, { width: 90 }];

  ws.mergeCells('A1:B1');
  styleTitle(ws.getCell('A1'));
  ws.getCell('A1').value = 'KALİMERA KITCHEN — MALİYET HESAP TABLOSU';
  ws.getRow(1).height = 44;

  ws.mergeCells('A2:B2');
  ws.getCell('A2').value = 'Kalkan / Antalya  ·  Premium Catering  ·  Kişi başı maliyet, kâr marjı ve alışveriş listesi';
  ws.getCell('A2').font = { italic: true, color: STYLE.brand, size: 11 };
  ws.getCell('A2').alignment = { horizontal: 'center' };
  ws.getRow(2).height = 22;

  const lines = [
    [3, '', ''],
    [4, 'NASIL KULLANILIR', 'header'],
    [5, '1', '“Malzeme Fiyat Listesi” sayfasını aç. Bu sayfa TEK fiyat kaynağıdır.'],
    [6, '2', 'Bir malzemenin fiyatı değiştiğinde sadece o satırdaki “Birim Fiyat (TL)” hücresini güncelle. Tüm menü maliyetleri OTOMATİK güncellenir.'],
    [7, '3', 'Her menü kendi sayfasında: 1 kişi, 6, 8 ve 10 kişi maliyetleri formüllerle hesaplanır. Satış geliri, brüt kâr ve kâr marjı (%) otomatik gösterilir.'],
    [8, '4', '“Alışveriş Özeti” sayfası: Menü ve kişi sayısı seç → tüm malzemelerin toplam alım miktarı ve toplam maliyeti otomatik gelir.'],
    [9, '5', 'GBP→TL kuru her menü sayfasında H2 hücresinde değiştirilebilir.'],
    [10, '', ''],
    [11, 'SAYFA YAPISI', 'header'],
    [12, '•', 'Malzeme Fiyat Listesi → Master DB (her şey buraya bağlı)'],
    [13, '•', 'Turkish Night, Mediterranean Night, BBQ Menu, Chicken BBQ, Fish BBQ, Breakfast → Her menü için ayrı maliyet sayfası'],
    [14, '•', 'Alışveriş Özeti → Menü × Kişi seç, alacağın tüm malzemeyi gör'],
    [15, '', ''],
    [16, 'ÖNEMLİ NOTLAR', 'header'],
    [17, '!', 'Fiyatlar 2026 Türkiye toptan tahminidir. Kendi tedarikçi fiyatlarınla güncelle.'],
    [18, '!', 'Birim çevirimleri otomatik (kg→g, lt→ml). Master birim sütunu ve mutfak gramajı uyumlu olmalı.'],
    [19, '!', 'Brüt kâr = Satış geliri − Malzeme maliyeti. Personel, kira, gaz, ulaşım dâhil DEĞİL.'],
    [20, '!', 'Yumurta/avokado/lavaş gibi adet bazlı malzemeler ondalık olabilir (örn. 0.3 adet).'],
    [21, '', ''],
    [22, 'MENÜ FİYATLARI VE MALZEME SAYISI', 'header'],
    [23, '•', `Turkish Night — 70£/kişi  ·  ${countItems('Turkish Night')} farklı malzeme`],
    [24, '•', `Mediterranean Night — 60£/kişi  ·  ${countItems('Mediterranean Night')} farklı malzeme`],
    [25, '•', `BBQ Menu — 50£/kişi  ·  ${countItems('BBQ Menu')} farklı malzeme`],
    [26, '•', `Chicken BBQ — 45£/kişi  ·  ${countItems('Chicken BBQ')} farklı malzeme`],
    [27, '•', `Fish BBQ — 65£/kişi  ·  ${countItems('Fish BBQ')} farklı malzeme`],
    [28, '•', `Breakfast — 20£/kişi  ·  ${countItems('Breakfast')} farklı malzeme`],
  ];

  function countItems(menuName) {
    const ids = new Set();
    MENUS[menuName].yemekler.forEach(y => y.items.forEach(it => ids.add(it.id)));
    return ids.size;
  }

  for (const [r, a, b] of lines) {
    if (b === 'header') {
      ws.mergeCells(r, 1, r, 2);
      const c = ws.getCell(r, 1);
      c.value = a;
      c.font = { bold: true, color: STYLE.cream, size: 12 };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: STYLE.brand };
      c.alignment = { horizontal: 'left', indent: 1, vertical: 'middle' };
      ws.getRow(r).height = 22;
    } else {
      ws.getCell(r, 1).value = a;
      ws.getCell(r, 1).font = { bold: true, color: STYLE.gold, size: 11 };
      ws.getCell(r, 1).alignment = { horizontal: 'center' };
      ws.getCell(r, 2).value = b;
      ws.getCell(r, 2).font = { color: STYLE.brand, size: 11 };
      ws.getCell(r, 2).alignment = { wrapText: true, vertical: 'middle', indent: 1 };
      ws.getRow(r).height = 20;
    }
  }
}

// ============================================================
// 8) ANA AKIŞ
// ============================================================
async function main() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Kalimera Kitchen';
  wb.created = new Date();
  wb.modified = new Date();
  wb.title = 'Maliyet Hesap Tablosu';
  wb.subject = 'Catering — Menü maliyet ve kâr analizi';

  // 1. Giriş sayfası
  const wsIntro = wb.addWorksheet('Başlangıç', { properties: { tabColor: STYLE.gold } });
  await buildIntroSheet(wsIntro);

  // 2. Malzeme fiyat listesi
  const wsPrice = wb.addWorksheet('Malzeme Fiyat Listesi', { properties: { tabColor: STYLE.brand2 } });
  await buildPriceSheet(wsPrice);

  // 3. Menü sayfaları
  for (const [name, data] of Object.entries(MENUS)) {
    const ws = wb.addWorksheet(name, { properties: { tabColor: STYLE.brand } });
    await buildMenuSheet(ws, name, data);
  }

  // 4. Alışveriş özeti
  const wsShop = wb.addWorksheet('Alışveriş Özeti', { properties: { tabColor: STYLE.gold } });
  await buildShoppingSheet(wsShop);

  const outPath = path.join(__dirname, 'Kalimera_Maliyet_Hesabi.xlsx');
  await wb.xlsx.writeFile(outPath);
  console.log(`✓ Excel dosyası oluşturuldu: ${outPath}`);
  console.log(`  Sayfa sayısı: ${wb.worksheets.length}`);
  console.log(`  Toplam malzeme: ${INGREDIENTS.length}`);
  console.log(`  Toplam menü: ${Object.keys(MENUS).length}`);
}

main().catch(err => {
  console.error('HATA:', err);
  process.exit(1);
});
