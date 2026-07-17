import { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import confetti from 'canvas-confetti';
import {
  ShoppingBag,
  Store,
  Sparkles,
  Clock,
  Database,
  CloudDownload,
  Upload,
  CheckCircle2,
  GitCompare,
  CheckSquare,
  Archive,
  Search,
  FileSpreadsheet,
  Copy,
  Ghost,
  Target,
  Star,
  X,
  ChevronLeft,
  ChevronRight,
  Info,
  AlertCircle,
  HelpCircle,
  TrendingDown
} from 'lucide-react';
import {
  Marketplace,
  TabType,
  Data1Source,
  ComparisonResults
} from './types';

export default function App() {
  // Primary States
  const [marketplace, setMarketplace] = useState<Marketplace>('shopee');
  const [data1Source, setData1Source] = useState<Data1Source>('auto_balist');
  const [file1, setFile1] = useState<File | null>(null);
  const [file3, setFile3] = useState<File | null>(null);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [hasProcessed, setHasProcessed] = useState<boolean>(false);
  const [activeTab, setActiveTab] = useState<TabType>('mismatch');

  // Search & Filters
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [currentPage, setCurrentPage] = useState<number>(1);
  const itemsPerPage = 100;

  // Filter Values
  const [min1, setMin1] = useState<string>('');
  const [max1, setMax1] = useState<string>('');
  const [minAio, setMinAio] = useState<string>('');
  const [maxAio, setMaxAio] = useState<string>('');

  const [categoryFilter, setCategoryFilter] = useState<string>('');
  const [brandFilter, setBrandFilter] = useState<string>('');
  const [sortOnly2, setSortOnly2] = useState<string>('');

  // Gamification & Mission States
  const [selectedForMission, setSelectedForMission] = useState<string[]>([]);
  const [selectedForMissionDisc, setSelectedForMissionDisc] = useState<string[]>([]);
  const [queueTimestamps, setQueueTimestamps] = useState<Record<string, number>>({});
  const [queueTimestampsDisc, setQueueTimestampsDisc] = useState<Record<string, number>>({});
  const [lastDurationRecord, setLastDurationRecord] = useState<string>("-");
  const [lastDurationRecordDisc, setLastDurationRecordDisc] = useState<string>("-");

  const [savedUploadData, setSavedUploadData] = useState<Record<string, boolean>>({});
  const [savedDiscData, setSavedDiscData] = useState<Record<string, boolean>>({});

  const dailyMissionTarget = 5;

  // Notification Toast State
  const [notification, setNotification] = useState<{ title: string; text: string; show: boolean; isWarning: boolean }>({
    title: "",
    text: "",
    show: false,
    isWarning: false
  });

  // Modal State for SKU Fixing Feature
  const [showFixSkuModal, setShowFixSkuModal] = useState<boolean>(false);
  const [modalSearchQuery, setModalSearchQuery] = useState<string>('');

  // Shopee Template State
  const [shopeeDetailsMap, setShopeeDetailsMap] = useState<Record<string, any>>({});
  const [shopeeHeaderRows, setShopeeHeaderRows] = useState<any[][]>([]);

  // Real-time Clock
  const [currentTime, setCurrentTime] = useState<string>("Memuat waktu...");
  useEffect(() => {
    const updateClock = () => {
      const now = new Date();
      setCurrentTime(
        now.toLocaleDateString('id-ID', {
          weekday: 'long',
          year: 'numeric',
          month: 'short',
          day: 'numeric',
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit',
        }).replace(/\./g, ':')
      );
    };
    updateClock();
    const interval = setInterval(updateClock, 1000);
    return () => clearInterval(interval);
  }, []);

  // Local Storage Load
  useEffect(() => {
    const today = new Date().toISOString().split('T')[0];
    const stored = localStorage.getItem('aioSyncData');
    if (stored) {
      try {
        const parsed = JSON.parse(stored);
        if (parsed.date === today) {
          setSavedUploadData(parsed.uploaded || {});
          setSavedDiscData(parsed.disc || {});
        } else {
          localStorage.setItem('aioSyncData', JSON.stringify({ date: today, uploaded: {}, disc: {} }));
        }
      } catch (e) {
        localStorage.setItem('aioSyncData', JSON.stringify({ date: today, uploaded: {}, disc: {} }));
      }
    } else {
      localStorage.setItem('aioSyncData', JSON.stringify({ date: today, uploaded: {}, disc: {} }));
    }
  }, []);

  // Save to Local Storage Helper
  const saveToLocal = (uploaded: Record<string, boolean>, disc: Record<string, boolean>) => {
    const today = new Date().toISOString().split('T')[0];
    localStorage.setItem('aioSyncData', JSON.stringify({
      date: today,
      uploaded,
      disc
    }));
  };

  // Results State
  const [comparisonResults, setComparisonResults] = useState<ComparisonResults>({
    mismatch: [],
    match: [],
    only1: [],
    only2: [],
    discontinued: [],
    shortSku: []
  });

  // Toast Trigger Helper
  const triggerNotification = (title: string, text: string, isWarning: boolean) => {
    setNotification({ title, text, show: true, isWarning });
  };

  useEffect(() => {
    if (notification.show) {
      const timer = setTimeout(() => {
        setNotification(prev => ({ ...prev, show: false }));
      }, 3000);
      return () => clearTimeout(timer);
    }
  }, [notification.show]);

  // Read Excel / CSV Helper
  const readExcelFile = (file: File, startRow: number): Promise<any[][]> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const wb = XLSX.read(data, { type: 'array' });
          const ws = wb.Sheets[wb.SheetNames[0]];
          const json = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1, raw: false, defval: "" });
          resolve(json.slice(startRow - 1));
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  };

  // Run Comparison Engine
  const startComparison = async () => {
    setIsProcessing(true);
    try {
      let data1: any[][] = [];
      let headerRows: any[][] = [];
      if (marketplace === 'shopee') {
        if (data1Source === 'manual') {
          if (!file1) {
            throw new Error("Harap unggah file Data 1 Shopee manual.");
          }
          const fullData1 = await readExcelFile(file1, 1);
          headerRows = fullData1.slice(0, 6);
          data1 = fullData1.slice(6);
        } else {
          // Fetch from Google Sheets published XLSX
          const url1 = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQUJWBw2EXirlxov14JNpI1h3ulExBcMQxQ5orpGZmpW7cMqUqMkU9E6OxJ4CBLd4ZvAW8tBmhmEEF6/pub?output=xlsx";
          const res1 = await fetch(url1);
          if (!res1.ok) throw new Error("Gagal mengunduh Data 1 dari Google Sheets");
          const buf1 = await res1.arrayBuffer();
          const wb1 = XLSX.read(buf1, { type: 'array' });
          const targetSheetName = data1Source === 'auto_balist' ? 'ShopeeBalist' : 'GomallShopee';
          if (!wb1.Sheets[targetSheetName]) {
            throw new Error(`Sheet '${targetSheetName}' tidak ditemukan di Google Sheets!`);
          }
          const ws1 = wb1.Sheets[targetSheetName];
          const rawData1 = XLSX.utils.sheet_to_json<any[]>(ws1, { header: 1, raw: false, defval: "" });
          headerRows = rawData1.slice(0, 6);
          data1 = rawData1.slice(6); // Mulai baris ke-7
        }
        setShopeeHeaderRows(headerRows);
      } else {
        if (!file1) {
          throw new Error("Harap unggah file Data 1 Tokopedia.");
        }
        data1 = await readExcelFile(file1, 4);
      }

      // Fetch DATA AIO from published CSV
      const aioUrl = "https://docs.google.com/spreadsheets/d/e/2PACX-1vTCxz1GPm7QU9IS1yBiSjvIdNTLUsvvplOCyT_R3XH4O-LuVbHoY_bXn1LTH5lpnlolJ29BhUgEdnFm/pub?gid=1564332470&single=true&output=csv";
      const aioRes = await fetch(aioUrl);
      if (!aioRes.ok) throw new Error("Gagal mengunduh DATA AIO dari Google Sheets");
      const aioBuffer = await aioRes.arrayBuffer();
      const aioWb = XLSX.read(aioBuffer, { type: 'array' });
      const aioWs = aioWb.Sheets[aioWb.SheetNames[0]];
      const rawDataAio = XLSX.utils.sheet_to_json<any[]>(aioWs, { header: 1, raw: false, defval: "" });
      const dataAio = rawDataAio.slice(1); // Mulai baris ke-2

      // Fetch optional Data 3
      const data3 = file3 ? await readExcelFile(file3, 2) : [];

      const map1 = new Map<string, { stock: number }>();
      const detailsMap: Record<string, any> = {};
      data1.forEach(row => {
        let sku = "";
        if (marketplace === 'shopee') {
          sku = String(row[5] || "").trim() || String(row[4] || "").trim();
          if (sku) {
            detailsMap[sku] = {
              productId: String(row[0] || ""),
              variationId: String(row[1] || ""),
              productName: String(row[2] || ""),
              variationName: String(row[3] || ""),
              parentSku: String(row[4] || ""),
              variationSku: String(row[5] || ""),
              price: row[6],
              currentStock: row[7],
              originalRow: row
            };
          }
        } else {
          sku = String(row[3] || "").trim();
        }
        let stock = parseFloat(row[8]);
        stock = isNaN(stock) ? 0 : stock;
        if (sku) map1.set(sku, { stock });
      });
      if (marketplace === 'shopee') {
        setShopeeDetailsMap(detailsMap);
      }

      const mapAio = new Map<string, { stock: number; name: string; category: string; brand: string }>();
      dataAio.forEach(row => {
        let sku = String(row[0] || "").trim();
        let name = String(row[2] || "").trim() || "-";
        let category = String(row[4] || "").trim() || "-";
        let brand = String(row[6] || "").trim() || "-";
        let stock = parseFloat(row[12]);
        stock = isNaN(stock) ? 0 : stock;
        if (sku) mapAio.set(sku, { stock, name, category, brand });
      });

      const map3 = new Map<string, { stock: number; name: string }>();
      data3.forEach(row => {
        let sku = String(row[0] || "").trim();
        let name = String(row[2] || "").trim() || "-";
        let stock = parseFloat(row[12]);
        stock = isNaN(stock) ? 0 : stock;
        if (sku) map3.set(sku, { stock, name });
      });

      const results: ComparisonResults = {
        mismatch: [],
        match: [],
        only1: [],
        only2: [],
        discontinued: [],
        shortSku: []
      };

      map1.forEach((val1, sku) => {
        // Cek Discontinued
        if (map3.has(sku)) {
          results.discontinued.push({
            sku,
            name: map3.get(sku)!.name,
            stock1: val1.stock,
            stock3: map3.get(sku)!.stock
          });
        }
        // Cek Normal
        if (mapAio.has(sku)) {
          const valAio = mapAio.get(sku)!;
          if (val1.stock === valAio.stock) {
            results.match.push({
              sku,
              name: valAio.name,
              stock1: val1.stock,
              stockAio: valAio.stock
            });
          } else {
            results.mismatch.push({
              sku,
              name: valAio.name,
              stock1: val1.stock,
              stockAio: valAio.stock
            });
          }
        } else {
          results.only1.push({
            sku,
            stock1: val1.stock
          });
        }
      });

      mapAio.forEach((valAio, sku) => {
        if (!map1.has(sku)) {
          results.only2.push({
            sku,
            name: valAio.name,
            stockAio: valAio.stock,
            category: valAio.category,
            brand: valAio.brand
          });
        }
      });

      // Analyze Short SKUs (< 5 digits)
      const shortSkusMap = new Map<string, { sku: string; suggestedSku: string; name: string; source: string; stock: number; existsInAio: boolean }>();

      // 1. Check Data 1
      map1.forEach((val, sku) => {
        if (sku.length > 0 && sku.length < 5 && /^\d+$/.test(sku)) {
          const suggested = sku.padStart(5, '0');
          const name = mapAio.get(sku)?.name || map3.get(sku)?.name || "-";
          const existsInAio = mapAio.has(suggested);
          const realName = existsInAio ? mapAio.get(suggested)!.name : name;

          shortSkusMap.set(sku, {
            sku,
            suggestedSku: suggested,
            name: realName,
            source: "Data 1",
            stock: val.stock,
            existsInAio
          });
        }
      });

      // 2. Check DATA AIO
      mapAio.forEach((val, sku) => {
        if (sku.length > 0 && sku.length < 5 && /^\d+$/.test(sku)) {
          const suggested = sku.padStart(5, '0');
          const existsInAio = mapAio.has(suggested);
          const realName = existsInAio ? mapAio.get(suggested)!.name : val.name;

          if (!shortSkusMap.has(sku)) {
            shortSkusMap.set(sku, {
              sku,
              suggestedSku: suggested,
              name: realName,
              source: "DATA AIO",
              stock: val.stock,
              existsInAio
            });
          } else {
            const item = shortSkusMap.get(sku)!;
            if (item.name === "-" || !item.name) {
              item.name = realName;
            }
          }
        }
      });

      // 3. Check Data 3
      map3.forEach((val, sku) => {
        if (sku.length > 0 && sku.length < 5 && /^\d+$/.test(sku)) {
          const suggested = sku.padStart(5, '0');
          const existsInAio = mapAio.has(suggested);
          const realName = existsInAio ? mapAio.get(suggested)!.name : val.name;

          if (!shortSkusMap.has(sku)) {
            shortSkusMap.set(sku, {
              sku,
              suggestedSku: suggested,
              name: realName,
              source: "Data 3",
              stock: val.stock,
              existsInAio
            });
          } else {
            const item = shortSkusMap.get(sku)!;
            if (item.name === "-" || !item.name) {
              item.name = realName;
            }
          }
        }
      });

      results.shortSku = Array.from(shortSkusMap.values());

      setComparisonResults(results);
      setHasProcessed(true);
      setActiveTab('mismatch');
      triggerNotification("Sukses", "Data berhasil dianalisis!", false);
    } catch (err: any) {
      console.error(err);
      triggerNotification("Peringatan", err?.message || "Gagal memproses data.", true);
    } finally {
      setIsProcessing(false);
    }
  };

  // Clipboard Copier
  const copyToClipboard = (text: string, typeName: string) => {
    navigator.clipboard.writeText(text).then(() => {
      triggerNotification("Tersalin", `${typeName} berhasil disalin ke papan klip.`, false);
    }).catch(() => {
      // Fallback
      const el = document.createElement('textarea');
      el.value = text;
      document.body.appendChild(el);
      el.select();
      document.execCommand('copy');
      document.body.removeChild(el);
      triggerNotification("Tersalin", `${typeName} berhasil disalin ke papan klip.`, false);
    });
  };

  // Filter & Search computation
  const getFilteredData = () => {
    let data: any[] = [];
    if (activeTab === 'mismatch') data = comparisonResults.mismatch;
    else if (activeTab === 'match') data = comparisonResults.match;
    else if (activeTab === 'only1') data = comparisonResults.only1;
    else if (activeTab === 'only2') data = comparisonResults.only2;
    else if (activeTab === 'discontinued') data = comparisonResults.discontinued;
    else if (activeTab === 'shortSku') data = comparisonResults.shortSku || [];

    if (searchQuery.trim()) {
      const q = searchQuery.toLowerCase();
      data = data.filter(item =>
        String(item.sku || "").toLowerCase().includes(q) ||
        String(item.name || "").toLowerCase().includes(q)
      );
    }

    if (activeTab === 'mismatch') {
      data = data.filter(item => {
        let pass = true;
        if (min1 !== "") pass = pass && item.stock1 >= parseFloat(min1);
        if (max1 !== "") pass = pass && item.stock1 <= parseFloat(max1);
        if (minAio !== "") pass = pass && item.stockAio >= parseFloat(minAio);
        if (maxAio !== "") pass = pass && item.stockAio <= parseFloat(maxAio);
        return pass;
      });
    }

    if (activeTab === 'only2') {
      if (categoryFilter) {
        data = data.filter(item => (String(item.category || "").trim() || "-") === categoryFilter);
      }
      if (brandFilter) {
        data = data.filter(item => (String(item.brand || "").trim() || "-") === brandFilter);
      }
      if (sortOnly2 === 'desc') {
        data = [...data].sort((a, b) => b.stockAio - a.stockAio);
      } else if (sortOnly2 === 'asc') {
        data = [...data].sort((a, b) => a.stockAio - b.stockAio);
      }
    }

    return data;
  };

  // Unique Categories/Brands computed options
  const getCategoryOptions = () => {
    const counts: Record<string, number> = {};
    comparisonResults.only2.forEach(item => {
      const cat = String(item.category || "").trim() || "-";
      counts[cat] = (counts[cat] || 0) + 1;
    });
    return Object.keys(counts).sort().map(cat => ({ value: cat, label: `${cat} (${counts[cat]})` }));
  };

  const getBrandOptions = () => {
    const counts: Record<string, number> = {};
    comparisonResults.only2.forEach(item => {
      const brand = String(item.brand || "").trim() || "-";
      counts[brand] = (counts[brand] || 0) + 1;
    });
    return Object.keys(counts).sort().map(brand => ({ value: brand, label: `${brand} (${counts[brand]})` }));
  };

  const filteredData = getFilteredData();
  const totalItems = filteredData.length;
  const totalPages = Math.ceil(totalItems / itemsPerPage);
  const paginatedData = filteredData.slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage);

  // Pagination Handlers
  const handlePrevPage = () => {
    if (currentPage > 1) {
      setCurrentPage(prev => prev - 1);
    }
  };

  const handleNextPage = () => {
    if (currentPage < totalPages) {
      setCurrentPage(prev => prev + 1);
    }
  };

  // Helper to generate Shopee-compatible 2D Array data (matching original/fallback Shopee Mass Update template)
  const generateShopeeExcelData = (items: any[], isShortSkuTab: boolean = false) => {
    const wsData: any[][] = [];
    
    if (shopeeHeaderRows && shopeeHeaderRows.length >= 6) {
      // Use original Shopee header rows to preserve formatting and exact language (ID/EN)
      shopeeHeaderRows.forEach(row => wsData.push([...row]));
    } else {
      // Fallback Shopee Mass Update Stock template structure
      wsData.push(["Ubah Stok Massal"]);
      wsData.push(["Instruksi: 1. Silakan masukkan stok terbaru Anda di kolom 'Stok Baru'."]);
      wsData.push(["2. Harap jangan mengubah isi kolom lainnya."]);
      wsData.push(["3. Setelah selesai, simpan file ini dan upload kembali ke Seller Centre."]);
      wsData.push([]);
      wsData.push([
        "ID Produk",
        "ID Variasi",
        "Nama Produk",
        "Nama Variasi",
        "SKU Induk",
        "Nomor Referensi SKU",
        "Harga",
        "Stok Saat Ini",
        "Stok Baru"
      ]);
    }

    items.forEach(item => {
      const details = shopeeDetailsMap[item.sku] || {};
      const row: any[] = [];
      row[0] = details.productId || "";
      row[1] = details.variationId || "";
      row[2] = details.productName || item.name || "";
      row[3] = details.variationName || "";

      let parentSku = details.parentSku || "";
      let variationSku = details.variationSku || "";

      if (isShortSkuTab) {
        // Replace with the 5-digit suggested SKU
        if (details.variationSku) {
          variationSku = item.suggestedSku;
        } else if (details.parentSku) {
          parentSku = item.suggestedSku;
        } else {
          variationSku = item.suggestedSku;
        }
      }

      row[4] = parentSku;
      row[5] = variationSku || item.sku;
      row[6] = details.price !== undefined ? details.price : "";
      row[7] = details.currentStock !== undefined ? details.currentStock : (item.stock1 !== undefined ? item.stock1 : "");

      // Determine stock value for last column
      let stockVal: any = "";
      if (isShortSkuTab) {
        stockVal = item.stock !== undefined ? item.stock : "";
      } else {
        if (activeTab === 'mismatch' || activeTab === 'match') {
          stockVal = item.stockAio !== undefined ? item.stockAio : "";
        } else if (activeTab === 'only1') {
          stockVal = item.stock1 !== undefined ? item.stock1 : "";
        } else if (activeTab === 'only2') {
          stockVal = item.stockAio !== undefined ? item.stockAio : "";
        } else if (activeTab === 'discontinued') {
          stockVal = item.stock3 !== undefined ? item.stock3 : "0";
        }
      }
      row[8] = stockVal;

      wsData.push(row);
    });

    return wsData;
  };

  // Helper to force specific Excel columns to be treated as pure text/string (prevents Excel from stripping leading zeros)
  const forceTextColumns = (ws: any, colIndices: number[], startRow: number = 7) => {
    if (!ws || !ws['!ref']) return;
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let r = startRow - 1; r <= range.e.r; r++) {
      colIndices.forEach(c => {
        const cellRef = XLSX.utils.encode_cell({ r, c });
        const cell = ws[cellRef];
        if (cell && cell.v !== undefined) {
          const valStr = String(cell.v).trim();
          ws[cellRef] = { t: 's', v: valStr };
        }
      });
    }
  };

  // Excel Downloader
  const handleExportToExcel = () => {
    if (filteredData.length === 0) {
      triggerNotification("Kosong", "Tidak ada data untuk diexport", true);
      return;
    }

    const wb = XLSX.utils.book_new();

    if (marketplace === 'shopee') {
      const wsData = generateShopeeExcelData(filteredData, activeTab === 'shortSku');
      const ws = XLSX.utils.aoa_to_sheet(wsData);
      
      // Force ID Produk (col 0), ID Variasi (col 1), SKU Induk (col 4), Nomor Referensi SKU (col 5) as pure text
      forceTextColumns(ws, [0, 1, 4, 5], 7);

      // Auto-adjust column widths for better readability
      const maxCols = wsData.reduce((max, row) => Math.max(max, row.length), 0);
      ws['!cols'] = Array.from({ length: maxCols }, () => ({ wch: 18 }));
      
      XLSX.utils.book_append_sheet(wb, ws, "Laporan Shopee");
      XLSX.writeFile(wb, `Ubah_Stok_Shopee_${activeTab}.xlsx`);
    } else {
      const exp = filteredData.map(item => {
        const r: Record<string, any> = { "Kode SKU": item.sku };
        if (activeTab !== 'only1') {
          r["Nama Produk"] = item.name;
        }

        if (activeTab === 'mismatch' || activeTab === 'match') {
          r["Stok Data 1"] = item.stock1;
          r["Stok AIO"] = item.stockAio;
        } else if (activeTab === 'only1') {
          r["Stok Data 1"] = item.stock1;
        } else if (activeTab === 'only2') {
          r["Stok AIO"] = item.stockAio;
          r["Kategori"] = item.category;
          r["Merk"] = item.brand;
          r["Status"] = savedUploadData[item.sku] ? "Sudah Terupload" : "Belum Terupload";
        } else if (activeTab === 'discontinued') {
          r["Stok Data 1"] = item.stock1;
          r["Stok Data 3"] = item.stock3;
          r["Status"] = savedDiscData[item.sku] ? "Sudah Diproses" : "Belum Diproses";
        } else if (activeTab === 'shortSku') {
          r["SKU Baru (5 Digit)"] = item.suggestedSku;
          r["Sumber Data"] = item.source;
          r["Stok"] = item.stock;
          r["Terdaftar di AIO"] = item.existsInAio ? "Ya" : "Tidak";
        }
        return r;
      });

      const ws = XLSX.utils.json_to_sheet(exp);
      
      // Force "Kode SKU" (col 0) as pure text, and also "SKU Baru (5 Digit)" (col 1) if on shortSku tab
      if (activeTab === 'shortSku') {
        forceTextColumns(ws, [0, 1], 2);
      } else {
        forceTextColumns(ws, [0], 2);
      }

      XLSX.utils.book_append_sheet(wb, ws, "Laporan");
      XLSX.writeFile(wb, `Laporan_${activeTab}_tokopedia.xlsx`);
    }
    triggerNotification("Berhasil", "File Excel telah didownload.", false);
  };

  // SKU Fix Downloader & Copier Helper Functions
  const handleDownloadSkuFixes = () => {
    const dataToExport = comparisonResults.shortSku || [];
    if (dataToExport.length === 0) {
      triggerNotification("Kosong", "Tidak ada data SKU < 5 digit untuk diperbaiki.", true);
      return;
    }

    const wb = XLSX.utils.book_new();

    if (marketplace === 'shopee') {
      const wsData = generateShopeeExcelData(dataToExport, true);
      const ws = XLSX.utils.aoa_to_sheet(wsData);
      
      // Force ID Produk (col 0), ID Variasi (col 1), SKU Induk (col 4), Nomor Referensi SKU (col 5) as pure text
      forceTextColumns(ws, [0, 1, 4, 5], 7);

      // Auto-adjust column widths for better readability
      const maxCols = wsData.reduce((max, row) => Math.max(max, row.length), 0);
      ws['!cols'] = Array.from({ length: maxCols }, () => ({ wch: 18 }));
      
      XLSX.utils.book_append_sheet(wb, ws, "Perbaikan SKU Shopee");
      XLSX.writeFile(wb, `Ubah_Stok_Shopee_Perbaikan_SKU.xlsx`);
    } else {
      const rows = dataToExport.map(item => ({
        "SKU Asli": item.sku,
        "SKU Baru (5 Digit)": item.suggestedSku,
        "Nama Produk": item.name,
        "Sumber Data": item.source,
        "Stok": item.stock,
        "Terdaftar di AIO": item.existsInAio ? "Ya" : "Tidak"
      }));

      const ws = XLSX.utils.json_to_sheet(rows);
      
      // Force "SKU Asli" (col 0) and "SKU Baru (5 Digit)" (col 1) as pure text
      forceTextColumns(ws, [0, 1], 2);

      XLSX.utils.book_append_sheet(wb, ws, "Perbaikan SKU");
      XLSX.writeFile(wb, `Perbaikan_SKU_Format_tokopedia.xlsx`);
    }
    triggerNotification("Sukses", "File rekomendasi perbaikan format SKU telah diunduh.", false);
  };

  const copyAllFixedSkus = () => {
    const dataToExport = comparisonResults.shortSku || [];
    if (dataToExport.length === 0) {
      triggerNotification("Kosong", "Tidak ada data SKU untuk disalin.", true);
      return;
    }
    const skusString = dataToExport.map(item => item.suggestedSku).join('\n');
    navigator.clipboard.writeText(skusString)
      .then(() => {
        triggerNotification("Berhasil", `${dataToExport.length} SKU Baru berhasil disalin ke clipboard.`, false);
      })
      .catch(() => {
        triggerNotification("Gagal", "Tidak dapat menyalin data ke clipboard.", true);
      });
  };

  // Gamification Actions
  const handleQueueAction = (sku: string, type: 'up' | 'disc') => {
    if (type === 'up') {
      if (selectedForMission.includes(sku)) {
        setSelectedForMission(prev => prev.filter(s => s !== sku));
        setQueueTimestamps(prev => {
          const next = { ...prev };
          delete next[sku];
          return next;
        });
      } else {
        setSelectedForMission(prev => [...prev, sku]);
        setQueueTimestamps(prev => ({ ...prev, [sku]: Date.now() }));
      }
    } else {
      if (selectedForMissionDisc.includes(sku)) {
        setSelectedForMissionDisc(prev => prev.filter(s => s !== sku));
        setQueueTimestampsDisc(prev => {
          const next = { ...prev };
          delete next[sku];
          return next;
        });
      } else {
        setSelectedForMissionDisc(prev => [...prev, sku]);
        setQueueTimestampsDisc(prev => ({ ...prev, [sku]: Date.now() }));
      }
    }
  };

  const formatDuration = (ms: number) => {
    if (ms < 0 || isNaN(ms)) return "-";
    const s = Math.floor(ms / 1000);
    const m = Math.floor(s / 60);
    const sc = s % 60;
    return m > 0 ? `${m}m ${sc}s` : `${sc} detik`;
  };

  const handleExecuteAction = (sku: string, type: 'up' | 'disc') => {
    const now = Date.now();
    if (type === 'up') {
      if (queueTimestamps[sku]) {
        setLastDurationRecord(`Durasi Upload: ${formatDuration(now - queueTimestamps[sku])}`);
      }
      setSelectedForMission(prev => prev.filter(s => s !== sku));
      setQueueTimestamps(prev => {
        const next = { ...prev };
        delete next[sku];
        return next;
      });

      const updatedUploads = { ...savedUploadData, [sku]: true };
      setSavedUploadData(updatedUploads);
      saveToLocal(updatedUploads, savedDiscData);

      const count = Object.keys(updatedUploads).length;
      if (count > 0 && count % dailyMissionTarget === 0) {
        fireBigConfetti();
        triggerNotification("Misi Selesai!", "Kerja bagus, target tercapai!", false);
      } else {
        fireSmallConfetti('up');
      }
    } else {
      if (queueTimestampsDisc[sku]) {
        setLastDurationRecordDisc(`Durasi Arsip: ${formatDuration(now - queueTimestampsDisc[sku])}`);
      }
      setSelectedForMissionDisc(prev => prev.filter(s => s !== sku));
      setQueueTimestampsDisc(prev => {
        const next = { ...prev };
        delete next[sku];
        return next;
      });

      const updatedDisc = { ...savedDiscData, [sku]: true };
      setSavedDiscData(updatedDisc);
      saveToLocal(savedUploadData, updatedDisc);

      const count = Object.keys(updatedDisc).length;
      if (count > 0 && count % dailyMissionTarget === 0) {
        fireBigConfetti();
        triggerNotification("Misi Selesai!", "Kerja bagus, target tercapai!", false);
      } else {
        fireSmallConfetti('disc');
      }
    }
  };

  const handleToggleAction = (sku: string, type: 'up' | 'disc') => {
    if (type === 'up') {
      const updatedUploads = { ...savedUploadData };
      delete updatedUploads[sku];
      setSavedUploadData(updatedUploads);
      saveToLocal(updatedUploads, savedDiscData);
    } else {
      const updatedDisc = { ...savedDiscData };
      delete updatedDisc[sku];
      setSavedDiscData(updatedDisc);
      saveToLocal(savedUploadData, updatedDisc);
    }
  };

  // Confetti Animations
  const fireBigConfetti = () => {
    const duration = 2.5 * 1000;
    const end = Date.now() + duration;

    const interval = setInterval(() => {
      const timeLeft = end - Date.now();
      if (timeLeft <= 0) return clearInterval(interval);

      const particleCount = 40 * (timeLeft / duration);
      confetti({
        startVelocity: 25,
        spread: 360,
        ticks: 60,
        zIndex: 100,
        particleCount,
        origin: { x: Math.random() * (0.3 - 0.1) + 0.1, y: Math.random() - 0.2 }
      });
      confetti({
        startVelocity: 25,
        spread: 360,
        ticks: 60,
        zIndex: 100,
        particleCount,
        origin: { x: Math.random() * (0.9 - 0.7) + 0.7, y: Math.random() - 0.2 }
      });
    }, 250);
  };

  const fireSmallConfetti = (type: 'up' | 'disc') => {
    confetti({
      particleCount: 40,
      spread: 50,
      origin: { y: 0.8 },
      zIndex: 100,
      colors: type === 'up'
        ? ['#6366f1', '#fbbf24', '#10b981']
        : ['#f59e0b', '#fbbf24', '#ef4444']
    });
  };

  // Gamification widget computation
  const activeMissionCount = Object.keys(activeTab === 'only2' ? savedUploadData : savedDiscData).length;
  let activeMissionProgress = activeMissionCount % dailyMissionTarget;
  if (activeMissionCount > 0 && activeMissionProgress === 0) {
    activeMissionProgress = dailyMissionTarget;
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-700 font-sans pb-32">
      <div className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-8 py-8 relative">

        {/* Marketplace Switcher */}
        <div className="flex justify-center mb-8">
          <div className="bg-slate-200/70 p-1.5 rounded-full inline-flex items-center gap-1 shadow-inner border border-slate-300">
            <button
              onClick={() => {
                setMarketplace('shopee');
                setFile1(null);
                setCurrentPage(1);
              }}
              className={`px-6 py-2 rounded-full text-sm font-bold transition-all duration-300 w-36 flex items-center justify-center gap-1.5 ${
                marketplace === 'shopee'
                  ? 'bg-orange-500 text-white shadow-md'
                  : 'text-slate-500 hover:text-slate-700 hover:bg-slate-300/30'
              }`}
            >
              <ShoppingBag className="w-4 h-4" />
              Shopee
            </button>
            <button
              onClick={() => {
                setMarketplace('tokopedia');
                setFile1(null);
                setCurrentPage(1);
              }}
              className={`px-6 py-2 rounded-full text-sm font-bold transition-all duration-300 w-36 flex items-center justify-center gap-1.5 ${
                marketplace === 'tokopedia'
                  ? 'bg-emerald-500 text-white shadow-md'
                  : 'text-slate-500 hover:text-slate-700 hover:bg-slate-300/30'
              }`}
            >
              <Store className="w-4 h-4" />
              Tokopedia
            </button>
          </div>
        </div>

        {/* Header Block */}
        <header className="mb-10 text-center md:text-left flex flex-col md:flex-row justify-between items-center gap-6">
          <div>
            <div className={`inline-flex items-center gap-2 px-3 py-1 rounded-full text-xs font-semibold mb-3 tracking-wide shadow-sm border transition-colors ${
              marketplace === 'shopee'
                ? 'bg-orange-50 border-orange-200 text-orange-600'
                : 'bg-emerald-50 border-emerald-200 text-emerald-600'
            }`}>
              <Sparkles className="w-3.5 h-3.5 animate-pulse" />
              Mode {marketplace === 'shopee' ? 'Shopee' : 'Tokopedia'} Aktif
            </div>
            <h1 className="text-3xl md:text-4xl font-bold text-slate-800 tracking-tight">
              Komparasi SKU & Stok
            </h1>

            <div className="flex items-center justify-center md:justify-start gap-2 text-xs font-semibold text-slate-500 bg-white border border-slate-200 px-3 py-2 rounded-lg shadow-sm w-fit mt-3 mx-auto md:mx-0">
              <Clock className={`w-4 h-4 ${marketplace === 'shopee' ? 'text-orange-500' : 'text-emerald-500'}`} />
              <span>{currentTime}</span>
            </div>
          </div>

          <button
            onClick={startComparison}
            disabled={isProcessing}
            className={`text-white px-7 py-3.5 rounded-xl font-bold transition-all active:scale-95 flex items-center gap-3 shadow-md disabled:opacity-50 disabled:cursor-not-allowed ${
              marketplace === 'shopee'
                ? 'bg-orange-500 hover:bg-orange-600 shadow-orange-200'
                : 'bg-emerald-500 hover:bg-emerald-600 shadow-emerald-200'
            }`}
          >
            <Sparkles className="w-5 h-5" />
            <span>{isProcessing ? 'Memproses...' : 'Mulai Analisis'}</span>
            {isProcessing && (
              <div className="w-4 h-4 border-2 border-white border-b-transparent rounded-full animate-spin"></div>
            )}
          </button>
        </header>

        {/* Upload Panels */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
          
          {/* File Data 1 Card */}
          <div className="bg-white border border-slate-200 p-6 rounded-2xl shadow-sm hover:border-slate-300 transition-colors relative flex flex-col justify-between">
            <div>
              <h3 className="text-sm font-bold uppercase tracking-wider text-slate-500 mb-1 flex items-center gap-2">
                <Database className="w-4 h-4 text-slate-400" />
                File Data 1
                <span className={`px-2 py-0.5 rounded text-[10px] uppercase border font-extrabold ${
                  marketplace === 'shopee'
                    ? 'bg-orange-100 text-orange-700 border-orange-200'
                    : 'bg-emerald-100 text-emerald-700 border-emerald-200'
                }`}>
                  {marketplace}
                </span>
              </h3>
              <p className="text-xs text-slate-500 mb-4 leading-relaxed">
                {marketplace === 'shopee'
                  ? 'Mulai baris ke-7 | SKU: Kolom 6 (jika kosong ambil Kol 5) | Stok: Kol 9'
                  : 'Mulai baris ke-4 | SKU: Kolom 4 | Stok: Kol 9'}
              </p>

              {/* Source selection specifically for Shopee */}
              {marketplace === 'shopee' && (
                <div className="relative mb-4">
                  <select
                    value={data1Source}
                    onChange={(e) => {
                      setData1Source(e.target.value as Data1Source);
                      setFile1(null);
                    }}
                    className="w-full px-3 py-2 bg-slate-50 hover:bg-slate-100 border border-slate-300 rounded-lg text-xs font-semibold cursor-pointer outline-none transition-colors text-slate-700 appearance-none pr-8"
                  >
                    <option value="auto_balist">Auto-Fetch: Sheet ShopeeBalist</option>
                    <option value="auto_gomall">Auto-Fetch: Sheet GomallShopee</option>
                    <option value="manual">Upload File Manual</option>
                  </select>
                  <div className="absolute right-2.5 top-1/2 -translate-y-1/2 pointer-events-none text-slate-500">
                    ▼
                  </div>
                </div>
              )}
            </div>

            {/* Display correct input depending on selected source */}
            {marketplace === 'shopee' && data1Source !== 'manual' ? (
              <div className="flex flex-col items-center justify-center w-full h-32 border-2 border-orange-200 border-dashed rounded-xl bg-orange-50/40 shadow-inner transition-colors">
                <CloudDownload className="w-8 h-8 text-orange-400 mb-1.5 animate-bounce" />
                <p className="text-xs font-bold text-orange-600">Auto-Fetch Data 1 Aktif</p>
                <p className="text-[10px] text-slate-500 mt-1 font-semibold bg-white px-2 py-0.5 rounded border border-orange-100 shadow-sm">
                  Sheet: {data1Source === 'auto_balist' ? 'ShopeeBalist' : 'GomallShopee'}
                </p>
              </div>
            ) : (
              <label className="flex flex-col items-center justify-center w-full h-32 border-2 border-slate-200 border-dashed rounded-xl cursor-pointer bg-slate-50 hover:bg-slate-100/70 transition-colors group">
                {!file1 ? (
                  <div className="flex flex-col items-center justify-center pt-4 pb-4">
                    <Upload className="w-8 h-8 text-slate-400 group-hover:text-slate-600 mb-1.5 transition-colors" />
                    <p className="text-xs font-bold text-slate-600">Pilih Data 1 {marketplace === 'shopee' ? 'Shopee' : 'Tokopedia'}</p>
                    <p className="text-[10px] text-slate-400 mt-0.5">XLSX, XLS, CSV</p>
                  </div>
                ) : (
                  <div className="flex flex-col items-center justify-center w-full h-full bg-emerald-50 rounded-xl border border-emerald-100 px-4 text-center">
                    <CheckCircle2 className="w-8 h-8 text-emerald-500 mb-1.5" />
                    <p className="text-xs font-bold text-emerald-700 truncate w-full">{file1.name}</p>
                    <button
                      type="button"
                      onClick={(e) => {
                        e.preventDefault();
                        setFile1(null);
                      }}
                      className="text-[10px] text-rose-500 hover:underline mt-1 font-bold"
                    >
                      Hapus
                    </button>
                  </div>
                )}
                <input
                  type="file"
                  className="hidden"
                  accept=".xlsx, .xls, .csv"
                  onChange={(e) => {
                    const f = e.target.files?.[0];
                    if (f) setFile1(f);
                  }}
                />
              </label>
            )}
          </div>

          {/* DATA AIO Panel */}
          <div className="bg-indigo-50/40 border border-indigo-100 p-6 rounded-2xl shadow-sm hover:border-indigo-200 transition-colors flex flex-col justify-between">
            <div>
              <h3 className="text-sm font-bold uppercase tracking-wider text-indigo-800 mb-1 flex items-center gap-2">
                <Database className="w-4 h-4 text-indigo-500" />
                DATA AIO (Auto-Fetch)
              </h3>
              <p className="text-xs text-slate-500 mb-4 leading-relaxed">
                Mulai baris ke-2 | SKU: Kol 1 | Stok: Kol 13 | Nama: Kol 3 | Kategori: Kol 5 | Merk: Kol 7
              </p>
            </div>

            <div className="flex flex-col items-center justify-center w-full h-32 border-2 border-indigo-200 border-dashed rounded-xl bg-indigo-50/60 shadow-inner transition-colors">
              <CloudDownload className="w-8 h-8 text-indigo-400 mb-1.5 animate-bounce" />
              <p className="text-xs font-bold text-indigo-600">Auto-Fetch AIO Aktif</p>
              <p className="text-[10px] text-slate-500 mt-1 font-semibold bg-white px-2 py-0.5 rounded border border-indigo-100 shadow-sm">
                Sheet: STOCK LIST
              </p>
            </div>
          </div>

          {/* Data 3 Panel */}
          <div className="bg-white border border-slate-200 p-6 rounded-2xl shadow-sm hover:border-slate-300 transition-colors flex flex-col justify-between">
            <div>
              <h3 className="text-sm font-bold uppercase tracking-wider text-slate-500 mb-1 flex items-center gap-2">
                <Archive className="w-4 h-4 text-slate-400" />
                File Data 3
                <span className="bg-slate-100 text-slate-500 px-2 py-0.5 rounded text-[10px] lowercase border border-slate-200">
                  opsional
                </span>
              </h3>
              <p className="text-xs text-slate-500 mb-4 leading-relaxed">
                Untuk cek Discontinued | Mulai baris ke-2 | SKU: Kolom 1
              </p>
            </div>

            <label className="flex flex-col items-center justify-center w-full h-32 border-2 border-slate-200 border-dashed rounded-xl cursor-pointer bg-slate-50 hover:bg-slate-100/70 transition-colors group">
              {!file3 ? (
                <div className="flex flex-col items-center justify-center pt-4 pb-4">
                  <Upload className="w-8 h-8 text-slate-400 group-hover:text-slate-600 mb-1.5 transition-colors" />
                  <p className="text-xs font-bold text-slate-600">Pilih Data 3 Discontinued</p>
                  <p className="text-[10px] text-slate-400 mt-0.5">XLSX, XLS, CSV</p>
                </div>
              ) : (
                <div className="flex flex-col items-center justify-center w-full h-full bg-emerald-50 rounded-xl border border-emerald-100 px-4 text-center">
                  <CheckCircle2 className="w-8 h-8 text-emerald-500 mb-1.5" />
                  <p className="text-xs font-bold text-emerald-700 truncate w-full">{file3.name}</p>
                  <button
                    type="button"
                    onClick={(e) => {
                      e.preventDefault();
                      setFile3(null);
                    }}
                    className="text-[10px] text-rose-500 hover:underline mt-1 font-bold"
                  >
                    Hapus
                  </button>
                </div>
              )}
              <input
                type="file"
                className="hidden"
                accept=".xlsx, .xls, .csv"
                onChange={(e) => {
                  const f = e.target.files?.[0];
                  if (f) setFile3(f);
                }}
              />
            </label>
          </div>
        </div>

        {/* Summary Metric Counters */}
        {hasProcessed && (
          <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-4 mb-6">
            
            <button
              onClick={() => {
                setActiveTab('mismatch');
                setCurrentPage(1);
              }}
              className={`p-4 rounded-2xl cursor-pointer text-left flex flex-col gap-2 transition-all duration-300 border ${
                activeTab === 'mismatch'
                  ? 'bg-rose-50 border-rose-300 shadow-sm'
                  : 'bg-white border-slate-200 hover:border-slate-300'
              }`}
            >
              <div className="flex items-center justify-between w-full">
                <div className="p-2 bg-rose-100 rounded-lg text-rose-500">
                  <GitCompare className="w-5 h-5" />
                </div>
                <h4 className="text-2xl font-black text-slate-800">
                  {comparisonResults.mismatch.length}
                </h4>
              </div>
              <p className="text-[10px] text-slate-500 font-extrabold uppercase tracking-wider">
                Selisih Stok
              </p>
            </button>

            <button
              onClick={() => {
                setActiveTab('match');
                setCurrentPage(1);
              }}
              className={`p-4 rounded-2xl cursor-pointer text-left flex flex-col gap-2 transition-all duration-300 border ${
                activeTab === 'match'
                  ? 'bg-emerald-50 border-emerald-300 shadow-sm'
                  : 'bg-white border-slate-200 hover:border-slate-300'
              }`}
            >
              <div className="flex items-center justify-between w-full">
                <div className="p-2 bg-emerald-100 rounded-lg text-emerald-600">
                  <CheckSquare className="w-5 h-5" />
                </div>
                <h4 className="text-2xl font-black text-slate-800">
                  {comparisonResults.match.length}
                </h4>
              </div>
              <p className="text-[10px] text-slate-500 font-extrabold uppercase tracking-wider">
                Stok Sama
              </p>
            </button>

            <button
              onClick={() => {
                setActiveTab('only1');
                setCurrentPage(1);
              }}
              className={`p-4 rounded-2xl cursor-pointer text-left flex flex-col gap-2 transition-all duration-300 border ${
                activeTab === 'only1'
                  ? 'bg-slate-100 border-slate-300 shadow-sm'
                  : 'bg-white border-slate-200 hover:border-slate-300'
              }`}
            >
              <div className="flex items-center justify-between w-full">
                <div className="p-2 bg-slate-100 rounded-lg text-slate-500">
                  <Database className="w-5 h-5" />
                </div>
                <h4 className="text-2xl font-black text-slate-800">
                  {comparisonResults.only1.length}
                </h4>
              </div>
              <p className="text-[10px] text-slate-500 font-extrabold uppercase tracking-wider">
                Hanya Data 1
              </p>
            </button>

            <button
              onClick={() => {
                setActiveTab('only2');
                setCurrentPage(1);
              }}
              className={`p-4 rounded-2xl cursor-pointer text-left flex flex-col gap-2 transition-all duration-300 border ${
                activeTab === 'only2'
                  ? 'bg-indigo-50 border-indigo-300 shadow-sm'
                  : 'bg-white border-slate-200 hover:border-slate-300'
              }`}
            >
              <div className="flex items-center justify-between w-full">
                <div className="p-2 bg-indigo-100 rounded-lg text-indigo-500">
                  <Upload className="w-5 h-5" />
                </div>
                <h4 className="text-2xl font-black text-slate-800">
                  {comparisonResults.only2.length}
                </h4>
              </div>
              <p className="text-[10px] text-slate-500 font-extrabold uppercase tracking-wider">
                Belum Upload
              </p>
            </button>

            <button
              onClick={() => {
                setActiveTab('discontinued');
                setCurrentPage(1);
              }}
              className={`p-4 rounded-2xl cursor-pointer text-left flex flex-col gap-2 transition-all duration-300 border ${
                activeTab === 'discontinued'
                  ? 'bg-amber-50 border-amber-300 shadow-sm'
                  : 'bg-white border-slate-200 hover:border-slate-300'
              }`}
            >
              <div className="flex items-center justify-between w-full">
                <div className="p-2 bg-amber-100 rounded-lg text-amber-500">
                  <TrendingDown className="w-5 h-5" />
                </div>
                <h4 className="text-2xl font-black text-slate-800">
                  {comparisonResults.discontinued.length}
                </h4>
              </div>
              <p className="text-[10px] text-slate-500 font-extrabold uppercase tracking-wider">
                Discontinued
              </p>
            </button>

            <button
              onClick={() => {
                setActiveTab('shortSku');
                setCurrentPage(1);
              }}
              className={`p-4 rounded-2xl cursor-pointer text-left flex flex-col gap-2 transition-all duration-300 border ${
                activeTab === 'shortSku'
                  ? 'bg-violet-50 border-violet-300 shadow-sm'
                  : 'bg-white border-slate-200 hover:border-slate-300'
              }`}
            >
              <div className="flex items-center justify-between w-full">
                <div className="p-2 bg-violet-100 rounded-lg text-violet-600">
                  <AlertCircle className="w-5 h-5" />
                </div>
                <h4 className="text-2xl font-black text-slate-800">
                  {comparisonResults.shortSku?.length || 0}
                </h4>
              </div>
              <p className="text-[10px] text-slate-500 font-extrabold uppercase tracking-wider">
                SKU &lt; 5 Digit
              </p>
            </button>
          </div>
        )}

        {/* Results Container Block */}
        {hasProcessed && (
          <div className="bg-white border border-slate-200 rounded-2xl overflow-hidden shadow-sm transition-all duration-500">
            
            {/* Filter & Search Bar */}
            <div className="p-4 border-b border-slate-200 bg-slate-50/80 flex flex-col lg:flex-row justify-between items-start lg:items-center gap-4">
              <div className="flex flex-wrap items-center gap-3 w-full lg:w-auto">
                <div className="relative w-full sm:w-72 text-slate-600">
                  <Search className="w-4.5 h-4.5 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => {
                      setSearchQuery(e.target.value);
                      setCurrentPage(1);
                    }}
                    placeholder="Cari SKU atau Nama..."
                    className="w-full pl-9 pr-4 py-2 bg-white border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-slate-200 text-sm shadow-sm"
                  />
                </div>

                <button
                  onClick={handleExportToExcel}
                  className="bg-emerald-50 text-emerald-600 border border-emerald-200 hover:bg-emerald-100/70 px-4 py-2 rounded-lg text-xs font-bold shadow-sm transition-colors flex items-center gap-1.5 shrink-0"
                >
                  <FileSpreadsheet className="w-4 h-4" />
                  <span>Download Excel</span>
                </button>

                {activeTab === 'shortSku' && (
                  <button
                    onClick={() => {
                      setModalSearchQuery('');
                      setShowFixSkuModal(true);
                    }}
                    className="bg-violet-600 hover:bg-violet-700 text-white border border-violet-700 px-4 py-2 rounded-lg text-xs font-black shadow-md hover:shadow-lg transition-all flex items-center gap-1.5 shrink-0 transform hover:-translate-y-0.5 active:translate-y-0"
                  >
                    <Sparkles className="w-4 h-4 text-violet-200 animate-pulse" />
                    <span>Terapkan Perbaikan</span>
                  </button>
                )}
              </div>

              {/* Dynamic Filter Controls */}
              <div className="flex flex-wrap items-center gap-3 w-full lg:w-auto">
                {activeTab === 'mismatch' && (
                  <div className="flex flex-wrap items-center gap-2 w-full sm:w-auto">
                    <input
                      type="number"
                      value={min1}
                      onChange={(e) => {
                        setMin1(e.target.value);
                        setCurrentPage(1);
                      }}
                      placeholder="Min D1"
                      className="w-20 px-2.5 py-1.5 border border-slate-300 rounded-lg text-xs font-semibold bg-white outline-none focus:border-indigo-400"
                    />
                    <input
                      type="number"
                      value={max1}
                      onChange={(e) => {
                        setMax1(e.target.value);
                        setCurrentPage(1);
                      }}
                      placeholder="Max D1"
                      className="w-20 px-2.5 py-1.5 border border-slate-300 rounded-lg text-xs font-semibold bg-white outline-none focus:border-indigo-400"
                    />
                    <input
                      type="number"
                      value={minAio}
                      onChange={(e) => {
                        setMinAio(e.target.value);
                        setCurrentPage(1);
                      }}
                      placeholder="Min AIO"
                      className="w-22 px-2.5 py-1.5 border border-slate-300 rounded-lg text-xs font-semibold bg-white outline-none focus:border-indigo-400"
                    />
                    <input
                      type="number"
                      value={maxAio}
                      onChange={(e) => {
                        setMaxAio(e.target.value);
                        setCurrentPage(1);
                      }}
                      placeholder="Max AIO"
                      className="w-22 px-2.5 py-1.5 border border-slate-300 rounded-lg text-xs font-semibold bg-white outline-none focus:border-indigo-400"
                    />
                  </div>
                )}

                {activeTab === 'only2' && (
                  <div className="flex flex-wrap items-center gap-2 w-full sm:w-auto">
                    <select
                      value={categoryFilter}
                      onChange={(e) => {
                        setCategoryFilter(e.target.value);
                        setCurrentPage(1);
                      }}
                      className="w-36 px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-semibold cursor-pointer outline-none text-slate-700"
                    >
                      <option value="">Semua Kategori</option>
                      {getCategoryOptions().map(opt => (
                        <option key={opt.value} value={opt.value}>{opt.label}</option>
                      ))}
                    </select>

                    <select
                      value={brandFilter}
                      onChange={(e) => {
                        setBrandFilter(e.target.value);
                        setCurrentPage(1);
                      }}
                      className="w-36 px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-semibold cursor-pointer outline-none text-slate-700"
                    >
                      <option value="">Semua Merk</option>
                      {getBrandOptions().map(opt => (
                        <option key={opt.value} value={opt.value}>{opt.label}</option>
                      ))}
                    </select>

                    <select
                      value={sortOnly2}
                      onChange={(e) => {
                        setSortOnly2(e.target.value);
                        setCurrentPage(1);
                      }}
                      className="w-36 px-2.5 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-semibold cursor-pointer outline-none text-slate-700"
                    >
                      <option value="">Urutkan Stok</option>
                      <option value="desc">Tinggi - Rendah</option>
                      <option value="asc">Rendah - Tinggi</option>
                    </select>

                    <div className="bg-indigo-100 text-indigo-700 px-3 py-1.5 rounded-lg border border-indigo-200 text-xs font-bold">
                      {filteredData.length} Tampil
                    </div>
                  </div>
                )}
              </div>
            </div>

            {/* Table Area */}
            <div className="overflow-x-auto bg-white">
              <table className="w-full text-left border-collapse">
                <thead className="bg-slate-50 text-[11px] uppercase font-extrabold text-slate-500 tracking-wider border-b border-slate-200">
                  <tr>
                    <th className="px-5 py-3 font-bold text-slate-600">Kode SKU</th>
                    {activeTab !== 'only1' && (
                      <th className="px-5 py-3 font-bold text-slate-600">Nama Produk</th>
                    )}
                    {activeTab === 'mismatch' && (
                      <>
                        <th className="px-5 py-3 font-bold text-center">Data 1</th>
                        <th className="px-5 py-3 font-bold text-center">AIO</th>
                      </>
                    )}
                    {activeTab === 'match' && (
                      <>
                        <th className="px-5 py-3 font-bold text-center">Data 1</th>
                        <th className="px-5 py-3 font-bold text-center">AIO</th>
                      </>
                    )}
                    {activeTab === 'only1' && (
                      <th className="px-5 py-3 font-bold text-center">Stok D1</th>
                    )}
                    {activeTab === 'only2' && (
                      <th className="px-5 py-3 font-bold text-center">Stok AIO</th>
                    )}
                    {activeTab === 'discontinued' && (
                      <>
                        <th className="px-5 py-3 font-bold text-center">Stok D1</th>
                        <th className="px-5 py-3 font-bold text-center">Stok D3</th>
                      </>
                    )}
                    {activeTab === 'shortSku' && (
                      <>
                        <th className="px-5 py-3 font-bold text-center">SKU Baru (5 Digit)</th>
                        <th className="px-5 py-3 font-bold text-center">Sumber</th>
                        <th className="px-5 py-3 font-bold text-center">Stok</th>
                      </>
                    )}
                    <th className="px-5 py-3 font-bold text-right">Status / Aksi</th>
                  </tr>
                </thead>
                <tbody className="text-sm divide-y divide-slate-100">
                  {paginatedData.length > 0 ? (
                    paginatedData.map((item, index) => {
                      return (
                        <tr key={item.sku + "-" + index} className="hover:bg-slate-50/70 transition-colors">
                          {/* SKU Cell */}
                          <td className="px-5 py-4 align-top">
                            <div className="flex items-center gap-2 font-bold text-slate-700">
                              <span>{item.sku}</span>
                              <button
                                onClick={() => copyToClipboard(item.sku, 'SKU')}
                                className="text-slate-300 hover:text-indigo-500 transition hover:scale-110"
                                title="Salin SKU"
                              >
                                <Copy className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          </td>

                          {/* Product Name Cell (not in only1) */}
                          {activeTab !== 'only1' && (
                            <td className="px-5 py-4 min-w-[280px] align-top">
                              <div className="flex items-start gap-2 group">
                                <div className="text-slate-700 font-semibold text-sm leading-relaxed">
                                  {item.name}
                                </div>
                                <button
                                  onClick={() => copyToClipboard(item.name, 'Nama Produk')}
                                  className="text-slate-300 opacity-0 group-hover:opacity-100 hover:text-indigo-500 transition hover:scale-110"
                                  title="Salin Nama"
                                >
                                  <Copy className="w-3.5 h-3.5" />
                                </button>
                              </div>

                              {activeTab === 'only2' && (
                                <div className="text-[10px] text-slate-500 mt-1.5 flex flex-wrap gap-1.5">
                                  <span className="bg-slate-100 px-2 py-0.5 rounded border border-slate-200">
                                    {item.brand || '-'}
                                  </span>
                                  <span className="bg-slate-100 px-2 py-0.5 rounded border border-slate-200">
                                    {item.category || '-'}
                                  </span>
                                </div>
                              )}
                            </td>
                          )}

                          {/* Stocks Display Cells */}
                          {activeTab === 'mismatch' && (
                            <>
                              <td className="px-5 py-4 text-center align-top">
                                <span className="font-mono font-bold bg-slate-100 px-2.5 py-1 rounded-lg border border-slate-200 text-xs">
                                  {item.stock1}
                                </span>
                              </td>
                              <td className="px-5 py-4 text-center align-top">
                                <span className="font-mono font-bold bg-slate-100 px-2.5 py-1 rounded-lg border border-slate-200 text-xs">
                                  {item.stockAio}
                                </span>
                              </td>
                            </>
                          )}

                          {activeTab === 'match' && (
                            <>
                              <td className="px-5 py-4 text-center align-top">
                                <span className="font-mono font-bold bg-slate-100 px-2.5 py-1 rounded-lg border border-slate-200 text-xs">
                                  {item.stock1}
                                </span>
                              </td>
                              <td className="px-5 py-4 text-center align-top">
                                <span className="font-mono font-bold bg-slate-100 px-2.5 py-1 rounded-lg border border-slate-200 text-xs">
                                  {item.stockAio}
                                </span>
                              </td>
                            </>
                          )}

                          {activeTab === 'only1' && (
                            <td className="px-5 py-4 text-center align-top">
                              <span className="font-mono font-bold bg-slate-100 px-2.5 py-1 rounded-lg border border-slate-200 text-xs">
                                {item.stock1}
                              </span>
                            </td>
                          )}

                          {activeTab === 'only2' && (
                            <td className="px-5 py-4 text-center align-top">
                              <span className="font-mono font-bold bg-slate-100 px-2.5 py-1 rounded-lg border border-slate-200 text-xs">
                                {item.stockAio}
                              </span>
                            </td>
                          )}

                          {activeTab === 'discontinued' && (
                            <>
                              <td className="px-5 py-4 text-center align-top">
                                <span className="font-mono font-bold bg-slate-100 px-2.5 py-1 rounded-lg border border-slate-200 text-xs">
                                  {item.stock1}
                                </span>
                              </td>
                              <td className="px-5 py-4 text-center align-top">
                                <span className="font-mono font-bold bg-slate-100 px-2.5 py-1 rounded-lg border border-slate-200 text-xs">
                                  {item.stock3}
                                </span>
                              </td>
                            </>
                          )}

                          {activeTab === 'shortSku' && (
                            <>
                              <td className="px-5 py-4 text-center align-top">
                                <div className="flex items-center justify-center gap-2 font-mono font-bold text-indigo-600 bg-indigo-50 px-2.5 py-1 rounded-lg border border-indigo-200 text-xs w-fit mx-auto">
                                  <span>{item.suggestedSku}</span>
                                  <button
                                    onClick={() => copyToClipboard(item.suggestedSku, 'SKU Baru')}
                                    className="text-indigo-400 hover:text-indigo-700 transition hover:scale-110"
                                    title="Salin SKU Baru"
                                  >
                                    <Copy className="w-3.5 h-3.5" />
                                  </button>
                                </div>
                              </td>
                              <td className="px-5 py-4 text-center align-top">
                                <span className="text-xs font-semibold bg-slate-100 border border-slate-200 px-2 py-1 rounded text-slate-600">
                                  {item.source}
                                </span>
                              </td>
                              <td className="px-5 py-4 text-center align-top">
                                <span className="font-mono font-bold bg-slate-100 px-2.5 py-1 rounded-lg border border-slate-200 text-xs">
                                  {item.stock}
                                </span>
                              </td>
                            </>
                          )}

                          {/* Action Cell */}
                          <td className="px-5 py-4 text-right align-top">
                            {activeTab === 'mismatch' && (
                              <span className="inline-flex items-center px-2.5 py-1 bg-rose-100 text-rose-600 rounded text-xs font-bold border border-rose-200">
                                Selisih
                              </span>
                            )}
                            {activeTab === 'match' && (
                              <span className="inline-flex items-center px-2.5 py-1 bg-emerald-100 text-emerald-600 rounded text-xs font-bold border border-emerald-200">
                                Sama
                              </span>
                            )}
                            {activeTab === 'only1' && (
                              <span className="inline-flex items-center px-2.5 py-1 bg-slate-100 text-slate-600 rounded text-xs font-bold border border-slate-200">
                                Data 1 Saja
                              </span>
                            )}

                            {activeTab === 'only2' && (
                              <div className="flex justify-end">
                                {savedUploadData[item.sku] ? (
                                  <button
                                    onClick={() => handleToggleAction(item.sku, 'up')}
                                    className="group flex items-center justify-center gap-1.5 bg-emerald-100 hover:bg-rose-100 text-emerald-700 hover:text-rose-600 border border-emerald-200 px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all w-36"
                                  >
                                    <span className="group-hover:hidden flex items-center gap-1.5">
                                      <CheckSquare className="w-3.5 h-3.5" /> Sudah Terupload
                                    </span>
                                    <span className="hidden group-hover:flex items-center gap-1.5">
                                      <X className="w-3.5 h-3.5" /> Batal
                                    </span>
                                  </button>
                                ) : selectedForMission.includes(item.sku) ? (
                                  <div className="flex items-center">
                                    <button
                                      onClick={() => handleExecuteAction(item.sku, 'up')}
                                      className="bg-indigo-600 hover:bg-indigo-700 text-white px-3.5 py-1.5 rounded-l-lg text-xs font-bold flex items-center gap-1"
                                    >
                                      <Upload className="w-3.5 h-3.5" /> Upload!
                                    </button>
                                    <button
                                      onClick={() => handleQueueAction(item.sku, 'up')}
                                      className="bg-white hover:bg-rose-50 text-slate-500 px-2 py-1.5 rounded-r-lg border-y border-r border-slate-300"
                                      title="Batalkan Target"
                                    >
                                      <X className="w-3.5 h-3.5" />
                                    </button>
                                  </div>
                                ) : (
                                  <button
                                    onClick={() => handleQueueAction(item.sku, 'up')}
                                    className="bg-white hover:bg-slate-50 text-slate-600 border border-slate-300 px-3 py-1.5 rounded-lg text-xs font-bold transition-colors"
                                  >
                                    Tambah Target
                                  </button>
                                )}
                              </div>
                            )}

                            {activeTab === 'discontinued' && (
                              <div className="flex justify-end">
                                {savedDiscData[item.sku] ? (
                                  <button
                                    onClick={() => handleToggleAction(item.sku, 'disc')}
                                    className="group flex items-center justify-center gap-1.5 bg-amber-100 hover:bg-rose-100 text-amber-700 hover:text-rose-600 border border-amber-200 px-3.5 py-1.5 rounded-lg text-xs font-bold transition-all w-36"
                                  >
                                    <span className="group-hover:hidden flex items-center gap-1.5">
                                      <CheckSquare className="w-3.5 h-3.5" /> Sudah Diproses
                                    </span>
                                    <span className="hidden group-hover:flex items-center gap-1.5">
                                      <X className="w-3.5 h-3.5" /> Batal
                                    </span>
                                  </button>
                                ) : selectedForMissionDisc.includes(item.sku) ? (
                                  <div className="flex items-center">
                                    <button
                                      onClick={() => handleExecuteAction(item.sku, 'disc')}
                                      className="bg-amber-500 hover:bg-amber-600 text-white px-3.5 py-1.5 rounded-l-lg text-xs font-bold flex items-center gap-1"
                                    >
                                      <Archive className="w-3.5 h-3.5" /> Arsip!
                                    </button>
                                    <button
                                      onClick={() => handleQueueAction(item.sku, 'disc')}
                                      className="bg-white hover:bg-rose-50 text-slate-500 px-2 py-1.5 rounded-r-lg border-y border-r border-slate-300"
                                      title="Batalkan Target"
                                    >
                                      <X className="w-3.5 h-3.5" />
                                    </button>
                                  </div>
                                ) : (
                                  <button
                                    onClick={() => handleQueueAction(item.sku, 'disc')}
                                    className="bg-white hover:bg-slate-50 text-slate-600 border border-slate-300 px-3 py-1.5 rounded-lg text-xs font-bold transition-colors"
                                  >
                                    Proses Data
                                  </button>
                                )}
                              </div>
                            )}

                            {activeTab === 'shortSku' && (
                              <div className="flex justify-end gap-2 items-center">
                                {item.existsInAio ? (
                                  <span className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-emerald-50 text-emerald-700 rounded-lg text-xs font-bold border border-emerald-200 shadow-sm">
                                    <CheckCircle2 className="w-3.5 h-3.5" /> Terdaftar di AIO
                                  </span>
                                ) : (
                                  <span className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-rose-50 text-rose-700 rounded-lg text-xs font-bold border border-rose-200 shadow-sm">
                                    <AlertCircle className="w-3.5 h-3.5" /> Belum di AIO
                                  </span>
                                )}
                                <button
                                  onClick={() => copyToClipboard(item.suggestedSku, 'SKU Baru')}
                                  className="bg-indigo-600 hover:bg-indigo-700 text-white px-3.5 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 shadow-sm transition-colors"
                                >
                                  <Copy className="w-3.5 h-3.5" /> Salin SKU
                                </button>
                              </div>
                            )}
                          </td>
                        </tr>
                      );
                    })
                  ) : (
                    <tr>
                      <td colSpan={10} className="py-16 text-center">
                        <Ghost className="w-10 h-10 text-slate-300 mx-auto mb-2" />
                        <p className="text-slate-700 font-bold">Tidak ada data ditemukan.</p>
                        <p className="text-slate-400 text-xs mt-0.5">Coba sesuaikan kata kunci pencarian atau filter Anda.</p>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination Controls */}
            {totalPages > 1 && (
              <div className="p-4 border-t border-slate-200 bg-slate-50 flex items-center justify-between">
                <p className="text-xs text-slate-500">
                  Tampil <span className="font-extrabold text-slate-700">{(currentPage - 1) * itemsPerPage + 1}</span> - <span className="font-extrabold text-slate-700">{Math.min(currentPage * itemsPerPage, totalItems)}</span> dari <span className="font-extrabold text-slate-700">{totalItems}</span>
                </p>
                <div className="flex gap-2">
                  <button
                    onClick={handlePrevPage}
                    disabled={currentPage <= 1}
                    className="px-3.5 py-2 bg-white border border-slate-300 rounded-lg text-xs font-bold transition-colors shadow-sm disabled:opacity-50 disabled:cursor-not-allowed hover:bg-slate-50"
                  >
                    <ChevronLeft className="w-4 h-4 inline" /> Sebelumnya
                  </button>
                  <button
                    onClick={handleNextPage}
                    disabled={currentPage >= totalPages}
                    className="px-3.5 py-2 bg-white border border-slate-300 rounded-lg text-xs font-bold transition-colors shadow-sm disabled:opacity-50 disabled:cursor-not-allowed hover:bg-slate-50"
                  >
                    Selanjutnya <ChevronRight className="w-4 h-4 inline" />
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Mission Bento Panel Widget (Only visible in only2 or discontinued tabs) */}
        {hasProcessed && (activeTab === 'only2' || activeTab === 'discontinued') && (
          <div className="fixed bottom-6 left-1/2 -translate-x-1/2 bg-white border border-slate-200 p-4 rounded-2xl shadow-xl z-40 flex items-center gap-4 transition-all duration-300 animate-slide-up max-w-[95%] sm:max-w-none">
            <div className="flex items-center gap-2.5">
              <div className={`p-2.5 rounded-xl text-white shadow-sm ${
                activeTab === 'only2' ? 'bg-indigo-500' : 'bg-amber-500'
              }`}>
                <Target className="w-5 h-5" />
              </div>
              <div className="flex flex-col ml-1 hidden sm:flex">
                <span className="text-[10px] text-slate-400 uppercase font-extrabold tracking-wider leading-none mb-1">
                  Misi Harian
                </span>
                <span className="text-xs font-extrabold text-slate-700 leading-none">
                  {activeTab === 'only2' ? 'Target Upload' : 'Target Arsip'}
                </span>
              </div>
            </div>

            <div className="h-10 w-px bg-slate-200"></div>

            <div className="flex flex-col items-center">
              <div className="flex items-center gap-1.5">
                {Array.from({ length: dailyMissionTarget }).map((_, i) => {
                  const filled = i < activeMissionProgress;
                  return (
                    <div
                      key={i}
                      className={`w-8 h-8 rounded-lg flex items-center justify-center transition-all duration-300 ${
                        filled
                          ? (activeTab === 'only2' ? 'bg-indigo-500 text-yellow-300 scale-105 shadow-sm' : 'bg-amber-500 text-yellow-300 scale-105 shadow-sm')
                          : 'bg-slate-100 border border-slate-200 text-slate-300 shadow-inner'
                      }`}
                    >
                      <Star className={`w-4 h-4 ${filled ? 'fill-current' : ''}`} />
                    </div>
                  );
                })}
              </div>

              <div className={`text-[10px] font-extrabold tracking-wide px-2.5 py-0.5 rounded mt-2 border ${
                activeTab === 'only2'
                  ? (lastDurationRecord !== "-" ? 'bg-indigo-50 text-indigo-600 border-indigo-200' : 'bg-slate-50 text-slate-400 border-slate-200')
                  : (lastDurationRecordDisc !== "-" ? 'bg-amber-50 text-amber-600 border-amber-200' : 'bg-slate-50 text-slate-400 border-slate-200')
              }`}>
                {activeTab === 'only2' ? lastDurationRecord : lastDurationRecordDisc}
              </div>
            </div>

            <div className="bg-slate-50 px-3 py-2 rounded-xl border border-slate-200 text-center shadow-inner self-center flex items-center justify-center gap-1 min-w-[70px]">
              <span className={`text-xl font-extrabold ${activeTab === 'only2' ? 'text-indigo-600' : 'text-amber-600'}`}>
                {activeMissionCount}
              </span>
              <span className="text-slate-400 font-bold text-xs">/ 5</span>
            </div>
          </div>
        )}

        {/* SKU Fixer Interactive Modal */}
        {showFixSkuModal && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-fade-in">
            <div className="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-4xl w-full max-h-[85vh] flex flex-col overflow-hidden animate-scale-up">
              
              {/* Modal Header */}
              <div className="p-6 bg-gradient-to-r from-violet-600 to-indigo-700 text-white flex justify-between items-center shrink-0">
                <div className="flex items-center gap-3">
                  <div className="p-2 bg-white/20 rounded-xl text-white backdrop-blur-xs">
                    <Sparkles className="w-6 h-6 animate-pulse" />
                  </div>
                  <div>
                    <h3 className="text-lg font-black tracking-tight">Rekomendasi Perbaikan Format SKU</h3>
                    <p className="text-xs text-violet-100 mt-0.5">Otomatis menormalisasi SKU kurang dari 5 digit dengan tambahan nol di depan</p>
                  </div>
                </div>
                <button
                  onClick={() => setShowFixSkuModal(false)}
                  className="p-1.5 rounded-xl hover:bg-white/15 text-white transition-colors cursor-pointer"
                  aria-label="Tutup"
                >
                  <X className="w-6 h-6" />
                </button>
              </div>

              {/* Action and Search Panel */}
              <div className="p-5 border-b border-slate-200 bg-slate-50 flex flex-col md:flex-row items-center justify-between gap-4 shrink-0">
                <div className="relative w-full md:w-72 text-slate-600">
                  <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
                  <input
                    type="text"
                    value={modalSearchQuery}
                    onChange={(e) => setModalSearchQuery(e.target.value)}
                    placeholder="Cari SKU asli atau baru..."
                    className="w-full pl-9 pr-4 py-2 bg-white border border-slate-300 rounded-lg text-xs font-medium focus:outline-none focus:ring-2 focus:ring-violet-200 shadow-sm"
                  />
                </div>

                <div className="flex items-center gap-3 w-full md:w-auto">
                  <button
                    onClick={copyAllFixedSkus}
                    className="flex-1 md:flex-none bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 px-4 py-2 rounded-xl text-xs font-bold shadow-sm transition-colors flex items-center justify-center gap-1.5 cursor-pointer"
                  >
                    <Copy className="w-4 h-4 text-slate-500" />
                    <span>Salin Semua SKU Baru</span>
                  </button>

                  <button
                    onClick={handleDownloadSkuFixes}
                    className="flex-1 md:flex-none bg-violet-600 hover:bg-violet-700 text-white px-4 py-2 rounded-xl text-xs font-extrabold shadow-sm hover:shadow-md transition-all flex items-center justify-center gap-1.5 cursor-pointer"
                  >
                    <FileSpreadsheet className="w-4 h-4 text-violet-200" />
                    <span>Unduh File Perbaikan</span>
                  </button>
                </div>
              </div>

              {/* Modal Table Content */}
              <div className="flex-1 overflow-y-auto p-5">
                <div className="border border-slate-200 rounded-2xl overflow-hidden bg-white shadow-inner">
                  <table className="w-full text-left border-collapse">
                    <thead className="bg-slate-50 text-[11px] uppercase font-extrabold text-slate-500 tracking-wider border-b border-slate-200">
                      <tr>
                        <th className="px-5 py-3 font-bold text-slate-600">SKU Asli</th>
                        <th className="px-5 py-3 font-bold text-center text-slate-600">Rekomendasi SKU (5 Digit)</th>
                        <th className="px-5 py-3 font-bold text-slate-600">Nama Produk</th>
                        <th className="px-5 py-3 font-bold text-center text-slate-600">Sumber</th>
                        <th className="px-5 py-3 font-bold text-center text-slate-600">Stok</th>
                        <th className="px-5 py-3 font-bold text-right text-slate-600">Status</th>
                      </tr>
                    </thead>
                    <tbody className="text-xs divide-y divide-slate-100">
                      {(() => {
                        const allFixes = comparisonResults.shortSku || [];
                        const q = modalSearchQuery.toLowerCase().trim();
                        const filteredFixes = q
                          ? allFixes.filter(item => 
                              item.sku.toLowerCase().includes(q) || 
                              item.suggestedSku.toLowerCase().includes(q) || 
                              item.name.toLowerCase().includes(q)
                            )
                          : allFixes;

                        if (filteredFixes.length === 0) {
                          return (
                            <tr>
                              <td colSpan={6} className="py-12 text-center text-slate-400 font-medium">
                                Tidak ada rekomendasi perbaikan SKU yang cocok.
                              </td>
                            </tr>
                          );
                        }

                        return filteredFixes.map((item, idx) => (
                          <tr key={item.sku + "-fix-" + idx} className="hover:bg-slate-50/50 transition-colors">
                            <td className="px-5 py-3.5 font-bold text-slate-700 font-mono">
                              {item.sku}
                            </td>
                            <td className="px-5 py-3.5 text-center">
                              <div className="inline-flex items-center gap-1.5 font-mono font-bold text-violet-600 bg-violet-50 px-2 py-0.5 rounded-md border border-violet-200 w-fit mx-auto">
                                <span>{item.suggestedSku}</span>
                                <button
                                  onClick={() => copyToClipboard(item.suggestedSku, 'SKU Baru')}
                                  className="text-violet-400 hover:text-violet-700 transition cursor-pointer"
                                  title="Salin"
                                >
                                  <Copy className="w-3 h-3" />
                                </button>
                              </div>
                            </td>
                            <td className="px-5 py-3.5 max-w-xs truncate text-slate-600 font-medium" title={item.name}>
                              {item.name}
                            </td>
                            <td className="px-5 py-3.5 text-center">
                              <span className="text-[10px] font-bold bg-slate-100 border border-slate-200 px-1.5 py-0.5 rounded text-slate-500">
                                {item.source}
                              </span>
                            </td>
                            <td className="px-5 py-3.5 text-center font-mono font-semibold text-slate-600">
                              {item.stock}
                            </td>
                            <td className="px-5 py-3.5 text-right">
                              {item.existsInAio ? (
                                <span className="inline-flex items-center gap-1 px-2 py-1 bg-emerald-50 text-emerald-700 rounded-md text-[10px] font-bold border border-emerald-100 shadow-xs">
                                  <CheckCircle2 className="w-3 h-3 text-emerald-600" /> Terdaftar
                                </span>
                              ) : (
                                <span className="inline-flex items-center gap-1 px-2 py-1 bg-rose-50 text-rose-700 rounded-md text-[10px] font-bold border border-rose-100 shadow-xs">
                                  <AlertCircle className="w-3 h-3 text-rose-600" /> Belum Terdaftar
                                </span>
                              )}
                            </td>
                          </tr>
                        ));
                      })()}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Modal Footer */}
              <div className="p-4 bg-slate-50 border-t border-slate-200 flex justify-between items-center shrink-0">
                <span className="text-[11px] text-slate-500 font-semibold">
                  Total SKU diperbaiki: <span className="text-slate-800 font-extrabold">{comparisonResults.shortSku?.length || 0} items</span>
                </span>
                <button
                  onClick={() => setShowFixSkuModal(false)}
                  className="bg-slate-200 hover:bg-slate-300 text-slate-700 px-4 py-2 rounded-xl text-xs font-bold transition-colors cursor-pointer"
                >
                  Tutup
                </button>
              </div>

            </div>
          </div>
        )}

        {/* Global Toast Notification */}
        <div className={`fixed top-6 right-6 bg-white text-slate-800 px-4 py-3 rounded-xl shadow-xl border border-slate-200 flex items-center gap-3 transition-all duration-300 z-50 transform ${
          notification.show ? 'translate-x-0 opacity-100' : 'translate-x-[150%] opacity-0'
        }`}>
          {notification.isWarning ? (
            <AlertCircle className="w-5 h-5 text-rose-500 shrink-0" />
          ) : (
            <CheckCircle2 className="w-5 h-5 text-emerald-500 shrink-0" />
          )}
          <div>
            <h4 className="text-sm font-extrabold text-slate-800 leading-none">{notification.title}</h4>
            <p className="text-xs text-slate-500 mt-1 font-medium leading-relaxed">{notification.text}</p>
          </div>
        </div>

      </div>
    </div>
  );
}
