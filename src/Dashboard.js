import React, { useState, useEffect, useCallback, useMemo } from "react";
import {
  collection,
  query,
  where,
  getDocs,
  limit,
  startAfter,
  orderBy,
  writeBatch,
  doc,
  serverTimestamp,
} from "firebase/firestore";
import { db } from "./firebase"; // Adjust the import path to your firebase config file
import * as XLSX from "xlsx";

/* ──────────────────────  UTILITIES  ────────────────────── */
const getTodayStart = () => {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return today;
};

const extractGB = (desc) =>
  !desc ? "N/A" : (desc.match(/(\d+)GB/)?.[1] ?? "N/A");

const formatPhoneNumber = (number) => {
  if (!number) return "N/A";
  const cleaned = number.replace(/^233/, "");
  return cleaned.length === 9 ? `0${cleaned}` : cleaned || "N/A";
};

const downloadExcel = (data, fileName, headers) => {
  const ws = XLSX.utils.json_to_sheet(data, { header: headers });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, fileName);
};

const debounce = (func, wait) => {
  let timeout;
  return (...args) => {
    clearTimeout(timeout);
    timeout = setTimeout(() => func(...args), wait);
  };
};

/* ──────────────────────────────────────────────────────── */

const Dashboard = () => {
  /* ──────────────────────  STATE  ────────────────────── */
  const [tabValue, setTabValue] = useState(0);
  const [numbers, setNumbers] = useState([]);
  const [transactions, setTransactions] = useState([]);
  const [ussdTransactions, setUssdTransactions] = useState([]);
  const [kaditoTransactions, setKaditoTransactions] = useState([]);

  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  // Pagination states
  const [numbersPage, setNumbersPage] = useState(1);
  const [transactionsPage, setTransactionsPage] = useState(1);
  const [ussdPage, setUssdPage] = useState(1);
  const [kaditoPage, setKaditoPage] = useState(1);

  const [numbersLastDocs, setNumbersLastDocs] = useState([]);
  const [transactionsLastDocs, setTransactionsLastDocs] = useState([]);
  const [ussdLastDocs, setUssdLastDocs] = useState([]);
  const [kaditoLastDocs, setKaditoLastDocs] = useState([]);

  const [hasMoreNumbers, setHasMoreNumbers] = useState(true);
  const [hasMoreTransactions, setHasMoreTransactions] = useState(true);
  const [hasMoreUssd, setHasMoreUssd] = useState(true);
  const [hasMoreKadito, setHasMoreKadito] = useState(true);

  // Confirm dialog
  const [showConfirmDialog, setShowConfirmDialog] = useState(false);
  const [confirmAction, setConfirmAction] = useState(null);
  const [recordCount, setRecordCount] = useState(0);

  // Caches
  const [numbersCache, setNumbersCache] = useState({});
  const [transactionsCache, setTransactionsCache] = useState({});
  const [ussdCache, setUssdCache] = useState({});
  const [kaditoCache, setKaditoCache] = useState({});

  // Totals
  const [totalNumbers, setTotalNumbers] = useState(0);
  const [totalTransactions, setTotalTransactions] = useState(0);
  const [totalUssd, setTotalUssd] = useState(0);
  const [totalKadito, setTotalKadito] = useState(0);

  // ─── Bundles Tab State ───
  const [bundlesTabValue, setBundlesTabValue] = useState(0); // 0 = mtn, 1 = tigo, 2 = telecel
  const [bundlesData, setBundlesData] = useState({
    mtn: { daily: [], weekly: [], monthly: [] },
    tigo: { daily: [], weekly: [], monthly: [] },
    telecel: { daily: [], weekly: [], monthly: [] },
  });
  const [bundlesLoading, setBundlesLoading] = useState(false);
  const [bundlesError, setBundlesError] = useState(null);
  const [priceChanges, setPriceChanges] = useState({});
  const [activeChanges, setActiveChanges] = useState({});

  const pageSize = 6;
  const maxExportRecords = 1000;
  const batchSize = 500;

  /* ──────────────────────  TAB HANDLING  ────────────────────── */
  const handleTabChange = useCallback((newValue) => {
    setTabValue(newValue);
    if (newValue === 0) {
      setNumbersPage(1);
      setNumbersLastDocs([]);
      setHasMoreNumbers(true);
    } else if (newValue === 1) {
      setTransactionsPage(1);
      setTransactionsLastDocs([]);
      setHasMoreTransactions(true);
    } else if (newValue === 2) {
      setUssdPage(1);
      setUssdLastDocs([]);
      setHasMoreUssd(true);
    } else if (newValue === 3) {
      setKaditoPage(1);
      setKaditoLastDocs([]);
      setHasMoreKadito(true);
    } else if (newValue === 4) {
      fetchBundles();
    }
    setError(null);
  }, []);

  /* ──────────────────────  TOTAL COUNTS  ────────────────────── */
  const fetchTotalNumbers = useCallback(async () => {
    try {
      const q = query(
        collection(db, "entries"),
        where("exported", "==", false),
      );
      const snap = await getDocs(q);
      setTotalNumbers(snap.size);
    } catch (e) {
      setError(`Total numbers: ${e.message}`);
    }
  }, []);

  const fetchTotalTransactions = useCallback(async () => {
    try {
      const today = getTodayStart();
      const q = query(
        collection(db, "approve_teller_transaction"),
        where("createdAt", ">=", today),
        where("status", "==", "approved"),
        where("exported", "==", false),
      );
      const snap = await getDocs(q);
      setTotalTransactions(snap.size);
    } catch (e) {
      setError(`Total transactions: ${e.message}`);
    }
  }, []);

  const fetchTotalUssd = useCallback(async () => {
    try {
      const today = getTodayStart();
      const q = query(
        collection(db, "teller_response"),
        where("createdAt", ">=", today),
        where("status", "==", "approved"),
        where("exported", "==", false),
      );
      const snap = await getDocs(q);
      setTotalUssd(snap.size);
    } catch (e) {
      setError(`Total USSD: ${e.message}`);
    }
  }, []);

  const fetchTotalKadito = useCallback(async () => {
    try {
      const today = getTodayStart();
      const q = query(
        collection(db, "kadis_purchase"),
        where("createdAt", ">=", today),
        where("status", "==", "approved"),
        where("exported", "==", false),
      );
      const snap = await getDocs(q);
      setTotalKadito(snap.size);
    } catch (e) {
      setError(`Total Kadito: ${e.message}`);
    }
  }, []);

  /* ──────────────────────  DATA FETCHERS  ────────────────────── */
  const fetchNumbers = useCallback(
    async (page = 1) => {
      if (numbersCache[page]) {
        setNumbers(numbersCache[page]);
        setLoading(false);
        return;
      }
      try {
        setLoading(true);
        let q = query(
          collection(db, "entries"),
          where("exported", "==", false),
          orderBy("phoneNumber"),
          limit(pageSize),
        );
        if (page > 1 && numbersLastDocs[page - 2]) {
          q = query(
            collection(db, "entries"),
            where("exported", "==", false),
            orderBy("phoneNumber"),
            startAfter(numbersLastDocs[page - 2]),
            limit(pageSize),
          );
        }
        const snap = await getDocs(q);
        const data = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        setNumbers(data);
        setNumbersCache((p) => ({ ...p, [page]: data }));
        if (snap.docs.length > 0) {
          const newLast = [...numbersLastDocs];
          newLast[page - 1] = snap.docs[snap.docs.length - 1];
          setNumbersLastDocs(newLast);
        }
        setHasMoreNumbers(snap.docs.length === pageSize);
        setLoading(false);
      } catch (e) {
        setError(`Numbers: ${e.message}`);
        setLoading(false);
      }
    },
    [numbersCache, numbersLastDocs],
  );

  const fetchTransactions = useCallback(
    async (page = 1) => {
      if (transactionsCache[page]) {
        setTransactions(transactionsCache[page]);
        setLoading(false);
        return;
      }
      try {
        setLoading(true);
        const today = getTodayStart();
        let q = query(
          collection(db, "approve_teller_transaction"),
          where("createdAt", ">=", today),
          where("status", "==", "approved"),
          where("exported", "==", false),
          orderBy("createdAt"),
          limit(pageSize),
        );
        if (page > 1 && transactionsLastDocs[page - 2]) {
          q = query(
            collection(db, "approve_teller_transaction"),
            where("createdAt", ">=", today),
            where("status", "==", "approved"),
            where("exported", "==", false),
            orderBy("createdAt"),
            startAfter(transactionsLastDocs[page - 2]),
            limit(pageSize),
          );
        }
        const snap = await getDocs(q);
        const data = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        setTransactions(data);
        setTransactionsCache((p) => ({ ...p, [page]: data }));
        if (snap.docs.length > 0) {
          const newLast = [...transactionsLastDocs];
          newLast[page - 1] = snap.docs[snap.docs.length - 1];
          setTransactionsLastDocs(newLast);
        }
        setHasMoreTransactions(snap.docs.length === pageSize);
        setLoading(false);
      } catch (e) {
        setError(`Transactions: ${e.message}`);
        setLoading(false);
      }
    },
    [transactionsCache, transactionsLastDocs],
  );

  const fetchUssdTransactions = useCallback(
    async (page = 1) => {
      if (ussdCache[page]) {
        setUssdTransactions(ussdCache[page]);
        setLoading(false);
        return;
      }
      try {
        setLoading(true);
        const today = getTodayStart();
        let q = query(
          collection(db, "teller_response"),
          where("createdAt", ">=", today),
          where("status", "==", "approved"),
          where("exported", "==", false),
          orderBy("createdAt"),
          limit(pageSize),
        );
        if (page > 1 && ussdLastDocs[page - 2]) {
          q = query(
            collection(db, "teller_response"),
            where("createdAt", ">=", today),
            where("status", "==", "approved"),
            where("exported", "==", false),
            orderBy("createdAt"),
            startAfter(ussdLastDocs[page - 2]),
            limit(pageSize),
          );
        }
        const snap = await getDocs(q);
        const data = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        setUssdTransactions(data);
        setUssdCache((p) => ({ ...p, [page]: data }));
        if (snap.docs.length > 0) {
          const newLast = [...ussdLastDocs];
          newLast[page - 1] = snap.docs[snap.docs.length - 1];
          setUssdLastDocs(newLast);
        }
        setHasMoreUssd(snap.docs.length === pageSize);
        setLoading(false);
      } catch (e) {
        setError(`USSD: ${e.message}`);
        setLoading(false);
      }
    },
    [ussdCache, ussdLastDocs],
  );

  const fetchKaditoTransactions = useCallback(
    async (page = 1) => {
      if (kaditoCache[page]) {
        setKaditoTransactions(kaditoCache[page]);
        setLoading(false);
        return;
      }
      try {
        setLoading(true);
        const today = getTodayStart();
        let q = query(
          collection(db, "kadis_purchase"),
          where("createdAt", ">=", today),
          where("status", "==", "approved"),
          where("exported", "==", false),
          orderBy("createdAt"),
          limit(pageSize),
        );
        if (page > 1 && kaditoLastDocs[page - 2]) {
          q = query(
            collection(db, "kadis_purchase"),
            where("createdAt", ">=", today),
            where("status", "==", "approved"),
            where("exported", "==", false),
            orderBy("createdAt"),
            startAfter(kaditoLastDocs[page - 2]),
            limit(pageSize),
          );
        }
        const snap = await getDocs(q);
        const data = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        setKaditoTransactions(data);
        setKaditoCache((p) => ({ ...p, [page]: data }));
        if (snap.docs.length > 0) {
          const newLast = [...kaditoLastDocs];
          newLast[page - 1] = snap.docs[snap.docs.length - 1];
          setKaditoLastDocs(newLast);
        }
        setHasMoreKadito(snap.docs.length === pageSize);
        setLoading(false);
      } catch (e) {
        setError(`Kadito fetch: ${e.message}`);
        setLoading(false);
      }
    },
    [kaditoCache, kaditoLastDocs],
  );

  /* ──────────────────────  Bundles Fetch & Save ────────────────────── */
  const fetchBundles = useCallback(async () => {
    setBundlesLoading(true);
    setBundlesError(null);
    try {
      const networks = ["mtn", "tigo", "telecel"];
      const periods = ["daily", "weekly", "monthly"];
      const newData = {
        mtn: { daily: [], weekly: [], monthly: [] },
        tigo: { daily: [], weekly: [], monthly: [] },
        telecel: { daily: [], weekly: [], monthly: [] },
      };

      for (const network of networks) {
        for (const period of periods) {
          const colRef = collection(db, "bundles", network, period);
          const q = query(colRef, orderBy("price", "asc"));
          const snap = await getDocs(q);
          newData[network][period] = snap.docs.map((d) => ({
            id: d.id,
            ...d.data(),
          }));
        }
      }
      setBundlesData(newData);
    } catch (err) {
      setBundlesError(err.message);
      console.error("Bundles fetch failed:", err);
    } finally {
      setBundlesLoading(false);
    }
  }, []);

  const handlePriceChange = (network, period, planId, value) => {
    const key = `${network}/${period}/${planId}`;
    const numValue = parseFloat(value);
    setPriceChanges((prev) => ({
      ...prev,
      [key]: isNaN(numValue) ? (prev[key] ?? 0) : numValue,
    }));
  };

  const handleActiveToggle = (network, period, planId, current) => {
    const key = `${network}/${period}/${planId}`;
    setActiveChanges((prev) => ({
      ...prev,
      [key]: !current,
    }));
  };

  const saveBundleChanges = async () => {
    if (
      Object.keys(priceChanges).length === 0 &&
      Object.keys(activeChanges).length === 0
    ) {
      alert("No changes to save");
      return;
    }

    if (!window.confirm("Save all price and active status changes?")) return;

    setBundlesLoading(true);
    try {
      const batch = writeBatch(db);

      Object.entries(priceChanges).forEach(([key, newPrice]) => {
        const [network, period, planId] = key.split("/");
        const ref = doc(db, "bundles", network, period, planId);
        batch.update(ref, { price: newPrice, updatedAt: serverTimestamp() });
      });

      Object.entries(activeChanges).forEach(([key, newActive]) => {
        const [network, period, planId] = key.split("/");
        const ref = doc(db, "bundles", network, period, planId);
        batch.update(ref, { active: newActive, updatedAt: serverTimestamp() });
      });

      await batch.commit();
      alert("Changes saved successfully");
      setPriceChanges({});
      setActiveChanges({});
      await fetchBundles();
    } catch (err) {
      alert("Save failed: " + err.message);
      console.error(err);
    } finally {
      setBundlesLoading(false);
    }
  };

  /* ──────────────────────  EFFECT – LOAD DATA  ────────────────────── */
  useEffect(() => {
    if (tabValue === 0) {
      fetchNumbers(numbersPage);
      fetchTotalNumbers();
    } else if (tabValue === 1) {
      fetchTransactions(transactionsPage);
      fetchTotalTransactions();
    } else if (tabValue === 2) {
      fetchUssdTransactions(ussdPage);
      fetchTotalUssd();
    } else if (tabValue === 3) {
      fetchKaditoTransactions(kaditoPage);
      fetchTotalKadito();
    } else if (tabValue === 4) {
      fetchBundles();
    }
  }, [
    tabValue,
    numbersPage,
    transactionsPage,
    ussdPage,
    kaditoPage,
    fetchNumbers,
    fetchTransactions,
    fetchUssdTransactions,
    fetchKaditoTransactions,
    fetchTotalNumbers,
    fetchTotalTransactions,
    fetchTotalUssd,
    fetchTotalKadito,
    fetchBundles,
  ]);

  /* ──────────────────────  DOWNLOAD HANDLERS  ────────────────────── */
  const handleDownloadNumbers = useCallback(async () => {
    try {
      setLoading(true);
      const q = query(
        collection(db, "entries"),
        where("exported", "==", false),
        limit(maxExportRecords),
      );
      const snap = await getDocs(q);
      const docs = snap.docs;
      const data = docs.map((d) => ({
        "Phone Number": formatPhoneNumber(d.data().phoneNumber),
        "Network Provider": d.data().networkProvider || "N/A",
      }));
      setRecordCount(docs.length);
      for (let i = 0; i < docs.length; i += batchSize) {
        const batch = writeBatch(db);
        docs.slice(i, i + batchSize).forEach((docSnap) => {
          batch.update(doc(db, "entries", docSnap.id), { exported: true });
        });
        await batch.commit();
      }
      downloadExcel(data, "Numbers.xlsx", ["Phone Number", "Network Provider"]);
      setNumbersCache({});
      setNumbersPage(1);
      setNumbersLastDocs([]);
      setHasMoreNumbers(true);
      await fetchNumbers(1);
      await fetchTotalNumbers();
    } catch (e) {
      setError(`Numbers download: ${e.message}`);
    } finally {
      setLoading(false);
    }
  }, [fetchNumbers, fetchTotalNumbers]);

  const handleDownloadTransactions = useCallback(async () => {
    try {
      setLoading(true);
      const today = getTodayStart();
      const q = query(
        collection(db, "approve_teller_transaction"),
        where("createdAt", ">=", today),
        where("status", "==", "approved"),
        where("exported", "==", false),
        limit(maxExportRecords),
      );
      const snap = await getDocs(q);
      const docs = snap.docs;
      const data = docs.map((d) => ({
        Number: formatPhoneNumber(
          d.data().subscriber_number || d.data().number,
        ),
        GB: d.data().gb || extractGB(d.data().desc) || "N/A",
      }));
      setRecordCount(docs.length);
      for (let i = 0; i < docs.length; i += batchSize) {
        const batch = writeBatch(db);
        docs.slice(i, i + batchSize).forEach((docSnap) => {
          batch.update(doc(db, "approve_teller_transaction", docSnap.id), {
            exported: true,
          });
        });
        await batch.commit();
      }
      downloadExcel(data, "Transactions.xlsx", ["Number", "GB"]);
      setTransactionsCache({});
      setTransactionsPage(1);
      setTransactionsLastDocs([]);
      setHasMoreTransactions(true);
      await fetchTransactions(1);
      await fetchTotalTransactions();
    } catch (e) {
      setError(`Transactions download: ${e.message}`);
    } finally {
      setLoading(false);
    }
  }, [fetchTransactions, fetchTotalTransactions]);

  const handleDownloadUssd = useCallback(async () => {
    try {
      setLoading(true);
      const today = getTodayStart();
      const q = query(
        collection(db, "teller_response"),
        where("createdAt", ">=", today),
        where("status", "==", "approved"),
        where("exported", "==", false),
        limit(maxExportRecords),
      );
      const snap = await getDocs(q);
      const docs = snap.docs;
      const data = docs.map((d) => ({
        Number: formatPhoneNumber(
          d.data().subscriber_number || d.data().number,
        ),
        GB: d.data().gb || extractGB(d.data().desc) || "N/A",
      }));
      setRecordCount(docs.length);
      for (let i = 0; i < docs.length; i += batchSize) {
        const batch = writeBatch(db);
        docs.slice(i, i + batchSize).forEach((docSnap) => {
          batch.update(doc(db, "teller_response", docSnap.id), {
            exported: true,
          });
        });
        await batch.commit();
      }
      downloadExcel(data, "UssdTransactions.xlsx", ["Number", "GB"]);
      setUssdCache({});
      setUssdPage(1);
      setUssdLastDocs([]);
      setHasMoreUssd(true);
      await fetchUssdTransactions(1);
      await fetchTotalUssd();
    } catch (e) {
      setError(`USSD download: ${e.message}`);
    } finally {
      setLoading(false);
    }
  }, [fetchUssdTransactions, fetchTotalUssd]);

  const handleDownloadKadito = useCallback(async () => {
    try {
      setLoading(true);
      const today = getTodayStart();
      const q = query(
        collection(db, "kadis_purchase"),
        where("createdAt", ">=", today),
        where("status", "==", "approved"),
        where("exported", "==", false),
        limit(maxExportRecords),
      );
      const snap = await getDocs(q);
      const docs = snap.docs;
      const data = docs.map((d) => ({
        Number: formatPhoneNumber(
          d.data().subscriber_number || d.data().number,
        ),
        GB: d.data().gb || extractGB(d.data().desc) || "N/A",
      }));
      setRecordCount(docs.length);
      for (let i = 0; i < docs.length; i += batchSize) {
        const batch = writeBatch(db);
        docs.slice(i, i + batchSize).forEach((docSnap) => {
          batch.update(doc(db, "kadis_purchase", docSnap.id), {
            exported: true,
          });
        });
        await batch.commit();
      }
      downloadExcel(data, "KaditoTransactions.xlsx", ["Number", "GB"]);
      setKaditoCache({});
      setKaditoPage(1);
      setKaditoLastDocs([]);
      setHasMoreKadito(true);
      await fetchKaditoTransactions(1);
      await fetchTotalKadito();
    } catch (e) {
      setError(`Kadito download: ${e.message}`);
    } finally {
      setLoading(false);
    }
  }, [fetchKaditoTransactions, fetchTotalKadito]);

  /* ──────────────────────  CONFIRM DIALOG LOGIC  ────────────────────── */
  const openConfirmDialog = useCallback(
    async (action) => {
      try {
        setLoading(true);
        let q;
        if (tabValue === 0) {
          q = query(
            collection(db, "entries"),
            where("exported", "==", false),
            limit(maxExportRecords),
          );
        } else if (tabValue === 1) {
          q = query(
            collection(db, "approve_teller_transaction"),
            where("createdAt", ">=", getTodayStart()),
            where("status", "==", "approved"),
            where("exported", "==", false),
            limit(maxExportRecords),
          );
        } else if (tabValue === 2) {
          q = query(
            collection(db, "teller_response"),
            where("createdAt", ">=", getTodayStart()),
            where("status", "==", "approved"),
            where("exported", "==", false),
            limit(maxExportRecords),
          );
        } else if (tabValue === 3) {
          q = query(
            collection(db, "kadis_purchase"),
            where("createdAt", ">=", getTodayStart()),
            where("status", "==", "approved"),
            where("exported", "==", false),
            limit(maxExportRecords),
          );
        }
        const snap = await getDocs(q);
        setRecordCount(snap.docs.length);
        setConfirmAction(() => action);
        setShowConfirmDialog(true);
      } catch (e) {
        setError(`Record count: ${e.message}`);
      } finally {
        setLoading(false);
      }
    },
    [tabValue],
  );

  const closeConfirmDialog = useCallback(() => {
    setShowConfirmDialog(false);
    setConfirmAction(null);
    setRecordCount(0);
  }, []);

  const confirmDownload = useCallback(() => {
    if (confirmAction) confirmAction();
    closeConfirmDialog();
  }, [confirmAction, closeConfirmDialog]);

  /* ──────────────────────  PAGINATION HANDLERS  ────────────────────── */
  const debouncedPrev = useMemo(
    () =>
      debounce(() => {
        if (tabValue === 0 && numbersPage > 1) setNumbersPage((p) => p - 1);
        else if (tabValue === 1 && transactionsPage > 1)
          setTransactionsPage((p) => p - 1);
        else if (tabValue === 2 && ussdPage > 1) setUssdPage((p) => p - 1);
        else if (tabValue === 3 && kaditoPage > 1) setKaditoPage((p) => p - 1);
      }, 300),
    [tabValue, numbersPage, transactionsPage, ussdPage, kaditoPage],
  );

  const debouncedNext = useMemo(
    () =>
      debounce(() => {
        if (tabValue === 0 && hasMoreNumbers) setNumbersPage((p) => p + 1);
        else if (tabValue === 1 && hasMoreTransactions)
          setTransactionsPage((p) => p + 1);
        else if (tabValue === 2 && hasMoreUssd) setUssdPage((p) => p + 1);
        else if (tabValue === 3 && hasMoreKadito) setKaditoPage((p) => p + 1);
      }, 300),
    [tabValue, hasMoreNumbers, hasMoreTransactions, hasMoreUssd, hasMoreKadito],
  );

  const handlePrevPage = useCallback(() => debouncedPrev(), [debouncedPrev]);
  const handleNextPage = useCallback(() => debouncedNext(), [debouncedNext]);

  /* ──────────────────────  CLEAR ERROR  ────────────────────── */
  useEffect(() => {
    if (error) {
      const t = setTimeout(() => setError(null), 5000);
      return () => clearTimeout(t);
    }
  }, [error]);

  /* ──────────────────────  RENDER  ────────────────────── */
  return (
    <div className="w-full max-w-7xl mx-auto p-4 sm:p-6 lg:p-8 min-h-screen bg-gray-100">
      {/* Confirmation Dialog */}
      {showConfirmDialog && (
        <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg p-6 w-full max-w-md mx-4">
            <h3 className="text-lg font-semibold text-gray-800 mb-4">
              Confirm Export
            </h3>
            <p className="text-gray-600 mb-6">
              Export {recordCount}{" "}
              {tabValue === 0
                ? "numbers"
                : tabValue === 1
                  ? "transactions"
                  : tabValue === 2
                    ? "ussd transactions"
                    : "kadito transactions"}
              ?
              {recordCount >= maxExportRecords &&
                ` (first ${maxExportRecords} only)`}
            </p>
            <div className="flex justify-end space-x-4">
              <button
                onClick={closeConfirmDialog}
                className="px-4 py-2 bg-gray-300 text-gray-800 rounded-lg hover:bg-gray-400"
              >
                Cancel
              </button>
              <button
                onClick={confirmDownload}
                className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700"
              >
                Confirm
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Tabs */}
      <div className="flex flex-wrap border-b border-gray-300 bg-white rounded-lg shadow-sm mb-6">
        <button
          className={`flex-1 px-4 py-3 text-sm font-semibold sm:text-base ${
            tabValue === 0
              ? "border-b-4 border-blue-600 text-blue-600 bg-blue-50"
              : "text-gray-600 hover:text-blue-600 hover:bg-gray-50"
          }`}
          onClick={() => handleTabChange(0)}
        >
          Numbers
        </button>
        <button
          className={`flex-1 px-4 py-3 text-sm font-semibold sm:text-base ${
            tabValue === 1
              ? "border-b-4 border-blue-600 text-blue-600 bg-blue-50"
              : "text-gray-600 hover:text-blue-600 hover:bg-gray-50"
          }`}
          onClick={() => handleTabChange(1)}
        >
          Website Transactions
        </button>
        <button
          className={`flex-1 px-4 py-3 text-sm font-semibold sm:text-base ${
            tabValue === 2
              ? "border-b-4 border-blue-600 text-blue-600 bg-blue-50"
              : "text-gray-600 hover:text-blue-600 hover:bg-gray-50"
          }`}
          onClick={() => handleTabChange(2)}
        >
          USSD Transactions
        </button>
        <button
          className={`flex-1 px-4 py-3 text-sm font-semibold sm:text-base ${
            tabValue === 3
              ? "border-b-4 border-blue-600 text-blue-600 bg-blue-50"
              : "text-gray-600 hover:text-blue-600 hover:bg-gray-50"
          }`}
          onClick={() => handleTabChange(3)}
        >
          Kadito Transaction
        </button>
        <button
          className={`flex-1 px-4 py-3 text-sm font-semibold sm:text-base ${
            tabValue === 4
              ? "border-b-4 border-blue-600 text-blue-600 bg-blue-50"
              : "text-gray-600 hover:text-blue-600 hover:bg-gray-50"
          }`}
          onClick={() => handleTabChange(4)}
        >
          Bundles
        </button>
      </div>

      {/* Loading / Error */}
      {loading && (
        <div className="mt-6 flex justify-center">
          <div className="animate-spin rounded-full h-8 w-8 border-t-2 border-b-2 border-blue-600"></div>
        </div>
      )}
      {!loading && error && (
        <p className="mt-6 text-center text-red-500 text-lg">{error}</p>
      )}

      {/* ─── Numbers Tab Content ─── */}
      {tabValue === 0 && !loading && !error && (
        <div className="mt-6">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-4">
            <h2 className="text-xl sm:text-2xl font-bold text-gray-800">
              New Numbers
            </h2>
            {numbers.length > 0 && (
              <button
                onClick={() => openConfirmDialog(handleDownloadNumbers)}
                className="mt-2 sm:mt-0 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 text-sm sm:text-base shadow-md"
              >
                Download Numbers (Excel)
              </button>
            )}
          </div>
          <p className="text-sm text-gray-600 mb-4">
            Total Records: {totalNumbers} | Current Page: {numbers.length} (Page{" "}
            {numbersPage})
          </p>
          {numbers.length > 0 ? (
            <>
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                {numbers.map((num) => (
                  <div
                    key={num.id}
                    className="p-4 bg-white rounded-lg shadow-md hover:shadow-lg"
                  >
                    <p className="text-sm sm:text-base">
                      <span className="font-semibold text-gray-700">
                        Phone Number:
                      </span>{" "}
                      {formatPhoneNumber(num.phoneNumber)}
                    </p>
                    <p className="text-sm sm:text-base">
                      <span className="font-semibold text-gray-700">
                        Network Provider:
                      </span>{" "}
                      {num.networkProvider || "N/A"}
                    </p>
                  </div>
                ))}
              </div>
              <div className="flex justify-between items-center mt-6">
                <button
                  onClick={handlePrevPage}
                  disabled={numbersPage === 1}
                  className={`px-4 py-2 rounded-lg text-sm sm:text-base ${
                    numbersPage === 1
                      ? "bg-gray-300 text-gray-500 cursor-not-allowed"
                      : "bg-blue-600 text-white hover:bg-blue-700"
                  }`}
                >
                  Previous
                </button>
                <span className="text-sm sm:text-base text-gray-600">
                  Page {numbersPage}
                </span>
                <button
                  onClick={handleNextPage}
                  disabled={!hasMoreNumbers}
                  className={`px-4 py-2 rounded-lg text-sm sm:text-base ${
                    !hasMoreNumbers
                      ? "bg-gray-300 text-gray-500 cursor-not-allowed"
                      : "bg-blue-600 text-white hover:bg-blue-700"
                  }`}
                >
                  Next
                </button>
              </div>
            </>
          ) : (
            <p className="text-gray-600 text-center text-lg">
              No numbers found.
            </p>
          )}
        </div>
      )}

      {/* ─── Website Transactions Tab Content ─── */}
      {tabValue === 1 && !loading && !error && (
        <div className="mt-6">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-4">
            <h2 className="text-xl sm:text-2xl font-bold text-gray-800">
              Today's Transactions
            </h2>
            {transactions.length > 0 && (
              <button
                onClick={() => openConfirmDialog(handleDownloadTransactions)}
                className="mt-2 sm:mt-0 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 text-sm sm:text-base shadow-md"
              >
                Download Transactions (Excel)
              </button>
            )}
          </div>
          <p className="text-sm text-gray-600 mb-4">
            Total Records: {totalTransactions} | Current Page:{" "}
            {transactions.length} (Page {transactionsPage})
          </p>
          {transactions.length > 0 ? (
            <>
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                {transactions.map((tx) => (
                  <div
                    key={tx.id}
                    className="p-4 bg-white rounded-lg shadow-md hover:shadow-lg"
                  >
                    <p className="text-sm sm:text-base">
                      <span className="font-semibold text-gray-700">
                        Number:
                      </span>{" "}
                      {formatPhoneNumber(tx.subscriber_number || tx.number)}
                    </p>
                    <p className="text-sm sm:text-base">
                      <span className="font-semibold text-gray-700">GB:</span>{" "}
                      {tx.gb || extractGB(tx.desc) || "N/A"}
                    </p>
                  </div>
                ))}
              </div>
              <div className="flex justify-between items-center mt-6">
                <button
                  onClick={handlePrevPage}
                  disabled={transactionsPage === 1}
                  className={`px-4 py-2 rounded-lg text-sm sm:text-base ${
                    transactionsPage === 1
                      ? "bg-gray-300 text-gray-500 cursor-not-allowed"
                      : "bg-blue-600 text-white hover:bg-blue-700"
                  }`}
                >
                  Previous
                </button>
                <span className="text-sm sm:text-base text-gray-600">
                  Page {transactionsPage}
                </span>
                <button
                  onClick={handleNextPage}
                  disabled={!hasMoreTransactions}
                  className={`px-4 py-2 rounded-lg text-sm sm:text-base ${
                    !hasMoreTransactions
                      ? "bg-gray-300 text-gray-500 cursor-not-allowed"
                      : "bg-blue-600 text-white hover:bg-blue-700"
                  }`}
                >
                  Next
                </button>
              </div>
            </>
          ) : (
            <p className="text-gray-600 text-center text-lg">
              No transactions found.
            </p>
          )}
        </div>
      )}

      {/* ─── USSD Transactions Tab Content ─── */}
      {tabValue === 2 && !loading && !error && (
        <div className="mt-6">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-4">
            <h2 className="text-xl sm:text-2xl font-bold text-gray-800">
              Today's USSD Transactions
            </h2>
            {ussdTransactions.length > 0 && (
              <button
                onClick={() => openConfirmDialog(handleDownloadUssd)}
                className="mt-2 sm:mt-0 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 text-sm sm:text-base shadow-md"
              >
                Download USSD Transactions (Excel)
              </button>
            )}
          </div>
          <p className="text-sm text-gray-600 mb-4">
            Total Records: {totalUssd} | Current Page: {ussdTransactions.length}{" "}
            (Page {ussdPage})
          </p>
          {ussdTransactions.length > 0 ? (
            <>
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                {ussdTransactions.map((tx) => (
                  <div
                    key={tx.id}
                    className="p-4 bg-white rounded-lg shadow-md hover:shadow-lg"
                  >
                    <p className="text-sm sm:text-base">
                      <span className="font-semibold text-gray-700">
                        Number:
                      </span>{" "}
                      {formatPhoneNumber(tx.subscriber_number || tx.number)}
                    </p>
                    <p className="text-sm sm:text-base">
                      <span className="font-semibold text-gray-700">GB:</span>{" "}
                      {tx.gb || extractGB(tx.desc) || "N/A"}
                    </p>
                  </div>
                ))}
              </div>
              <div className="flex justify-between items-center mt-6">
                <button
                  onClick={handlePrevPage}
                  disabled={ussdPage === 1}
                  className={`px-4 py-2 rounded-lg text-sm sm:text-base ${
                    ussdPage === 1
                      ? "bg-gray-300 text-gray-500 cursor-not-allowed"
                      : "bg-blue-600 text-white hover:bg-blue-700"
                  }`}
                >
                  Previous
                </button>
                <span className="text-sm sm:text-base text-gray-600">
                  Page {ussdPage}
                </span>
                <button
                  onClick={handleNextPage}
                  disabled={!hasMoreUssd}
                  className={`px-4 py-2 rounded-lg text-sm sm:text-base ${
                    !hasMoreUssd
                      ? "bg-gray-300 text-gray-500 cursor-not-allowed"
                      : "bg-blue-600 text-white hover:bg-blue-700"
                  }`}
                >
                  Next
                </button>
              </div>
            </>
          ) : (
            <p className="text-gray-600 text-center text-lg">
              No USSD transactions found.
            </p>
          )}
        </div>
      )}

      {/* ─── Kadito Transactions Tab Content ─── */}
      {tabValue === 3 && !loading && !error && (
        <div className="mt-6">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-4">
            <h2 className="text-xl sm:text-2xl font-bold text-gray-800">
              Today's Kadito Transactions
            </h2>
            {kaditoTransactions.length > 0 && (
              <button
                onClick={() => openConfirmDialog(handleDownloadKadito)}
                className="mt-2 sm:mt-0 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 text-sm sm:text-base shadow-md"
              >
                Download Kadito (Excel)
              </button>
            )}
          </div>

          <p className="text-sm text-gray-600 mb-4">
            Total Records: {totalKadito} | Current Page:{" "}
            {kaditoTransactions.length} (Page {kaditoPage})
          </p>

          {kaditoTransactions.length > 0 ? (
            <>
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                {kaditoTransactions.map((tx) => (
                  <div
                    key={tx.id}
                    className="p-4 bg-white rounded-lg shadow-md hover:shadow-lg transition-shadow"
                  >
                    <p className="text-sm sm:text-base">
                      <span className="font-semibold text-gray-700">
                        Number:
                      </span>{" "}
                      {formatPhoneNumber(tx.subscriber_number || tx.number)}
                    </p>
                    <p className="text-sm sm:text-base">
                      <span className="font-semibold text-gray-700">GB:</span>{" "}
                      {tx.gb || extractGB(tx.desc) || "N/A"}
                    </p>
                  </div>
                ))}
              </div>

              <div className="flex justify-between items-center mt-6">
                <button
                  onClick={handlePrevPage}
                  disabled={kaditoPage === 1}
                  className={`px-4 py-2 rounded-lg text-sm sm:text-base ${
                    kaditoPage === 1
                      ? "bg-gray-300 text-gray-500 cursor-not-allowed"
                      : "bg-blue-600 text-white hover:bg-blue-700"
                  }`}
                >
                  Previous
                </button>
                <span className="text-sm sm:text-base text-gray-600">
                  Page {kaditoPage}
                </span>
                <button
                  onClick={handleNextPage}
                  disabled={!hasMoreKadito}
                  className={`px-4 py-2 rounded-lg text-sm sm:text-base ${
                    !hasMoreKadito
                      ? "bg-gray-300 text-gray-500 cursor-not-allowed"
                      : "bg-blue-600 text-white hover:bg-blue-700"
                  }`}
                >
                  Next
                </button>
              </div>
            </>
          ) : (
            <p className="text-gray-600 text-center text-lg">
              No Kadito transactions found.
            </p>
          )}
        </div>
      )}

      {/* ─── Bundles Management Tab ─── */}
      {tabValue === 4 && (
        <div className="mt-6">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6">
            <h2 className="text-2xl font-bold text-gray-800">
              Manage Data Bundles
            </h2>
            <div className="mt-3 sm:mt-0 flex items-center space-x-4">
              {(Object.keys(priceChanges).length > 0 ||
                Object.keys(activeChanges).length > 0) && (
                <button
                  onClick={saveBundleChanges}
                  disabled={bundlesLoading}
                  className="px-5 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 disabled:opacity-50 shadow-md"
                >
                  Save Changes
                </button>
              )}
              <button
                onClick={fetchBundles}
                disabled={bundlesLoading}
                className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 disabled:opacity-50"
              >
                Refresh
              </button>
            </div>
          </div>

          {bundlesLoading && (
            <div className="flex justify-center my-12">
              <div className="animate-spin rounded-full h-10 w-10 border-t-2 border-b-2 border-blue-600"></div>
            </div>
          )}

          {bundlesError && (
            <p className="text-red-600 text-center text-lg">{bundlesError}</p>
          )}

          {!bundlesLoading && !bundlesError && (
            <>
              {/* Network tabs */}
              <div className="flex border-b mb-6 overflow-x-auto">
                {["mtn", "tigo", "telecel"].map((net, idx) => (
                  <button
                    key={net}
                    className={`flex-1 py-3 px-4 font-medium uppercase whitespace-nowrap ${
                      bundlesTabValue === idx
                        ? "border-b-4 border-blue-600 text-blue-700"
                        : "text-gray-500 hover:text-gray-700"
                    }`}
                    onClick={() => setBundlesTabValue(idx)}
                  >
                    {net.toUpperCase()}
                  </button>
                ))}
              </div>

              {["mtn", "tigo", "telecel"].map((network, idx) => (
                <div
                  key={network}
                  className={bundlesTabValue === idx ? "block" : "hidden"}
                >
                  {["daily", "weekly", "monthly"].map((period) => {
                    const plans = bundlesData[network]?.[period] || [];
                    if (plans.length === 0) return null;

                    return (
                      <div key={period} className="mb-10">
                        <h3 className="text-xl font-semibold capitalize mb-4 text-gray-800">
                          {period} Plans
                        </h3>
                        <div className="overflow-x-auto">
                          <table className="min-w-full bg-white border border-gray-200">
                            <thead className="bg-gray-100">
                              <tr>
                                <th className="px-4 py-3 text-left text-sm font-medium text-gray-700">
                                  ID
                                </th>
                                <th className="px-4 py-3 text-left text-sm font-medium text-gray-700">
                                  Name
                                </th>
                                <th className="px-4 py-3 text-left text-sm font-medium text-gray-700">
                                  Size
                                </th>
                                <th className="px-4 py-3 text-left text-sm font-medium text-gray-700">
                                  Price (GHS)
                                </th>
                                <th className="px-4 py-3 text-left text-sm font-medium text-gray-700">
                                  Active
                                </th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-gray-200">
                              {plans.map((plan) => {
                                const key = `${network}/${period}/${plan.id}`;
                                const displayPrice =
                                  priceChanges[key] !== undefined
                                    ? priceChanges[key]
                                    : plan.price;
                                const displayActive =
                                  activeChanges[key] !== undefined
                                    ? activeChanges[key]
                                    : (plan.active ?? true);

                                return (
                                  <tr
                                    key={plan.id}
                                    className="hover:bg-gray-50"
                                  >
                                    <td className="px-4 py-3 text-sm font-medium text-gray-900">
                                      {plan.id}
                                    </td>
                                    <td className="px-4 py-3 text-sm text-gray-900">
                                      {plan.name}
                                    </td>
                                    <td className="px-4 py-3 text-sm text-gray-900">
                                      {plan.size}
                                    </td>
                                    <td className="px-4 py-3">
                                      <input
                                        type="number"
                                        step="0.5"
                                        min="0"
                                        value={displayPrice}
                                        onChange={(e) =>
                                          handlePriceChange(
                                            network,
                                            period,
                                            plan.id,
                                            e.target.value,
                                          )
                                        }
                                        className="w-24 px-2 py-1 border border-gray-300 rounded focus:ring-blue-500 focus:border-blue-500"
                                      />
                                    </td>
                                    <td className="px-4 py-3">
                                      <input
                                        type="checkbox"
                                        checked={displayActive}
                                        onChange={() =>
                                          handleActiveToggle(
                                            network,
                                            period,
                                            plan.id,
                                            plan.active ?? true,
                                          )
                                        }
                                        className="h-5 w-5 text-blue-600 rounded border-gray-300 focus:ring-blue-500"
                                      />
                                    </td>
                                  </tr>
                                );
                              })}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    );
                  })}
                </div>
              ))}
            </>
          )}
        </div>
      )}
    </div>
  );
};

export default Dashboard;


