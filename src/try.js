import React, { useState, useEffect, useCallback } from "react";
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
} from "firebase/firestore";
import { db } from "./firebase"; // Import Firestore instance from your firebase config
import * as XLSX from "xlsx";

// Utility to get start of today for date filtering
const getTodayStart = () => {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return today;
};

// Utility to extract GB from desc field for Type 1 documents
const extractGB = (desc) => {
  if (!desc) return "N/A";
  const match = desc.match(/(\d+)GB/);
  return match ? match[1] : "N/A";
};

// Utility to format phone number (e.g., "233549856098" to "0549856098")
const formatPhoneNumber = (number) => {
  if (!number) return "N/A";
  const cleanedNumber = number.replace(/^233/, "");
  return cleanedNumber.length === 9
    ? `0${cleanedNumber}`
    : cleanedNumber || "N/A";
};

// Utility to download data as Excel
const downloadExcel = (data, fileName, headers) => {
  const worksheet = XLSX.utils.json_to_sheet(data, { header: headers });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  XLSX.writeFile(workbook, fileName);
};

// Utility to debounce a function
const debounce = (func, wait) => {
  let timeout;
  return (...args) => {
    clearTimeout(timeout);
    timeout = setTimeout(() => func(...args), wait);
  };
};

const Dashboard = () => {
  const [tabValue, setTabValue] = useState(0);
  const [numbers, setNumbers] = useState([]);
  const [transactions, setTransactions] = useState([]);
  const [ussdTransactions, setUssdTransactions] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [numbersPage, setNumbersPage] = useState(1);
  const [transactionsPage, setTransactionsPage] = useState(1);
  const [ussdPage, setUssdPage] = useState(1);
  const [numbersLastDocs, setNumbersLastDocs] = useState([]);
  const [transactionsLastDocs, setTransactionsLastDocs] = useState([]);
  const [ussdLastDocs, setUssdLastDocs] = useState([]);
  const [hasMoreNumbers, setHasMoreNumbers] = useState(true);
  const [hasMoreTransactions, setHasMoreTransactions] = useState(true);
  const [hasMoreUssd, setHasMoreUssd] = useState(true);
  const [showConfirmDialog, setShowConfirmDialog] = useState(false);
  const [confirmAction, setConfirmAction] = useState(null);
  const [recordCount, setRecordCount] = useState(0);
  const [numbersCache, setNumbersCache] = useState({});
  const [transactionsCache, setTransactionsCache] = useState({});
  const [ussdCache, setUssdCache] = useState({});
  const [totalNumbers, setTotalNumbers] = useState(0);
  const [totalTransactions, setTotalTransactions] = useState(0);
  const [totalUssd, setTotalUssd] = useState(0);
  const pageSize = 6; // Number of records per page
  const maxExportRecords = 1000; // Limit for export to prevent overload
  const batchSize = 500; // Firestore batch write limit

  // Handle tab change with pagination reset
  const handleTabChange = (newValue) => {
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
    }
    setError(null); // Clear error on tab change
  };

  // Fetch total count of numbers
  const fetchTotalNumbers = async () => {
    try {
      const q = query(
        collection(db, "entries"),
        where("exported", "==", false)
      );
      const querySnapshot = await getDocs(q);
      setTotalNumbers(querySnapshot.size);
    } catch (err) {
      setError("Failed to fetch total numbers count: " + err.message);
    }
  };

  // Fetch total count of transactions
  const fetchTotalTransactions = async () => {
    try {
      const today = getTodayStart();
      const q = query(
        collection(db, "data_approve_teller_transaction"),
        where("createdAt", ">=", today),
        where("status", "==", "approved"),
        where("exported", "==", false)
      );
      const querySnapshot = await getDocs(q);
      setTotalTransactions(querySnapshot.size);
    } catch (err) {
      setError("Failed to fetch total transactions count: " + err.message);
      console.log(err.message);
    }
  };

  // Fetch total count of USSD transactions
  const fetchTotalUssd = async () => {
    try {
      const today = getTodayStart();
      const q = query(
        collection(db, "teller-response-calls"),
        where("createdAt", ">=", today),
        where("status", "==", "approved"),
        where("exported", "==", false)
      );
      const querySnapshot = await getDocs(q);
      setTotalUssd(querySnapshot.size);
    } catch (err) {
      setError("Failed to fetch total USSD transactions count: " + err.message);
      console.log(err.message);
    }
  };

  // Fetch numbers with pagination and caching
  const fetchNumbers = async (page = 1) => {
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
        limit(pageSize)
      );

      if (page > 1 && numbersLastDocs[page - 2]) {
        q = query(
          collection(db, "entries"),
          where("exported", "==", false),
          orderBy("phoneNumber"),
          startAfter(numbersLastDocs[page - 2]),
          limit(pageSize)
        );
      }

      const querySnapshot = await getDocs(q);
      const data = querySnapshot.docs.map((doc) => ({
        id: doc.id,
        ...doc.data(),
      }));

      setNumbers(data);
      setNumbersCache((prev) => ({ ...prev, [page]: data }));
      if (querySnapshot.docs.length > 0) {
        const newLastDocs = [...numbersLastDocs];
        newLastDocs[page - 1] =
          querySnapshot.docs[querySnapshot.docs.length - 1];
        setNumbersLastDocs(newLastDocs);
      }
      setHasMoreNumbers(querySnapshot.docs.length === pageSize);
      setLoading(false);
    } catch (err) {
      setError("Failed to fetch numbers: " + err.message);
      setLoading(false);
    }
  };

  // Fetch transactions with pagination and caching
  const fetchTransactions = async (page = 1) => {
    if (transactionsCache[page]) {
      setTransactions(transactionsCache[page]);
      setLoading(false);
      return;
    }

    try {
      setLoading(true);
      const today = getTodayStart();
      let q = query(
        collection(db, "data_approve_teller_transaction"),
        where("createdAt", ">=", today),
        where("status", "==", "approved"),
        where("exported", "==", false),
        orderBy("createdAt"),
        limit(pageSize)
      );

      if (page > 1 && transactionsLastDocs[page - 2]) {
        q = query(
          collection(db, "data_approve_teller_transaction"),
          where("createdAt", ">=", today),
          where("status", "==", "approved"),
          where("exported", "==", false),
          orderBy("createdAt"),
          startAfter(transactionsLastDocs[page - 2]),
          limit(pageSize)
        );
      }

      const querySnapshot = await getDocs(q);
      const data = querySnapshot.docs.map((doc) => ({
        id: doc.id,
        ...doc.data(),
      }));

      setTransactions(data);
      setTransactionsCache((prev) => ({ ...prev, [page]: data }));
      if (querySnapshot.docs.length > 0) {
        const newLastDocs = [...transactionsLastDocs];
        newLastDocs[page - 1] =
          querySnapshot.docs[querySnapshot.docs.length - 1];
        setTransactionsLastDocs(newLastDocs);
      }
      setHasMoreTransactions(querySnapshot.docs.length === pageSize);
      setLoading(false);
    } catch (err) {
      setError("Failed to fetch transactions: " + err.message);
      setLoading(false);
    }
  };

  // Fetch USSD transactions with pagination and caching
  const fetchUssdTransactions = async (page = 1) => {
    if (ussdCache[page]) {
      setUssdTransactions(ussdCache[page]);
      setLoading(false);
      return;
    }

    try {
      setLoading(true);
      const today = getTodayStart();
      let q = query(
        collection(db, "teller-response-calls"),
        where("createdAt", ">=", today),
        where("status", "==", "approved"),
        where("exported", "==", false),
        orderBy("createdAt"),
        limit(pageSize)
      );

      if (page > 1 && ussdLastDocs[page - 2]) {
        q = query(
          collection(db, "teller-response-calls"),
          where("createdAt", ">=", today),
          where("status", "==", "approved"),
          where("exported", "==", false),
          orderBy("createdAt"),
          startAfter(ussdLastDocs[page - 2]),
          limit(pageSize)
        );
      }

      const querySnapshot = await getDocs(q);
      const data = querySnapshot.docs.map((doc) => ({
        id: doc.id,
        ...doc.data(),
      }));

      setUssdTransactions(data);
      setUssdCache((prev) => ({ ...prev, [page]: data }));
      if (querySnapshot.docs.length > 0) {
        const newLastDocs = [...ussdLastDocs];
        newLastDocs[page - 1] =
          querySnapshot.docs[querySnapshot.docs.length - 1];
        setUssdLastDocs(newLastDocs);
      }
      setHasMoreUssd(querySnapshot.docs.length === pageSize);
      setLoading(false);
    } catch (err) {
      setError("Failed to fetch USSD transactions: " + err.message);
      setLoading(false);
    }
  };

  // Fetch data and total counts on component mount and page/tab change
  useEffect(() => {
    if (tabValue === 0) {
      fetchNumbers(numbersPage);
      fetchTotalNumbers(); // Fetch total numbers count
    } else if (tabValue === 1) {
      fetchTransactions(transactionsPage);
      fetchTotalTransactions(); // Fetch total transactions count
    } else if (tabValue === 2) {
      fetchUssdTransactions(ussdPage);
      fetchTotalUssd(); // Fetch total USSD transactions count
    }
  }, [tabValue, numbersPage, transactionsPage, ussdPage]);

  // Handle download for Numbers with batch update
  const handleDownloadNumbers = async () => {
    try {
      setLoading(true);
      const q = query(
        collection(db, "entries"),
        where("exported", "==", false),
        limit(maxExportRecords)
      );
      const querySnapshot = await getDocs(q);
      const docs = querySnapshot.docs;
      const data = docs.map((doc) => ({
        "Phone Number": formatPhoneNumber(doc.data().phoneNumber),
        "Network Provider": doc.data().networkProvider || "N/A",
      }));

      setRecordCount(docs.length);

      // Split batch writes into chunks of 500
      for (let i = 0; i < docs.length; i += batchSize) {
        const batch = writeBatch(db);
        const batchDocs = docs.slice(i, i + batchSize);
        batchDocs.forEach((docSnap) => {
          const docRef = doc(db, "entries", docSnap.id);
          batch.update(docRef, { exported: true });
        });
        await batch.commit();
      }

      // Generate and download Excel
      downloadExcel(data, "Numbers.xlsx", ["Phone Number", "Network Provider"]);

      // Clear cache and refresh data
      setNumbersCache({});
      setNumbersPage(1);
      setNumbersLastDocs([]);
      setHasMoreNumbers(true);
      await fetchNumbers(1);
      await fetchTotalNumbers(); // Refresh total numbers count
    } catch (err) {
      setError("Failed to download or update numbers: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  // Handle download for Transactions with batch update
  const handleDownloadTransactions = async () => {
    try {
      setLoading(true);
      const today = getTodayStart();
      const q = query(
        collection(db, "data_approve_teller_transaction"),
        where("createdAt", ">=", today),
        where("status", "==", "approved"),
        where("exported", "==", false),
        limit(maxExportRecords)
      );
      const querySnapshot = await getDocs(q);
      const docs = querySnapshot.docs;
      const data = docs.map((doc) => ({
        Number: formatPhoneNumber(
          doc.data().subscriber_number || doc.data().number
        ),
        GB: doc.data().gb || extractGB(doc.data().desc) || "N/A",
      }));

      setRecordCount(docs.length);

      // Split batch writes into chunks of 500
      for (let i = 0; i < docs.length; i += batchSize) {
        const batch = writeBatch(db);
        const batchDocs = docs.slice(i, i + batchSize);
        batchDocs.forEach((docSnap) => {
          const docRef = doc(db, "data_approve_teller_transaction", docSnap.id);
          batch.update(docRef, { exported: true });
        });
        await batch.commit();
      }

      // Generate and download Excel
      downloadExcel(data, "Transactions.xlsx", ["Number", "GB"]);

      // Clear cache and refresh data
      setTransactionsCache({});
      setTransactionsPage(1);
      setTransactionsLastDocs([]);
      setHasMoreTransactions(true);
      await fetchTransactions(1);
      await fetchTotalTransactions(); // Refresh total transactions count
    } catch (err) {
      setError("Failed to download or update transactions: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  // Handle download for USSD Transactions with batch update
  const handleDownloadUssd = async () => {
    try {
      setLoading(true);
      const today = getTodayStart();
      const q = query(
        collection(db, "teller-response-calls"),
        where("createdAt", ">=", today),
        where("status", "==", "approved"),
        where("exported", "==", false),
        limit(maxExportRecords)
      );
      const querySnapshot = await getDocs(q);
      const docs = querySnapshot.docs;
      const data = docs.map((doc) => ({
        Number: formatPhoneNumber(
          doc.data().subscriber_number || doc.data().number
        ),
        GB: doc.data().gb || extractGB(doc.data().desc) || "N/A",
      }));

      setRecordCount(docs.length);

      // Split batch writes into chunks of 500
      for (let i = 0; i < docs.length; i += batchSize) {
        const batch = writeBatch(db);
        const batchDocs = docs.slice(i, i + batchSize);
        batchDocs.forEach((docSnap) => {
          const docRef = doc(db, "teller-response-calls", docSnap.id);
          batch.update(docRef, { exported: true });
        });
        await batch.commit();
      }

      // Generate and download Excel
      downloadExcel(data, "UssdTransactions.xlsx", ["Number", "GB"]);

      // Clear cache and refresh data
      setUssdCache({});
      setUssdPage(1);
      setUssdLastDocs([]);
      setHasMoreUssd(true);
      await fetchUssdTransactions(1);
      await fetchTotalUssd(); // Refresh total USSD transactions count
    } catch (err) {
      setError(
        "Failed to download or update USSD transactions: " + err.message
      );
    } finally {
      setLoading(false);
    }
  };

  // Handle confirmation dialog
  const openConfirmDialog = async (action) => {
    try {
      setLoading(true);
      let q;
      if (tabValue === 0) {
        q = query(
          collection(db, "entries"),
          where("exported", "==", false),
          limit(maxExportRecords)
        );
      } else if (tabValue === 1) {
        q = query(
          collection(db, "data_approve_teller_transaction"),
          where("createdAt", ">=", getTodayStart()),
          where("status", "==", "approved"),
          where("exported", "==", false),
          limit(maxExportRecords)
        );
      } else if (tabValue === 2) {
        q = query(
          collection(db, "teller-response-calls"),
          where("createdAt", ">=", getTodayStart()),
          where("status", "==", "approved"),
          where("exported", "==", false),
          limit(maxExportRecords)
        );
      }
      const querySnapshot = await getDocs(q);
      setRecordCount(querySnapshot.docs.length);
      setConfirmAction(() => action);
      setShowConfirmDialog(true);
    } catch (err) {
      setError("Failed to fetch record count: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  const closeConfirmDialog = () => {
    setShowConfirmDialog(false);
    setConfirmAction(null);
    setRecordCount(0);
  };

  const confirmDownload = () => {
    if (confirmAction) {
      confirmAction();
    }
    closeConfirmDialog();
  };

  // Debounced pagination controls
  const handlePrevPage = useCallback(
    debounce(() => {
      if (tabValue === 0 && numbersPage > 1) {
        setNumbersPage((prev) => prev - 1);
      } else if (tabValue === 1 && transactionsPage > 1) {
        setTransactionsPage((prev) => prev - 1);
      } else if (tabValue === 2 && ussdPage > 1) {
        setUssdPage((prev) => prev - 1);
      }
    }, 300),
    [tabValue, numbersPage, transactionsPage, ussdPage]
  );

  const handleNextPage = useCallback(
    debounce(() => {
      if (tabValue === 0 && hasMoreNumbers) {
        setNumbersPage((prev) => prev + 1);
      } else if (tabValue === 1 && hasMoreTransactions) {
        setTransactionsPage((prev) => prev + 1);
      } else if (tabValue === 2 && hasMoreUssd) {
        setUssdPage((prev) => prev + 1);
      }
    }, 300),
    [tabValue, hasMoreNumbers, hasMoreTransactions, hasMoreUssd]
  );

  // Clear error after 5 seconds
  useEffect(() => {
    if (error) {
      const timer = setTimeout(() => setError(null), 5000);
      return () => clearTimeout(timer);
    }
  }, [error]);

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
              Are you sure you want to export {recordCount}{" "}
              {tabValue === 0
                ? "numbers"
                : tabValue === 1
                ? "transactions"
                : "ussd transactions"}
              ? This will mark them as exported.
              {recordCount >= maxExportRecords &&
                ` Only the first ${maxExportRecords} records will be exported due to system limits.`}
            </p>
            <div className="flex justify-end space-x-4">
              <button
                onClick={closeConfirmDialog}
                className="px-4 py-2 bg-gray-300 text-gray-800 rounded-lg hover:bg-gray-400 transition-colors duration-200"
              >
                Cancel
              </button>
              <button
                onClick={confirmDownload}
                className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors duration-200"
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
          className={`flex-1 px-4 py-3 text-sm font-semibold transition-colors duration-200 sm:text-base ${
            tabValue === 0
              ? "border-b-4 border-blue-600 text-blue-600 bg-blue-50"
              : "text-gray-600 hover:text-blue-600 hover:bg-gray-50"
          }`}
          onClick={() => handleTabChange(0)}
        >
          Numbers
        </button>
        <button
          className={`flex-1 px-4 py-3 text-sm font-semibold transition-colors duration-200 sm:text-base ${
            tabValue === 1
              ? "border-b-4 border-blue-600 text-blue-600 bg-blue-50"
              : "text-gray-600 hover:text-blue-600 hover:bg-gray-50"
          }`}
          onClick={() => handleTabChange(1)}
        >
          Website Transactions
        </button>
        <button
          className={`flex-1 px-4 py-3 text-sm font-semibold transition-colors duration-200 sm:text-base ${
            tabValue === 2
              ? "border-b-4 border-blue-600 text-blue-600 bg-blue-50"
              : "text-gray-600 hover:text-blue-600 hover:bg-gray-50"
          }`}
          onClick={() => handleTabChange(2)}
        >
          USSD Transactions
        </button>
      </div>

      {/* Loading Spinner */}
      {loading && (
        <div className="mt-6 flex justify-center">
          <div className="animate-spin rounded-full h-8 w-8 border-t-2 border-b-2 border-blue-600"></div>
        </div>
      )}

      {/* Error State */}
      {!loading && error && (
        <p className="mt-6 text-center text-red-500 text-lg">{error}</p>
      )}

      {/* Numbers Tab (entries collection) */}
      {tabValue === 0 && !loading && !error && (
        <div className="mt-6">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-4">
            <h2 className="text-xl sm:text-2xl font-bold text-gray-800">
              New Numbers
            </h2>
            {numbers.length > 0 && (
              <button
                onClick={() => openConfirmDialog(handleDownloadNumbers)}
                className="mt-2 sm:mt-0 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors duration-200 text-sm sm:text-base shadow-md"
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
                    className="p-4 bg-white rounded-lg shadow-md hover:shadow-lg transition-shadow duration-200"
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
                  } transition-colors duration-200`}
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
                  } transition-colors duration-200`}
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

      {/* Transactions Tab (data_approve_teller_transaction collection) */}
      {tabValue === 1 && !loading && !error && (
        <div className="mt-6">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-4">
            <h2 className="text-xl sm:text-2xl font-bold text-gray-800">
              Today's Transactions
            </h2>
            {transactions.length > 0 && (
              <button
                onClick={() => openConfirmDialog(handleDownloadTransactions)}
                className="mt-2 sm:mt-0 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors duration-200 text-sm sm:text-base shadow-md"
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
                    className="p-4 bg-white rounded-lg shadow-md hover:shadow-lg transition-shadow duration-200"
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
                  } transition-colors duration-200`}
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
                  } transition-colors duration-200`}
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

      {/* USSD Transactions Tab (teller-response-calls collection) */}
      {tabValue === 2 && !loading && !error && (
        <div className="mt-6">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-4">
            <h2 className="text-xl sm:text-2xl font-bold text-gray-800">
              Today's USSD Transactions
            </h2>
            {ussdTransactions.length > 0 && (
              <button
                onClick={() => openConfirmDialog(handleDownloadUssd)}
                className="mt-2 sm:mt-0 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors duration-200 text-sm sm:text-base shadow-md"
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
                    className="p-4 bg-white rounded-lg shadow-md hover:shadow-lg transition-shadow duration-200"
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
                  } transition-colors duration-200`}
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
                  } transition-colors duration-200`}
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
    </div>
  );
};

export default Dashboard;
