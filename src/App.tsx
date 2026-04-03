/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo, useRef, Component, ErrorInfo, ReactNode } from 'react';
import { 
  Search, RefreshCw, AlertCircle, LayoutGrid, 
  List as ListIcon, Edit2, Plus, X, Save, Trash2, LogIn, LogOut, User,
  Upload, Settings, FileSpreadsheet, Filter, ChevronDown, MoreVertical,
  Database, Table as TableIcon, ArrowLeft, ExternalLink
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  collection, onSnapshot, query, addDoc, updateDoc, deleteDoc, 
  doc, getDoc, setDoc, serverTimestamp, writeBatch, orderBy, where,
  getDocFromServer, Timestamp, limit, startAfter, getDocs, getCountFromServer
} from 'firebase/firestore';
import { signInWithPopup, GoogleAuthProvider, onAuthStateChanged, signOut, User as FirebaseUser } from 'firebase/auth';
import { db, auth, firebaseConfig } from './firebase';
import * as XLSX from 'xlsx';

interface DataItem {
  id: string;
  [key: string]: any;
}

interface ErrorBoundaryProps {
  children: ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
  error: Error | null;
}

// Error Boundary Component
class ErrorBoundary extends Component<any, any> {
  public state = {
    hasError: false,
    error: null
  };

  constructor(props: any) {
    super(props);
  }

  static getDerivedStateFromError(error: Error) {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    console.error("ErrorBoundary caught an error:", error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen bg-slate-50 flex items-center justify-center p-6">
          <div className="max-w-md w-full bg-white rounded-[2.5rem] shadow-2xl p-10 text-center">
            <div className="w-20 h-20 bg-red-50 rounded-3xl flex items-center justify-center mx-auto mb-6 text-red-600">
              <AlertCircle className="w-10 h-10" />
            </div>
            <h2 className="text-2xl font-bold text-slate-900 mb-2">Something went wrong</h2>
            <p className="text-slate-500 text-sm mb-8">
              The application encountered a rendering error. This often happens when data types are inconsistent.
            </p>
            <div className="bg-red-50 rounded-2xl p-4 text-left mb-8 overflow-x-auto">
              <code className="text-xs text-red-600 font-mono">
                {this.state.error?.message}
              </code>
            </div>
            <button
              onClick={() => window.location.reload()}
              className="w-full bg-indigo-600 text-white py-4 rounded-2xl font-bold hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-500/20"
            >
              Reload Application
            </button>
          </div>
        </div>
      );
    }

    return (this as any).props.children;
  }
}

export default function App() {
  return (
    <ErrorBoundary>
      <AppContent />
    </ErrorBoundary>
  );
}

function AppContent() {
  const [items, setItems] = useState<any[]>([]);
  const [totalCount, setTotalCount] = useState(0);
  const [lastVisible, setLastVisible] = useState<any>(null);
  const [columns, setColumns] = useState<string[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const deferredSearchQuery = React.useDeferredValue(searchQuery);
  const [user, setUser] = useState<FirebaseUser | null>(null);
  const [editingItem, setEditingItem] = useState<any | null>(null);
  const [isImporting, setIsImporting] = useState(false);
  const [rootCollection, setRootCollection] = useState('bib_orders');
  const [importCollection, setImportCollection] = useState('');
  const [importType, setImportType] = useState<'excel' | 'tsv' | null>(null);
  const [tsvUrl, setTsvUrl] = useState('');
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isTypeManagerOpen, setIsTypeManagerOpen] = useState(false);
  const [pendingTypeChanges, setPendingTypeChanges] = useState<Record<string, string>>({});
  const [visibleCount, setVisibleCount] = useState(20);
  const [uploadProgress, setUploadProgress] = useState(0);
  const [uploadStatus, setUploadStatus] = useState<string | null>(null);
  const [dbStatus, setDbStatus] = useState<'checking' | 'online' | 'offline'>('checking');
  const [dbError, setDbError] = useState<string | null>(null);
  const [deleteConfirm, setDeleteConfirm] = useState<{ type: 'dataset' | 'item', id: string, name?: string } | null>(null);
  const [configForm, setConfigForm] = useState({
    apiKey: firebaseConfig.apiKey || '',
    authDomain: firebaseConfig.authDomain || '',
    projectId: firebaseConfig.projectId || '',
    storageBucket: firebaseConfig.storageBucket || '',
    messagingSenderId: firebaseConfig.messagingSenderId || '',
    appId: firebaseConfig.appId || '',
    firestoreDatabaseId: firebaseConfig.firestoreDatabaseId || '(default)'
  });
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Connection test
  useEffect(() => {
    const testConnection = async () => {
      setDbStatus('checking');
      setDbError(null);
      try {
        // Try to fetch a non-existent doc from a system collection to test connectivity
        await getDocFromServer(doc(db, '_system_', 'connectivity_test'));
        setDbStatus('online');
      } catch (err: any) {
        console.error("Connection Test Failed:", err);
        const errorCode = err.code || 'unknown';
        const errorMessage = err.message || String(err);
        
        // If we get a permission error, we are actually "online" (connected to server)
        if (errorCode === 'permission-denied' || errorMessage.toLowerCase().includes('permission')) {
          setDbStatus('online');
        } else if (errorMessage.toLowerCase().includes('offline') || errorCode === 'unavailable') {
          setDbStatus('offline');
          setDbError(`Error [${errorCode}]: ${errorMessage}`);
          setError("Database Connection Error: The client is offline or the Firebase configuration is incorrect.");
        } else {
          // Some other errors might still mean we reached the server
          setDbStatus('online');
        }
      }
    };
    testConnection();
  }, []);

  // Auth listener
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (u) => {
      setUser(u);
    });
    return () => unsubscribe();
  }, []);

  // Items listener for the current root collection
  useEffect(() => {
    if (!user || !rootCollection || isImporting) {
      if (!isImporting) {
        setItems([]);
        setLoading(false);
      }
      return;
    }

    setLoading(true);
    
    // Get total count
    const fetchCount = async () => {
      try {
        const coll = collection(db, rootCollection);
        const snapshot = await getCountFromServer(coll);
        setTotalCount(snapshot.data().count);
      } catch (e) {
        console.error("Count error:", e);
      }
    };
    fetchCount();

    // Initial query with limit
    const q = query(
      collection(db, rootCollection), 
      orderBy('__name__'), // Order by document ID by default
      limit(1000)
    );

    const unsubscribe = onSnapshot(q, (snapshot) => {
      let metadataDoc: any = null;
      const fetchedItems = snapshot.docs
        .filter(doc => {
          if (doc.id === '_metadata_') {
            metadataDoc = doc.data();
            return false;
          }
          return true;
        })
        .map(doc => {
          const data = doc.data();
          return {
            ...data,
            id: doc.id
          };
        });
      
      setItems(fetchedItems);
      setLastVisible(snapshot.docs[snapshot.docs.length - 1]);
      
      // Use column order from metadata if available, otherwise derive from first item
      if (metadataDoc && metadataDoc.columnOrder) {
        setColumns(metadataDoc.columnOrder);
      } else if (fetchedItems.length > 0) {
        const firstItem = fetchedItems[0];
        const cols = Object.keys(firstItem).filter(k => k !== 'id');
        setColumns(cols);
      } else {
        setColumns([]);
      }
      
      setLoading(false);
    }, (err) => {
      console.error("Firestore Error:", err);
      setError("Database Error: " + err.message);
      setLoading(false);
    });

    return () => unsubscribe();
  }, [rootCollection, user, isImporting]);

  const handleLoadMore = async () => {
    if (!lastVisible || !rootCollection) return;
    setLoading(true);
    try {
      const q = query(
        collection(db, rootCollection),
        orderBy('__name__'),
        startAfter(lastVisible),
        limit(1000)
      );
      
      const snapshot = await getDocs(q);
      const newItems = snapshot.docs
        .filter(doc => doc.id !== '_metadata_')
        .map(doc => ({ ...doc.data(), id: doc.id }));
      
      if (newItems.length > 0) {
        setItems(prev => [...prev, ...newItems]);
        setLastVisible(snapshot.docs[snapshot.docs.length - 1]);
      }
    } catch (err: any) {
      setError("Load More Error: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  const handleLogin = async () => {
    try {
      const provider = new GoogleAuthProvider();
      await signInWithPopup(auth, provider);
    } catch (err: any) {
      console.error("Login Error Details:", err);
      let message = "Login Error: " + (err instanceof Error ? err.message : String(err));
      if (err.code === 'auth/unauthorized-domain') {
        message = "Lỗi: Tên miền này chưa được cấp phép trong Firebase Console. Hãy thêm tên miền Vercel của bạn vào 'Authorized domains' trong Firebase Authentication settings.";
      } else if (err.code === 'auth/popup-blocked') {
        message = "Lỗi: Trình duyệt đã chặn cửa sổ đăng nhập. Hãy cho phép popup và thử lại.";
      }
      setError(message);
    }
  };

  const handleSaveConfig = () => {
    try {
      localStorage.setItem('custom_firebase_config', JSON.stringify(configForm));
      alert("Configuration saved! The application will now reload to apply changes.");
      window.location.reload();
    } catch (e) {
      alert("Error saving configuration: " + e);
    }
  };

  const handleResetConfig = () => {
    if (window.confirm("Reset to default configuration?")) {
      localStorage.removeItem('custom_firebase_config');
      window.location.reload();
    }
  };

  const handleLogout = () => {
    signOut(auth);
  };

  const handleBulkTypeUpdate = async (fieldTypes: Record<string, string>) => {
    if (!user || !rootCollection || items.length === 0) return;
    
    setLoading(true);
    try {
      const BATCH_SIZE = 500;
      for (let i = 0; i < items.length; i += BATCH_SIZE) {
        const batch = writeBatch(db);
        const chunk = items.slice(i, i + BATCH_SIZE);
        
        chunk.forEach((item) => {
          const updatedData: any = { ...item };
          delete updatedData.id;
          
          Object.entries(fieldTypes).forEach(([field, targetType]) => {
            const val = item[field];
            
            if (targetType === 'null') {
              updatedData[field] = null;
              return;
            }

            if (val === undefined || val === null) return;
            
            let convertedVal = val;
            try {
              // Helper to get raw value if it's a Timestamp
              const getRawValue = (v: any) => (v && typeof v.toDate === 'function') ? v.toDate() : v;

              if (targetType === 'number') {
                const raw = getRawValue(val);
                convertedVal = Number(raw);
                if (isNaN(convertedVal)) convertedVal = val;
              } else if (targetType === 'boolean') {
                const raw = getRawValue(val);
                convertedVal = raw === 'true' || raw === true || raw === 1 || raw === '1';
              } else if (targetType === 'string') {
                const raw = getRawValue(val);
                if (raw instanceof Date) {
                  convertedVal = raw.toLocaleString();
                } else if (typeof raw === 'object') {
                  convertedVal = JSON.stringify(raw);
                } else {
                  convertedVal = String(raw);
                }
              } else if (targetType === 'timestamp') {
                const raw = getRawValue(val);
                const date = (raw instanceof Date) ? raw : new Date(raw);
                if (!isNaN(date.getTime())) {
                  convertedVal = Timestamp.fromDate(date);
                }
              } else if (targetType === 'array') {
                const raw = getRawValue(val);
                if (Array.isArray(raw)) {
                  convertedVal = raw;
                } else if (typeof raw === 'string') {
                  try {
                    const parsed = JSON.parse(raw);
                    convertedVal = Array.isArray(parsed) ? parsed : [raw];
                  } catch (e) {
                    convertedVal = [raw];
                  }
                } else {
                  convertedVal = [raw];
                }
              } else if (targetType === 'object') {
                const raw = getRawValue(val);
                if (raw && typeof raw === 'object' && !Array.isArray(raw) && !(raw instanceof Date)) {
                  convertedVal = raw;
                } else if (typeof raw === 'string') {
                  try {
                    const parsed = JSON.parse(raw);
                    convertedVal = (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) ? parsed : { value: raw };
                  } catch (e) {
                    convertedVal = { value: raw };
                  }
                } else {
                  convertedVal = { value: raw };
                }
              }
            } catch (e) {
              console.error(`Error converting ${field} to ${targetType}`, e);
            }
            updatedData[field] = convertedVal;
          });
          
          batch.update(doc(db, rootCollection, item.id), updatedData);
        });
        
        let retryCount = 0;
        const maxRetries = 3;
        
        while (retryCount <= maxRetries) {
          try {
            await batch.commit();
            break;
          } catch (err: any) {
            if (err.code === 'resource-exhausted' || err.code === 'unavailable' || err.message?.includes('quota')) {
              retryCount++;
              if (retryCount > maxRetries) throw err;
              const waitTime = Math.pow(2, retryCount) * 5000; // Exponential backoff: 10s, 20s, 40s
              setUploadStatus(`Rate limit detected! Retrying in ${waitTime/1000}s... (Attempt ${retryCount}/${maxRetries})`);
              await new Promise(resolve => setTimeout(resolve, waitTime));
            } else {
              throw err;
            }
          }
        }
        
        // Wait 3 seconds before next batch if not the last one
        if (i + BATCH_SIZE < items.length) {
          await new Promise(resolve => setTimeout(resolve, 3000));
        }
      }
      
      setIsTypeManagerOpen(false);
      alert(`Successfully updated data types for ${items.length} records!`);
    } catch (err) {
      setError("Bulk Update Error: " + (err instanceof Error ? err.message : String(err)));
    } finally {
      setLoading(false);
    }
  };

  const processImport = async (rows: any[], columns: string[]) => {
    if (!user || !importCollection) return;

    setLoading(true);
    setUploadProgress(0);
    setUploadStatus("Preparing upload...");
    
    try {
      const idKey = columns.find(k => k.toLowerCase() === 'id');
      
      // Save column order to metadata
      await setDoc(doc(db, importCollection, '_metadata_'), {
        columnOrder: columns.filter(c => c.toLowerCase() !== 'id'),
        updatedAt: serverTimestamp()
      });

      const BATCH_SIZE = 500;
      const totalBatches = Math.ceil(rows.length / BATCH_SIZE);
      let completedBatches = 0;
      
      for (let i = 0; i < rows.length; i += BATCH_SIZE) {
        const chunk = rows.slice(i, i + BATCH_SIZE);
        const batch = writeBatch(db);
        
        chunk.forEach((row) => {
          const docId = idKey ? String(row[idKey]) : undefined;
          const itemRef = docId 
            ? doc(db, importCollection, docId)
            : doc(collection(db, importCollection));
          batch.set(itemRef, row);
        });
        
        let retryCount = 0;
        const maxRetries = 3;
        
        while (retryCount <= maxRetries) {
          try {
            await batch.commit();
            break;
          } catch (err: any) {
            if (err.code === 'resource-exhausted' || err.code === 'unavailable' || err.message?.includes('quota')) {
              retryCount++;
              if (retryCount > maxRetries) throw err;
              const waitTime = Math.pow(2, retryCount) * 5000; // Exponential backoff: 10s, 20s, 40s
              setUploadStatus(`Rate limit detected! Retrying in ${waitTime/1000}s... (Attempt ${retryCount}/${maxRetries})`);
              await new Promise(resolve => setTimeout(resolve, waitTime));
            } else {
              throw err;
            }
          }
        }

        completedBatches++;
        setUploadStatus(`Uploaded batch ${completedBatches} of ${totalBatches}. Waiting 3s...`);
        setUploadProgress(Math.round((completedBatches / totalBatches) * 100));
        
        // Wait 3 seconds before next batch if not the last one
        if (completedBatches < totalBatches) {
          await new Promise(resolve => setTimeout(resolve, 3000));
        }
      }

      setUploadStatus("Upload complete!");
      setUploadProgress(100);
      
      setTimeout(() => {
        setRootCollection(importCollection);
        setIsImporting(false);
        setTsvUrl('');
        setUploadStatus(null);
        setUploadProgress(0);
      }, 1500);
    } catch (err) {
      setError("Import Error: " + (err instanceof Error ? err.message : String(err)));
      setUploadStatus("Upload failed");
    } finally {
      setLoading(false);
    }
  };

  const handleTsvImport = async () => {
    if (!importCollection) {
      alert("Please enter a Target Collection Name");
      return;
    }
    if (!tsvUrl) {
      alert("Please enter a TSV URL");
      return;
    }
    
    setLoading(true);
    setUploadProgress(10);
    setUploadStatus("Fetching TSV data...");
    
    try {
      const fetchUrl = `${tsvUrl}${tsvUrl.includes('?') ? '&' : '?'}t=${Date.now()}`;
      const response = await fetch(fetchUrl);

      if (!response.ok) {
        throw new Error(`Failed to fetch TSV: ${response.status} ${response.statusText}`);
      }
      
      setUploadProgress(40);
      setUploadStatus("Parsing TSV data...");
      
      const text = await response.text();
      const lines = text.split(/\r?\n/).filter(line => line.trim() !== '');
      
      if (lines.length <= 1) throw new Error("No data found in TSV.");

      const headers = lines[0].split('\t').map(h => h.trim()).filter(h => h !== '');
      const rows = lines.slice(1).map(line => {
        const values = line.split('\t');
        const rowData: any = {};
        headers.forEach((h, i) => {
          rowData[h] = values[i]?.trim() || '';
        });
        return rowData;
      });

      setUploadProgress(60);
      await processImport(rows, headers);
    } catch (err) {
      setError("TSV Error: " + (err instanceof Error ? err.message : String(err)));
      setUploadStatus("TSV Import failed");
    } finally {
      setLoading(false);
    }
  };

  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || !user || !importCollection) return;

    setLoading(true);
    setUploadProgress(0);
    setUploadStatus("Reading file...");
    
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const data = evt.target?.result;
        const wb = XLSX.read(data, { 
          type: 'array',
          cellDates: true,
          cellNF: false,
          cellText: false
        });
        
        setUploadProgress(30);
        setUploadStatus("Parsing data...");
        
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rawData = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
        
        if (rawData.length === 0) {
          setLoading(false);
          setUploadStatus(null);
          return;
        }

        const headers = (rawData[0] as any[]).map(h => String(h || '').trim()).filter(h => h !== '');
        
        if (headers.length > 1000) {
          throw new Error("Too many columns detected (>1000). Please check your file format.");
        }

        setUploadProgress(50);
        setUploadStatus("Formatting records (this may take a moment)...");

        // Process in chunks to prevent UI freeze
        const CHUNK_SIZE = 2000;
        const jsonData: any[] = [];
        
        const processChunks = async (startIndex: number) => {
          const endIndex = Math.min(startIndex + CHUNK_SIZE, rawData.length);
          
          for (let i = startIndex; i < endIndex; i++) {
            if (i === 0) continue; // Skip header
            const row = rawData[i];
            const obj: any = {};
            let hasData = false;
            
            headers.forEach((header, index) => {
              const val = row[index];
              if (val !== undefined && val !== null && val !== '') {
                obj[header] = val;
                hasData = true;
              }
            });
            
            if (hasData) jsonData.push(obj);
          }

          setUploadProgress(50 + Math.round((endIndex / rawData.length) * 20));
          
          if (endIndex < rawData.length) {
            // Give UI a chance to breathe
            await new Promise(resolve => setTimeout(resolve, 0));
            await processChunks(endIndex);
          }
        };

        await processChunks(0);

        if (jsonData.length === 0) {
          alert("No valid data found in file.");
          setLoading(false);
          setUploadStatus(null);
          return;
        }

        setUploadProgress(70);
        await processImport(jsonData, headers);
      } catch (err) {
        setError("Excel Error: " + (err instanceof Error ? err.message : String(err)));
        setLoading(false);
        setUploadStatus("Error reading file");
      }
    };
    reader.onerror = () => {
      setError("File Reading Error");
      setLoading(false);
      setUploadStatus("Error reading file");
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const handleSaveItem = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!user || !editingItem || !rootCollection) return;

    try {
      const { id, ...dataToSave } = editingItem;
      if (id === 'new') {
        await addDoc(collection(db, rootCollection), dataToSave);
      } else {
        await updateDoc(doc(db, rootCollection, id), dataToSave);
      }
      setEditingItem(null);
    } catch (err) {
      setError("Save Error: " + (err instanceof Error ? err.message : String(err)));
    }
  };

  const handleDeleteItem = async (id: string) => {
    if (!rootCollection) return;
    try {
      await deleteDoc(doc(db, rootCollection, id));
      setDeleteConfirm(null);
    } catch (err) {
      setError("Delete Error: " + (err instanceof Error ? err.message : String(err)));
    }
  };

  const filteredItems = useMemo(() => {
    const filtered = items.filter(item => {
      const searchStr = deferredSearchQuery.toLowerCase();
      return Object.values(item).some(val => 
        String(val).toLowerCase().includes(searchStr)
      );
    });
    return filtered;
  }, [items, deferredSearchQuery]);

  const displayedItems = useMemo(() => {
    return filteredItems.slice(0, visibleCount);
  }, [filteredItems, visibleCount]);

  // Reset pagination when search changes or collection changes
  useEffect(() => {
    setVisibleCount(20);
  }, [deferredSearchQuery, rootCollection]);

  if (!user) {
    return (
      <div className="min-h-screen bg-[#0f172a] flex items-center justify-center p-6">
        <div className="max-w-md w-full bg-white/5 backdrop-blur-xl border border-white/10 rounded-[2.5rem] p-12 text-center">
          <div className="w-20 h-20 bg-indigo-600 rounded-3xl flex items-center justify-center mx-auto mb-8 shadow-2xl shadow-indigo-500/20">
            <Database className="text-white w-10 h-10" />
          </div>
          <h1 className="text-3xl font-bold text-white mb-4">DataHub Pro</h1>
          <p className="text-gray-400 mb-10">Manage multiple datasets with real-time Firestore sync and dynamic schema support.</p>
          <button 
            onClick={handleLogin}
            className="w-full bg-indigo-600 text-white py-4 rounded-2xl font-bold hover:bg-indigo-700 transition-all flex items-center justify-center gap-3 shadow-xl shadow-indigo-500/20"
          >
            <LogIn className="w-5 h-5" />
            Sign In with Google
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="h-screen flex flex-col bg-[#f8fafc] text-[#1e293b] font-sans overflow-hidden">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 px-8 py-4 z-40 flex-none">
        <div className="max-w-[1600px] mx-auto flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center">
                <Database className="text-white w-6 h-6" />
              </div>
              <div>
                <h1 className="text-lg font-bold leading-tight">Flat Data Manager</h1>
                <p className="text-[10px] text-slate-400 font-bold uppercase tracking-widest flex items-center gap-2">
                  Collection: {rootCollection}
                  <span className="bg-slate-100 text-slate-500 px-1.5 py-0.5 rounded-md border border-slate-200">
                    {totalCount} rows
                  </span>
                </p>
              </div>
            </div>
          </div>

          <div className="flex items-center gap-6">
            <div className="flex items-center gap-3 bg-slate-100 px-4 py-2 rounded-2xl border border-slate-200 focus-within:border-indigo-500/30 transition-all">
              <Database className="w-4 h-4 text-indigo-600" />
              <div className="flex flex-col">
                <span className="text-[8px] font-bold text-slate-400 uppercase tracking-widest leading-none mb-1">Current Root</span>
                <input 
                  value={rootCollection}
                  onChange={(e) => setRootCollection(e.target.value.toLowerCase().replace(/[^a-z0-9_]/g, ''))}
                  className="bg-transparent text-xs font-bold text-slate-900 outline-none w-32"
                  placeholder="datasets"
                />
              </div>
            </div>

            <div className="flex items-center gap-2">
              <div className={`w-2 h-2 rounded-full ${dbStatus === 'online' ? 'bg-emerald-500' : dbStatus === 'offline' ? 'bg-red-500' : 'bg-amber-500 animate-pulse'}`} />
              <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400">
                {dbStatus === 'online' ? 'Database Online' : dbStatus === 'offline' ? 'Database Offline' : 'Checking...'}
              </span>
            </div>

            <button 
              onClick={() => setIsSettingsOpen(true)}
              className="p-2 hover:bg-slate-100 rounded-xl text-slate-400 transition-colors"
              title="Settings"
            >
              <Settings className="w-5 h-5" />
            </button>

            <div className="flex items-center gap-3">
              <div className="text-right hidden sm:block">
                <p className="text-sm font-bold text-slate-900">{user.displayName}</p>
                <p className="text-[10px] text-slate-400 font-medium">{user.email}</p>
              </div>
              <img src={user.photoURL || ''} className="w-9 h-9 rounded-full border-2 border-white shadow-sm" alt="User" />
              <button onClick={handleLogout} className="p-2 rounded-xl hover:bg-slate-100 text-slate-400 transition-colors">
                <LogOut className="w-5 h-5" />
              </button>
            </div>
          </div>
        </div>
      </header>

      <main className="flex-1 overflow-y-auto custom-scrollbar">
        <div className="max-w-[1600px] mx-auto px-8 py-10">
        {error && (
          <motion.div 
            initial={{ opacity: 0, y: -20 }}
            animate={{ opacity: 1, y: 0 }}
            className="mb-8 bg-red-50 border border-red-100 rounded-2xl p-6 flex items-start gap-4 text-red-600"
          >
            <AlertCircle className="w-6 h-6 shrink-0" />
            <div className="flex-1">
              <p className="font-bold text-sm">System Error</p>
              <p className="text-sm opacity-80">{error}</p>
            </div>
            <button onClick={() => setError(null)} className="p-1 hover:bg-red-100 rounded-lg transition-colors">
              <X className="w-5 h-5" />
            </button>
          </motion.div>
        )}

        <div className="space-y-8">
          {/* Table Toolbar */}
          <div className="flex flex-col lg:flex-row items-center justify-between gap-6">
            <div className="flex items-center gap-4 flex-1 w-full">
              <div className="relative flex-1">
                <Search className="absolute left-5 top-1/2 -translate-y-1/2 w-5 h-5 text-slate-400" />
                <input
                  type="text"
                  placeholder={`Search in ${rootCollection}...`}
                  className="w-full bg-white border border-slate-200 rounded-2xl py-4 pl-14 pr-6 text-sm focus:ring-4 focus:ring-indigo-500/10 transition-all outline-none shadow-sm"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                />
                <div className="absolute right-5 top-1/2 -translate-y-1/2 flex items-center gap-3">
                  {items.length < totalCount && searchQuery && (
                    <span className="text-[10px] font-bold text-amber-500 uppercase tracking-widest bg-amber-50 px-2 py-1 rounded-lg border border-amber-100">
                      Searching loaded only
                    </span>
                  )}
                  <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest bg-slate-50 px-2 py-1 rounded-lg border border-slate-100">
                    {searchQuery ? `${filteredItems.length} found` : `${totalCount} total rows`}
                  </span>
                </div>
              </div>
            </div>
            
            <div className="flex items-center gap-3 w-full lg:w-auto">
              <button 
                onClick={() => {
                  const initialTypes: Record<string, string> = {};
                  columns.forEach(col => {
                    const firstVal = items[0]?.[col];
                    initialTypes[col] = typeof firstVal === 'number' ? 'number' : 
                                      typeof firstVal === 'boolean' ? 'boolean' :
                                      (firstVal instanceof Timestamp || (firstVal && typeof firstVal.toDate === 'function')) ? 'date' : 'string';
                  });
                  setPendingTypeChanges(initialTypes);
                  setIsTypeManagerOpen(true);
                }}
                className="flex-1 lg:flex-none bg-amber-50 text-amber-600 px-8 py-4 rounded-2xl font-bold hover:bg-amber-100 transition-all flex items-center justify-center gap-3"
                title="Manage Field Data Types"
              >
                <Database className="w-5 h-5" />
                Manage Types
              </button>
              <button 
                onClick={() => {
                  setImportCollection(rootCollection);
                  setIsImporting(true);
                }}
                className="flex-1 lg:flex-none bg-indigo-50 text-indigo-600 px-8 py-4 rounded-2xl font-bold hover:bg-indigo-100 transition-all flex items-center justify-center gap-3"
              >
                <Upload className="w-5 h-5" />
                Import TSV/Excel
              </button>
              <button 
                onClick={() => {
                  const newItem: any = { id: 'new' };
                  columns.forEach(col => newItem[col] = '');
                  setEditingItem(newItem);
                }}
                className="flex-1 lg:flex-none bg-indigo-600 text-white px-8 py-4 rounded-2xl font-bold hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-500/20 flex items-center justify-center gap-3"
              >
                <Plus className="w-5 h-5" />
                Add Record
              </button>
            </div>
          </div>

          {/* Dynamic Table */}
          <div className="bg-white rounded-[2.5rem] border border-slate-200 shadow-sm overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-slate-50/50 border-b border-slate-100">
                    <th className="px-8 py-5 text-[10px] font-bold text-slate-400 uppercase tracking-widest">Doc ID</th>
                    {columns.map(col => (
                      <th key={col} className="px-8 py-5 text-[10px] font-bold text-slate-400 uppercase tracking-widest">
                        {col}
                      </th>
                    ))}
                    <th className="px-8 py-5 text-[10px] font-bold text-slate-400 uppercase tracking-widest text-right">Actions</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-50">
                  <AnimatePresence>
                    {displayedItems.map((item) => (
                      <motion.tr
                        key={item.id}
                        layout
                        initial={{ opacity: 0 }}
                        animate={{ opacity: 1 }}
                        exit={{ opacity: 0 }}
                        className="hover:bg-slate-50/50 transition-colors group"
                      >
                        <td className="px-8 py-5">
                          <span className="text-[10px] font-mono text-slate-400 bg-slate-100 px-2 py-1 rounded-lg">
                            {item.id.substring(0, 12)}
                          </span>
                        </td>
                        {columns.map(col => (
                          <td key={col} className="px-8 py-5">
                            <p className="text-sm font-medium text-slate-700 truncate max-w-[200px]">
                              {(() => {
                                try {
                                  const val = item[col];
                                  if (val === undefined || val === null) return '-';
                                  if (val && typeof val.toDate === 'function') {
                                    return val.toDate().toLocaleString();
                                  }
                                  if (typeof val === 'boolean') return String(val);
                                  if (typeof val === 'object') return JSON.stringify(val);
                                  return String(val);
                                } catch (e) {
                                  return 'Error';
                                }
                              })()}
                            </p>
                          </td>
                        ))}
                        <td className="px-8 py-5 text-right">
                          <div className="flex items-center justify-end gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                            <button 
                              onClick={() => setEditingItem(item)}
                              className="p-2 rounded-xl hover:bg-amber-50 text-amber-500 transition-colors"
                            >
                              <Edit2 className="w-4 h-4" />
                            </button>
                            <button 
                              onClick={() => setDeleteConfirm({ type: 'item', id: item.id })}
                              className="p-2 rounded-xl hover:bg-red-50 text-red-500 transition-colors"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </div>
                        </td>
                      </motion.tr>
                    ))}
                  </AnimatePresence>
                  {items.length < totalCount && (
                    <tr>
                      <td colSpan={columns.length + 2} className="px-8 py-8 text-center bg-slate-50/30">
                        <button 
                          onClick={async () => {
                            if (visibleCount + 20 <= items.length) {
                              setVisibleCount(prev => prev + 20);
                            } else {
                              await handleLoadMore();
                              setVisibleCount(prev => prev + 20);
                            }
                          }}
                          disabled={loading}
                          className="px-8 py-3 bg-white border border-slate-200 rounded-2xl font-bold text-slate-600 hover:bg-slate-50 hover:border-indigo-500 hover:text-indigo-600 transition-all shadow-sm flex items-center gap-2 mx-auto disabled:opacity-50"
                        >
                          {loading ? <RefreshCw className="w-4 h-4 animate-spin" /> : <Plus className="w-4 h-4" />}
                          Load More (Showing {Math.min(visibleCount, items.length)} of {totalCount})
                        </button>
                      </td>
                    </tr>
                  )}
                  {filteredItems.length === 0 && !loading && (
                    <tr>
                      <td colSpan={columns.length + 2} className="px-8 py-32 text-center">
                        <div className="flex flex-col items-center gap-4">
                          <div className="w-20 h-20 bg-slate-50 rounded-full flex items-center justify-center">
                            <Search className="w-10 h-10 text-slate-200" />
                          </div>
                          <p className="text-slate-400 font-medium">No records found in "{rootCollection}"</p>
                        </div>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
        </div>
      </main>

      {/* Import Modal */}
      <AnimatePresence>
        {isImporting && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-6">
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsImporting(false)}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
            />
            <motion.div
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="relative w-full max-w-xl bg-white rounded-[3rem] shadow-2xl p-10"
            >
              <div className="flex items-center justify-between mb-10">
                <div>
                  <h2 className="text-2xl font-bold text-slate-900">Import Data</h2>
                  <p className="text-slate-400 text-sm mt-1">Create a new collection from TSV or Excel.</p>
                </div>
                <button onClick={() => setIsImporting(false)} className="p-2 rounded-full hover:bg-slate-100 text-slate-400">
                  <X className="w-6 h-6" />
                </button>
              </div>

              <div className="space-y-8">
                {uploadStatus && (
                  <div className="bg-indigo-50 rounded-3xl p-8 border border-indigo-100">
                    <div className="flex items-center justify-between mb-4">
                      <span className="text-sm font-bold text-indigo-900">{uploadStatus}</span>
                      <span className="text-sm font-mono font-bold text-indigo-600">{uploadProgress}%</span>
                    </div>
                    <div className="w-full h-3 bg-indigo-100 rounded-full overflow-hidden">
                      <motion.div 
                        initial={{ width: 0 }}
                        animate={{ width: `${uploadProgress}%` }}
                        className="h-full bg-indigo-600 rounded-full"
                      />
                    </div>
                  </div>
                )}

                <div className="grid grid-cols-1 gap-6">
                  <div>
                    <label className="text-[10px] font-bold text-slate-400 uppercase tracking-widest ml-1 mb-2 block">Target Collection Name</label>
                    <input
                      type="text"
                      placeholder="e.g. bib_orders, items, products..."
                      className="w-full bg-slate-50 border-2 border-transparent focus:border-indigo-500/20 focus:bg-white rounded-2xl px-5 py-4 outline-none transition-all"
                      value={importCollection}
                      onChange={(e) => setImportCollection(e.target.value.toLowerCase().replace(/[^a-z0-9_]/g, ''))}
                    />
                    <p className="text-[10px] text-slate-400 mt-2 ml-1 italic">* Data will be imported directly into this collection.</p>
                  </div>
                </div>

                {!importType ? (
                  <div className="grid grid-cols-2 gap-4">
                    <button 
                      onClick={() => setImportType('tsv')}
                      className="p-8 rounded-[2rem] border-2 border-slate-100 hover:border-indigo-500 hover:bg-indigo-50/30 transition-all text-center group"
                    >
                      <RefreshCw className="w-10 h-10 text-slate-300 mx-auto mb-4 group-hover:text-indigo-600" />
                      <span className="font-bold text-slate-900 block">Google Sheets</span>
                      <span className="text-xs text-slate-400">Sync via TSV URL</span>
                    </button>
                    <button 
                      onClick={() => setImportType('excel')}
                      className="p-8 rounded-[2rem] border-2 border-slate-100 hover:border-indigo-500 hover:bg-indigo-50/30 transition-all text-center group"
                    >
                      <FileSpreadsheet className="w-10 h-10 text-slate-300 mx-auto mb-4 group-hover:text-indigo-600" />
                      <span className="font-bold text-slate-900 block">Excel / CSV</span>
                      <span className="text-xs text-slate-400">Upload from local</span>
                    </button>
                  </div>
                ) : importType === 'tsv' ? (
                  <div className="space-y-6">
                    <div className="flex items-center justify-between">
                      <button onClick={() => setImportType(null)} className="text-xs font-bold text-indigo-600 hover:underline">← Back</button>
                      <span className="text-xs text-slate-400 font-bold uppercase tracking-widest">TSV Import</span>
                    </div>
                    <input
                      type="text"
                      placeholder="Paste TSV URL here..."
                      className="w-full bg-slate-50 border-2 border-transparent focus:border-indigo-500/20 focus:bg-white rounded-2xl px-5 py-4 outline-none transition-all"
                      value={tsvUrl}
                      onChange={(e) => setTsvUrl(e.target.value)}
                    />
                    <button 
                      onClick={handleTsvImport}
                      disabled={!tsvUrl || !rootCollection || loading}
                      className="w-full bg-indigo-600 text-white py-5 rounded-2xl font-bold hover:bg-indigo-700 transition-all disabled:opacity-50 shadow-xl shadow-indigo-500/20 flex items-center justify-center gap-2"
                    >
                      {loading ? (
                        <>
                          <RefreshCw className="w-5 h-5 animate-spin" />
                          Processing...
                        </>
                      ) : 'Start Import'}
                    </button>
                  </div>
                ) : (
                  <div className="space-y-6">
                    <div className="flex items-center justify-between">
                      <button onClick={() => setImportType(null)} className="text-xs font-bold text-indigo-600 hover:underline">← Back</button>
                      <span className="text-xs text-slate-400 font-bold uppercase tracking-widest">Excel Upload</span>
                    </div>
                    <div 
                      onClick={() => !importCollection ? alert('Please enter Target Collection Name first') : fileInputRef.current?.click()}
                      className={`border-2 border-dashed rounded-[2rem] p-12 text-center transition-all ${!importCollection ? 'border-slate-100 opacity-50 cursor-not-allowed' : 'border-slate-200 hover:bg-slate-50 cursor-pointer'}`}
                    >
                      <Upload className="w-12 h-12 text-slate-300 mx-auto mb-4" />
                      <p className="font-bold text-slate-900">Click to upload file</p>
                      <p className="text-xs text-slate-400 mt-1">Supports .xlsx, .xls, .csv</p>
                    </div>
                    <input 
                      type="file" 
                      ref={fileInputRef} 
                      className="hidden" 
                      accept=".xlsx, .xls, .csv" 
                      onChange={handleExcelUpload} 
                    />
                  </div>
                )}
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Edit Item Modal */}
      <AnimatePresence>
        {editingItem && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-6">
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setEditingItem(null)}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
            />
            <motion.div
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="relative w-full max-w-2xl bg-white rounded-[3rem] shadow-2xl overflow-hidden"
            >
              <div className="p-10">
                <div className="flex items-center justify-between mb-10">
                  <div>
                    <h2 className="text-2xl font-bold text-slate-900">
                      {editingItem.id === 'new' ? 'New Record' : 'Edit Record'}
                    </h2>
                    <p className="text-xs text-slate-400 font-bold uppercase tracking-widest mt-1">Collection: {rootCollection}</p>
                  </div>
                  <button onClick={() => setEditingItem(null)} className="p-2 rounded-full hover:bg-slate-100 text-slate-400">
                    <X className="w-6 h-6" />
                  </button>
                </div>

                <form onSubmit={handleSaveItem} className="space-y-6">
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-6 max-h-[50vh] overflow-y-auto pr-4 custom-scrollbar">
                    {columns.map(col => (
                      <div key={col}>
                        <label className="text-[10px] font-bold text-slate-400 uppercase tracking-widest ml-1 mb-2 block">{col}</label>
                        <input
                          className="w-full bg-slate-50 border-2 border-transparent focus:border-indigo-500/20 focus:bg-white rounded-2xl px-5 py-4 outline-none transition-all"
                          value={(() => {
                            const val = editingItem[col];
                            if (val === undefined || val === null) return '';
                            if (val instanceof Timestamp || (val && typeof val.toDate === 'function')) {
                              return val.toDate().toISOString().slice(0, 16);
                            }
                            return String(val);
                          })()}
                          onChange={e => setEditingItem({...editingItem, [col]: e.target.value})}
                        />
                      </div>
                    ))}
                  </div>

                  <div className="pt-8 flex gap-4">
                    <button
                      type="submit"
                      className="flex-1 bg-indigo-600 text-white py-5 rounded-2xl font-bold hover:bg-indigo-700 transition-all shadow-xl shadow-indigo-500/20 flex items-center justify-center gap-3"
                    >
                      <Save className="w-6 h-6" />
                      Save Record
                    </button>
                    <button
                      type="button"
                      onClick={() => setEditingItem(null)}
                      className="px-10 bg-slate-100 text-slate-600 py-5 rounded-2xl font-bold hover:bg-slate-200 transition-all"
                    >
                      Cancel
                    </button>
                  </div>
                </form>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Delete Confirmation Modal */}
      <AnimatePresence>
        {deleteConfirm && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-6">
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setDeleteConfirm(null)}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
            />
            <motion.div
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="relative w-full max-w-sm bg-white rounded-[2.5rem] shadow-2xl p-10 text-center"
            >
              <div className="w-20 h-20 bg-red-50 rounded-3xl flex items-center justify-center mx-auto mb-6 text-red-600">
                <Trash2 className="w-10 h-10" />
              </div>
              <h3 className="text-xl font-bold text-slate-900 mb-2">Confirm Delete</h3>
              <p className="text-slate-500 text-sm mb-8">
                {deleteConfirm.type === 'dataset' 
                  ? `Are you sure you want to delete the entire dataset "${deleteConfirm.name}"? This action cannot be undone.`
                  : 'Are you sure you want to delete this record?'}
              </p>
              <div className="flex gap-4">
                <button 
                  onClick={() => handleDeleteItem(deleteConfirm.id)}
                  className="flex-1 bg-red-600 text-white py-4 rounded-2xl font-bold hover:bg-red-700 transition-all shadow-lg shadow-red-500/20"
                >
                  Delete
                </button>
                <button 
                  onClick={() => setDeleteConfirm(null)}
                  className="flex-1 bg-slate-100 text-slate-600 py-4 rounded-2xl font-bold hover:bg-slate-200 transition-all"
                >
                  Cancel
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Type Manager Modal */}
      <AnimatePresence>
        {isTypeManagerOpen && (
          <div className="fixed inset-0 z-[60] flex items-center justify-center p-6">
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsTypeManagerOpen(false)}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
            />
            <motion.div
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="relative w-full max-w-4xl bg-white rounded-[3rem] shadow-2xl overflow-hidden flex flex-col max-h-[90vh]"
            >
              <div className="p-10 border-b border-slate-100">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-4">
                    <div className="w-12 h-12 bg-amber-100 rounded-2xl flex items-center justify-center text-amber-600">
                      <Database className="w-6 h-6" />
                    </div>
                    <div>
                      <h2 className="text-2xl font-bold text-slate-900">Field Type Manager</h2>
                      <p className="text-slate-400 text-sm mt-1">Convert data types across all {items.length} records in "{rootCollection}".</p>
                    </div>
                  </div>
                  <button onClick={() => setIsTypeManagerOpen(false)} className="p-2 rounded-full hover:bg-slate-100 text-slate-400">
                    <X className="w-6 h-6" />
                  </button>
                </div>
              </div>

              <div className="flex-1 overflow-y-auto p-10 custom-scrollbar">
                <div className="space-y-4">
                  <div className="grid grid-cols-12 gap-4 px-4 mb-4">
                    <div className="col-span-4 text-[10px] font-bold text-slate-400 uppercase tracking-widest">Field Name</div>
                    <div className="col-span-3 text-[10px] font-bold text-slate-400 uppercase tracking-widest">Target Type</div>
                    <div className="col-span-5 text-[10px] font-bold text-slate-400 uppercase tracking-widest">Preview (First Record)</div>
                  </div>

                  {columns.map(col => {
                    const firstVal = items[0]?.[col];
                    const currentType = typeof firstVal === 'number' ? 'number' : 
                                      typeof firstVal === 'boolean' ? 'boolean' :
                                      (firstVal instanceof Timestamp || (firstVal && typeof firstVal.toDate === 'function')) ? 'timestamp' : 
                                      Array.isArray(firstVal) ? 'array' :
                                      (firstVal && typeof firstVal === 'object') ? 'object' : 'string';
                    const targetType = pendingTypeChanges[col] || currentType;
                    
                    return (
                      <div key={col} className="grid grid-cols-12 gap-4 items-center p-4 bg-slate-50 rounded-2xl border border-slate-100 hover:border-amber-200 transition-all group">
                        <div className="col-span-4">
                          <span className="font-bold text-slate-900 block truncate">{col}</span>
                          <span className="text-[10px] text-slate-400 uppercase font-bold">Current: {currentType}</span>
                        </div>
                        <div className="col-span-3">
                          <select 
                            className="w-full bg-white border border-slate-200 rounded-xl px-3 py-2 text-xs font-bold text-slate-700 outline-none focus:border-amber-500 transition-all"
                            value={targetType}
                            onChange={(e) => {
                              setPendingTypeChanges(prev => ({ ...prev, [col]: e.target.value }));
                            }}
                          >
                            <option value="string">String</option>
                            <option value="number">Number</option>
                            <option value="boolean">Boolean</option>
                            <option value="timestamp">Timestamp (Firestore)</option>
                            <option value="array">Array (List)</option>
                            <option value="object">Object (Map)</option>
                            <option value="null">Null (Empty)</option>
                          </select>
                        </div>
                        <div className="col-span-5 flex items-center gap-3">
                          <div className="flex-1 bg-white border border-slate-200 rounded-xl px-4 py-2 text-xs font-mono text-slate-500 truncate">
                            {(() => {
                              const val = firstVal;
                              if (val === undefined || val === null) return 'null';
                              
                              try {
                                if (targetType === 'number') return String(Number(val));
                                if (targetType === 'boolean') return String(val === 'true' || val === true || val === 1 || val === '1');
                                if (targetType === 'string') return String(val);
                                if (targetType === 'timestamp') {
                                  const d = (val instanceof Timestamp || (val && typeof val.toDate === 'function')) ? val.toDate() : new Date(val);
                                  return isNaN(d.getTime()) ? 'Invalid Date' : d.toLocaleString();
                                }
                                if (targetType === 'array') return `[Array(${Array.isArray(val) ? val.length : 1})]`;
                                if (targetType === 'object') return '{Object}';
                                if (targetType === 'null') return 'null';
                              } catch (e) { return 'Error'; }
                              return String(val);
                            })()}
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className="p-10 border-t border-slate-100 bg-slate-50/50 flex gap-4">
                <button
                  onClick={() => handleBulkTypeUpdate(pendingTypeChanges)}
                  disabled={loading}
                  className="flex-1 bg-amber-600 text-white py-5 rounded-2xl font-bold hover:bg-amber-700 transition-all shadow-xl shadow-amber-500/20 flex items-center justify-center gap-3 disabled:opacity-50"
                >
                  {loading ? <RefreshCw className="w-6 h-6 animate-spin" /> : <Save className="w-6 h-6" />}
                  Apply Changes to All Records
                </button>
                <button
                  onClick={() => setIsTypeManagerOpen(false)}
                  className="px-10 bg-white border border-slate-200 text-slate-600 py-5 rounded-2xl font-bold hover:bg-slate-50 transition-all"
                >
                  Cancel
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Settings Modal */}
      <AnimatePresence>
        {isSettingsOpen && (
          <div className="fixed inset-0 z-[60] flex items-center justify-center p-6">
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsSettingsOpen(false)}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
            />
            <motion.div
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="relative w-full max-w-md bg-white rounded-[3rem] shadow-2xl overflow-hidden"
            >
              <div className="p-10">
                <div className="flex items-center justify-between mb-8">
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 bg-slate-100 rounded-xl flex items-center justify-center text-slate-600">
                      <Settings className="w-6 h-6" />
                    </div>
                    <h2 className="text-xl font-bold text-slate-900">System Settings</h2>
                  </div>
                  <button onClick={() => setIsSettingsOpen(false)} className="p-2 rounded-full hover:bg-slate-100 text-slate-400">
                    <X className="w-6 h-6" />
                  </button>
                </div>

                <div className="space-y-6">
                  <div className="p-6 bg-slate-50 rounded-3xl border border-slate-100">
                    <div className="flex items-center justify-between mb-4">
                      <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Database Status</span>
                      <div className="flex items-center gap-2">
                        <div className={`w-2 h-2 rounded-full ${dbStatus === 'online' ? 'bg-emerald-500' : dbStatus === 'offline' ? 'bg-red-500' : 'bg-amber-500 animate-pulse'}`} />
                        <span className={`text-[10px] font-bold uppercase tracking-wider ${dbStatus === 'online' ? 'text-emerald-600' : dbStatus === 'offline' ? 'text-red-600' : 'text-amber-600'}`}>
                          {dbStatus === 'online' ? 'Connected' : dbStatus === 'offline' ? 'Offline' : 'Checking...'}
                        </span>
                      </div>
                    </div>
                    
                    <div className="space-y-4 max-h-[40vh] overflow-y-auto pr-2 custom-scrollbar">
                      <div>
                        <label className="text-[9px] font-bold text-slate-400 uppercase tracking-widest block mb-1">API Key</label>
                        <input 
                          className="w-full bg-white border border-slate-200 rounded-xl px-4 py-2 text-xs outline-none focus:border-indigo-500 transition-all"
                          value={configForm.apiKey}
                          onChange={e => setConfigForm({...configForm, apiKey: e.target.value})}
                          placeholder="AIzaSy..."
                        />
                      </div>
                      <div>
                        <label className="text-[9px] font-bold text-slate-400 uppercase tracking-widest block mb-1">Auth Domain</label>
                        <input 
                          className="w-full bg-white border border-slate-200 rounded-xl px-4 py-2 text-xs outline-none focus:border-indigo-500 transition-all"
                          value={configForm.authDomain}
                          onChange={e => setConfigForm({...configForm, authDomain: e.target.value})}
                          placeholder="project.firebaseapp.com"
                        />
                      </div>
                      <div>
                        <label className="text-[9px] font-bold text-slate-400 uppercase tracking-widest block mb-1">Project ID</label>
                        <input 
                          className="w-full bg-white border border-slate-200 rounded-xl px-4 py-2 text-xs outline-none focus:border-indigo-500 transition-all"
                          value={configForm.projectId}
                          onChange={e => setConfigForm({...configForm, projectId: e.target.value})}
                          placeholder="project-id-123"
                        />
                      </div>
                      <div>
                        <label className="text-[9px] font-bold text-slate-400 uppercase tracking-widest block mb-1">Storage Bucket</label>
                        <input 
                          className="w-full bg-white border border-slate-200 rounded-xl px-4 py-2 text-xs outline-none focus:border-indigo-500 transition-all"
                          value={configForm.storageBucket}
                          onChange={e => setConfigForm({...configForm, storageBucket: e.target.value})}
                          placeholder="project.firebasestorage.app"
                        />
                      </div>
                      <div>
                        <label className="text-[9px] font-bold text-slate-400 uppercase tracking-widest block mb-1">Messaging Sender ID</label>
                        <input 
                          className="w-full bg-white border border-slate-200 rounded-xl px-4 py-2 text-xs outline-none focus:border-indigo-500 transition-all"
                          value={configForm.messagingSenderId}
                          onChange={e => setConfigForm({...configForm, messagingSenderId: e.target.value})}
                          placeholder="123456789"
                        />
                      </div>
                      <div>
                        <label className="text-[9px] font-bold text-slate-400 uppercase tracking-widest block mb-1">App ID</label>
                        <input 
                          className="w-full bg-white border border-slate-200 rounded-xl px-4 py-2 text-xs outline-none focus:border-indigo-500 transition-all"
                          value={configForm.appId}
                          onChange={e => setConfigForm({...configForm, appId: e.target.value})}
                          placeholder="1:12345:web:abcde"
                        />
                      </div>
                      <div>
                        <label className="text-[9px] font-bold text-slate-400 uppercase tracking-widest block mb-1">Database ID</label>
                        <input 
                          className="w-full bg-white border border-slate-200 rounded-xl px-4 py-2 text-xs outline-none focus:border-indigo-500 transition-all"
                          value={configForm.firestoreDatabaseId}
                          onChange={e => setConfigForm({...configForm, firestoreDatabaseId: e.target.value})}
                          placeholder="(default)"
                        />
                      </div>
                    </div>
                  </div>

                  <div className="flex gap-3">
                    <button 
                      onClick={handleSaveConfig}
                      className="flex-1 bg-indigo-600 text-white py-4 rounded-2xl font-bold hover:bg-indigo-700 transition-all flex items-center justify-center gap-2 shadow-lg shadow-indigo-500/20"
                    >
                      <Save className="w-4 h-4" />
                      Save & Apply
                    </button>
                    <button 
                      onClick={handleResetConfig}
                      className="px-6 bg-slate-100 text-slate-600 py-4 rounded-2xl font-bold hover:bg-slate-200 transition-all"
                    >
                      Reset
                    </button>
                  </div>

                  {dbError && (
                    <div className="p-4 bg-red-50 border border-red-100 rounded-2xl">
                      <div className="flex items-center gap-2 mb-2">
                        <AlertCircle className="w-4 h-4 text-red-600" />
                        <span className="text-[10px] font-bold text-red-600 uppercase tracking-widest">Connection Error Details</span>
                      </div>
                      <p className="text-[10px] text-red-700 font-mono break-all leading-relaxed">
                        {dbError}
                      </p>
                    </div>
                  )}
                  
                  <p className="text-[10px] text-center text-slate-400">
                    Configuration is saved locally in your browser.
                  </p>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Error Toast */}
      <AnimatePresence>
        {error && (
          <motion.div
            initial={{ opacity: 0, y: 50 }}
            animate={{ opacity: 1, y: 0 }}
            exit={{ opacity: 0, y: 50 }}
            className="fixed bottom-10 right-10 z-[100] bg-red-600 text-white px-8 py-5 rounded-2xl shadow-2xl flex items-center gap-4"
          >
            <AlertCircle className="w-6 h-6" />
            <div>
              <p className="text-sm font-bold">System Error</p>
              <p className="text-xs opacity-80">{error}</p>
            </div>
            <button onClick={() => setError(null)} className="ml-6 p-2 hover:bg-white/20 rounded-xl transition-colors">
              <X className="w-5 h-5" />
            </button>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
