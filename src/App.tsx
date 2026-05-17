import React, { useState, useEffect, useCallback, useMemo } from 'react';
import { UploadCloud, Settings, Play, Brain, ChevronDown, ChevronUp, FileSpreadsheet, AlertCircle, Save, FolderOpen, LogOut, LogIn, Trash2, Filter, ExternalLink, CheckSquare, Square, Globe } from 'lucide-react';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import { cn } from './lib/utils';
import { auth, db } from './firebase';
import { onAuthStateChanged, signInWithPopup, GoogleAuthProvider, signOut, User } from 'firebase/auth';
import { collection, doc, setDoc, onSnapshot, query, orderBy, serverTimestamp, deleteDoc, getDocFromServer } from 'firebase/firestore';

// --- Types ---
type QueryData = {
  query: string;
  urls: string[];
};

type Group = {
  id: string;
  queries: string[];
  sharedUrls: string[];
  intent?: string;
  recommendation?: string;
};

type Project = {
  id: string;
  name: string;
  threshold: number;
  parsedData: QueryData[];
  groups: Group[];
  createdAt: any;
  updatedAt: any;
};

// --- Error Handling ---
enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: any;
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId,
      providerInfo: auth.currentUser?.providerData.map(provider => ({
        providerId: provider.providerId,
        displayName: provider.displayName,
        email: provider.email,
        photoUrl: provider.photoURL
      })) || []
    },
    operationType,
    path
  };
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

class ErrorBoundary extends React.Component<{children: React.ReactNode}, {hasError: boolean, error: Error | null}> {
  constructor(props: {children: React.ReactNode}) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error: Error) {
    return { hasError: true, error };
  }

  render() {
    if (this.state.hasError) {
      return (
        <div className="min-h-screen bg-red-50 flex items-center justify-center p-6">
          <div className="bg-white p-8 rounded-2xl shadow-xl max-w-lg w-full border border-red-100">
            <div className="flex items-center gap-3 text-red-600 mb-4">
              <AlertCircle className="w-8 h-8" />
              <h1 className="text-2xl font-bold">Произошла ошибка</h1>
            </div>
            <p className="text-neutral-700 mb-4">Что-то пошло не так. Пожалуйста, обновите страницу или обратитесь в поддержку.</p>
            <pre className="bg-neutral-100 p-4 rounded-lg text-xs overflow-auto text-neutral-600 max-h-64">
              {this.state.error?.message}
            </pre>
            <button 
              onClick={() => window.location.reload()}
              className="mt-6 w-full bg-red-600 hover:bg-red-700 text-white py-3 rounded-xl font-medium transition-colors"
            >
              Обновить страницу
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

// --- Parsing Logic ---
async function parseArsenkinFile(file: File): Promise<QueryData[]> {
  return new Promise((resolve, reject) => {
    const processData = (rows: any[][]) => {
      if (rows.length === 0) return resolve([]);

      // Find header row (skip empty rows at the beginning)
      let headerIdx = 0;
      for (let i = 0; i < Math.min(5, rows.length); i++) {
        if (rows[i] && rows[i].filter(Boolean).length > 1) {
          headerIdx = i;
          break;
        }
      }

      const header = rows[headerIdx] || [];
      let queryColIdx = -1;
      const urlColIndices: number[] = [];

      // 1. Try to find columns by header names
      header.forEach((colName: any, idx: number) => {
        if (!colName) return;
        const lower = String(colName).toLowerCase();
        if (lower.includes('запрос') || lower.includes('фраза') || lower.includes('keyword') || lower.includes('query') || lower.includes('ключ')) {
          if (queryColIdx === -1) queryColIdx = idx;
        }
        if (lower.includes('url') || lower.includes('ссылка') || lower.includes('link') || lower.match(/url\s*\d+/) || lower.includes('домен') || lower.includes('сайт')) {
          urlColIndices.push(idx);
        }
      });

      // 2. If header detection failed, guess based on the first data row
      if (rows.length > headerIdx + 1) {
        const firstDataRow = rows[headerIdx + 1];
        
        if (queryColIdx === -1) {
          for (let i = 0; i < firstDataRow.length; i++) {
            const val = String(firstDataRow[i] || '').trim();
            const isUrl = val.match(/^https?:\/\//) || val.match(/^[a-z0-9-]+\.[a-z]{2,}(\/.*)?$/i);
            const isNumber = !isNaN(Number(val)) && val !== '';
            
            if (!isUrl && !isNumber && val.length > 0) {
              queryColIdx = i;
              break;
            }
          }
        }

        if (urlColIndices.length === 0) {
          for (let i = 0; i < firstDataRow.length; i++) {
            const val = String(firstDataRow[i] || '').trim();
            const isUrl = val.match(/^https?:\/\//) || val.match(/^[a-z0-9-]+\.[a-z]{2,}(\/.*)?$/i);
            if (isUrl && i !== queryColIdx) {
              urlColIndices.push(i);
            }
          }
        }
      }

      // 3. Ultimate fallback
      if (queryColIdx === -1) queryColIdx = 0;
      if (urlColIndices.length === 0) {
        for (let i = 0; i < header.length; i++) {
          if (i !== queryColIdx) urlColIndices.push(i);
        }
      }

      const queryMap = new Map<string, Set<string>>();

      for (let i = headerIdx + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0 || (row.length === 1 && !row[0])) continue;

        let query = String(row[queryColIdx] || '').trim();
        
        // If the extracted query is just a number (e.g. row number), try to find the actual text query in this row
        if (query && !isNaN(Number(query))) {
          const textCol = row.find((val, idx) => 
            idx !== queryColIdx && 
            !urlColIndices.includes(idx) && 
            String(val).trim().length > 0 && 
            isNaN(Number(val))
          );
          if (textCol) query = String(textCol).trim();
        }

        if (!query) continue;

        if (!queryMap.has(query)) queryMap.set(query, new Set());

        urlColIndices.forEach(idx => {
          const url = row[idx];
          if (url && typeof url === 'string') {
            const cleanUrl = url.trim()
              .replace(/^https?:\/\//, '')
              .replace(/^www\./, '')
              .replace(/\/$/, '')
              .split('?')[0];
            if (cleanUrl && cleanUrl.includes('.')) {
              queryMap.get(query)!.add(cleanUrl);
            }
          }
        });
      }

      const result: QueryData[] = [];
      queryMap.forEach((urls, query) => {
        if (urls.size > 0) {
          result.push({ query, urls: Array.from(urls) });
        }
      });

      resolve(result);
    };

    if (file.name.toLowerCase().endsWith('.csv')) {
      file.arrayBuffer().then(buffer => {
        let text = '';
        try {
          const decoder = new TextDecoder('utf-8', { fatal: true });
          text = decoder.decode(buffer);
        } catch (e) {
          const decoder = new TextDecoder('windows-1251');
          text = decoder.decode(buffer);
        }
        
        Papa.parse(text, {
          complete: (results) => {
            processData(results.data as any[][]);
          },
          error: (error: any) => {
            reject(new Error("Ошибка парсинга CSV: " + error.message));
          },
          skipEmptyLines: true,
        });
      }).catch(err => reject(new Error("Ошибка чтения файла: " + err.message)));
    } else {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const rows = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1, defval: '' });
          processData(rows);
        } catch (err) {
          console.error(err);
          reject(new Error("Не удалось разобрать файл. Убедитесь, что это корректный экспорт Arsenkin (CSV/XLSX)."));
        }
      };
      reader.onerror = () => reject(new Error("Ошибка чтения файла"));
      reader.readAsArrayBuffer(file);
    }
  });
}

// --- Clustering Logic ---
function clusterQueries(queries: QueryData[], threshold: number): Group[] {
  const groups: Group[] = [];
  const unassigned = [...queries];
  
  unassigned.sort((a, b) => b.urls.length - a.urls.length || a.query.localeCompare(b.query));

  let groupId = 1;

  while (unassigned.length > 0) {
    const center = unassigned.shift()!;
    const currentGroup: Group = {
      id: `G${groupId++}`,
      queries: [center.query],
      sharedUrls: [],
    };

    const toRemove: number[] = [];
    const urlCounts: Record<string, number> = {};
    
    center.urls.forEach(url => {
      urlCounts[url] = 1;
    });

    for (let i = 0; i < unassigned.length; i++) {
      const candidate = unassigned[i];
      const intersection = candidate.urls.filter(url => center.urls.includes(url));
      
      if (intersection.length >= threshold) {
        currentGroup.queries.push(candidate.query);
        candidate.urls.forEach(url => {
          urlCounts[url] = (urlCounts[url] || 0) + 1;
        });
        toRemove.push(i);
      }
    }

    if (currentGroup.queries.length === 1) {
      currentGroup.sharedUrls = center.urls;
    } else {
      currentGroup.sharedUrls = Object.entries(urlCounts)
        .filter(([_, count]) => count > 1)
        .sort((a, b) => b[1] - a[1])
        .map(([url, _]) => url);
    }

    for (let i = toRemove.length - 1; i >= 0; i--) {
      unassigned.splice(toRemove[i], 1);
    }

    groups.push(currentGroup);
  }

  return groups;
}

// --- Main Component ---
function MainApp() {
  const [user, setUser] = useState<User | null>(null);
  const [isAuthReady, setIsAuthReady] = useState(false);
  const [projects, setProjects] = useState<Project[]>([]);
  const [showProjectsModal, setShowProjectsModal] = useState(false);

  const [projectName, setProjectName] = useState<string>('');
  const [file, setFile] = useState<File | null>(null);
  const [parsedData, setParsedData] = useState<QueryData[]>([]);
  const [threshold, setThreshold] = useState<number>(3);
  const [groups, setGroups] = useState<Group[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [analyzeProgress, setAnalyzeProgress] = useState<{current: number, total: number} | null>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [expandedGroups, setExpandedGroups] = useState<Set<string>>(new Set());
  const [excludedDomains, setExcludedDomains] = useState<Set<string>>(new Set());
  const [showDomainsFilter, setShowDomainsFilter] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const resultsRef = React.useRef<HTMLDivElement>(null);

  // Connection Test
  useEffect(() => {
    async function testConnection() {
      try {
        await getDocFromServer(doc(db, 'test', 'connection'));
      } catch (error) {
        if(error instanceof Error && error.message.includes('the client is offline')) {
          console.error("Please check your Firebase configuration.");
        }
      }
    }
    testConnection();
  }, []);

  // Auth Listener
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser);
      setIsAuthReady(true);
    });
    return () => unsubscribe();
  }, []);

  // Projects Listener
  useEffect(() => {
    if (!isAuthReady || !user) {
      setProjects([]);
      return;
    }

    const path = `users/${user.uid}/projects`;
    const q = query(collection(db, path), orderBy('updatedAt', 'desc'));
    
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const loadedProjects: Project[] = [];
      snapshot.forEach((doc) => {
        const data = doc.data();
        try {
          loadedProjects.push({
            id: doc.id,
            name: data.name,
            threshold: data.threshold,
            parsedData: JSON.parse(data.parsedData || '[]'),
            groups: JSON.parse(data.groups || '[]'),
            createdAt: data.createdAt,
            updatedAt: data.updatedAt,
          });
        } catch (e) {
          console.error("Error parsing project data", e);
        }
      });
      setProjects(loadedProjects);
    }, (error) => {
      handleFirestoreError(error, OperationType.LIST, path);
    });

    return () => unsubscribe();
  }, [user, isAuthReady]);

  const handleLogin = async () => {
    try {
      const provider = new GoogleAuthProvider();
      await signInWithPopup(auth, provider);
    } catch (err: any) {
      setError("Ошибка авторизации: " + err.message);
    }
  };

  const handleLogout = async () => {
    try {
      await signOut(auth);
      setParsedData([]);
      setGroups([]);
      setProjectName('');
      setFile(null);
    } catch (err: any) {
      setError("Ошибка выхода: " + err.message);
    }
  };

  const processFile = async (selectedFile: File) => {
    setFile(selectedFile);
    if (!projectName) setProjectName(selectedFile.name.replace(/\.[^/.]+$/, ""));
    setError(null);
    setIsProcessing(true);
    
    try {
      const data = await parseArsenkinFile(selectedFile);
      setParsedData(data);
      if (data.length === 0) {
        setError("В файле не найдено запросов и URL. Проверьте формат.");
      }
    } catch (err: any) {
      setError(err.message || "Ошибка при обработке файла");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = e.target.files?.[0];
    if (selectedFile) processFile(selectedFile);
  };

  const handleDragOver = (e: React.DragEvent<HTMLLabelElement>) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent<HTMLLabelElement>) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent<HTMLLabelElement>) => {
    e.preventDefault();
    setIsDragging(false);
    const selectedFile = e.dataTransfer.files?.[0];
    if (selectedFile) processFile(selectedFile);
  };

  const domainStats = useMemo(() => {
    if (parsedData.length === 0) return [];
    const stats: Record<string, number> = {};
    parsedData.forEach(item => {
      item.urls.forEach(url => {
        let domain = url;
        if (domain.includes('/')) {
          domain = domain.split('/')[0];
        }
        stats[domain] = (stats[domain] || 0) + 1;
      });
    });
    return Object.entries(stats).sort((a, b) => b[1] - a[1]);
  }, [parsedData]);

  const toggleDomain = (domain: string) => {
    const next = new Set(excludedDomains);
    if (next.has(domain)) next.delete(domain);
    else next.add(domain);
    setExcludedDomains(next);
  };

  const handleCluster = () => {
    if (parsedData.length === 0) return;
    
    const filteredData = parsedData.map(item => ({
      query: item.query,
      urls: item.urls.filter(url => {
        let domain = url;
        if (domain.includes('/')) domain = domain.split('/')[0];
        return !excludedDomains.has(domain);
      })
    }));

    const newGroups = clusterQueries(filteredData, threshold);
    setGroups(newGroups);
    setExpandedGroups(new Set());
    setTimeout(() => {
      resultsRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }, 100);
  };

  const handleAIAnalysis = async () => {
    if (groups.length === 0) return;
    setIsAnalyzing(true);
    setError(null);
    setAnalyzeProgress({ current: 0, total: groups.length });

    try {
      const batchSize = 10;
      const updatedGroups = [...groups];

      for (let i = 0; i < groups.length; i += batchSize) {
        const batch = groups.slice(i, i + batchSize);
        const payload = batch.map(g => ({
          id: g.id,
          queries: g.queries.slice(0, 10) // Limit to top 10 queries per group to save tokens and prevent truncation
        }));

        const prompt = `
Ты опытный SEO-специалист. Я передаю тебе список групп поисковых запросов, которые были сгруппированы на основе пересечения URL в ТОП-10 выдачи.
Для каждой группы определи основной интент запросов (Информационный, Коммерческий, Транзакционный, Навигационный) и дай краткую рекомендацию (1-2 предложения) о том, какую страницу лучше создать для этой группы (например, "Статья в блог", "Карточка товара", "Категория каталога", "Главная страница", "Страница услуги").

Ответь строго в формате JSON. Это должен быть массив объектов.
Пример:
[
  {
    "id": "G1",
    "intent": "Коммерческий",
    "recommendation": "Создать страницу категории каталога с фильтрами и листингом товаров."
  }
]

Данные для анализа:
${JSON.stringify(payload, null, 2)}
        `;

        const response = await fetch("https://routerai.ru/api/v1/chat/completions", {
          method: "POST",
          headers: {
            "Authorization": "sk-idWLIk8WBHJJiwn-Y2oyMNdW0ckjsfIa",
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            model: "anthropic/claude-sonnet-4.6",
            messages: [
              { role: "user", content: prompt }
            ]
          })
        });

        if (!response.ok) {
          throw new Error(`Ошибка API: ${response.status} ${response.statusText}`);
        }

        const data = await response.json();
        let text = data.choices?.[0]?.message?.content || "";
        
        // Extract JSON array from markdown if present
        const jsonMatch = text.match(/\[[\s\S]*\]/);
        if (jsonMatch) {
          text = jsonMatch[0];
        }

        if (text) {
          const aiResults = JSON.parse(text);
          aiResults.forEach((result: any) => {
            const groupIndex = updatedGroups.findIndex(g => g.id === result.id);
            if (groupIndex !== -1) {
              updatedGroups[groupIndex] = {
                ...updatedGroups[groupIndex],
                intent: result.intent,
                recommendation: result.recommendation
              };
            }
          });
          setGroups([...updatedGroups]);
        }
        
        setAnalyzeProgress({ current: Math.min(i + batchSize, groups.length), total: groups.length });
        
        // Add a delay between batches to avoid hitting rate limits
        if (i + batchSize < groups.length) {
          await new Promise(resolve => setTimeout(resolve, 2000));
        }
      }
    } catch (err: any) {
      console.error(err);
      setError("Ошибка при анализе ИИ: " + (err.message || "Неизвестная ошибка"));
    } finally {
      setIsAnalyzing(false);
      setAnalyzeProgress(null);
    }
  };

  const handleSaveProject = async () => {
    if (!user) return;
    if (parsedData.length === 0 && groups.length === 0) {
      setError("Нет данных для сохранения");
      return;
    }

    setIsSaving(true);
    setError(null);

    try {
      const projectId = Date.now().toString();
      const path = `users/${user.uid}/projects/${projectId}`;
      
      const parsedDataStr = JSON.stringify(parsedData);
      const groupsStr = JSON.stringify(groups);

      if (parsedDataStr.length > 1000000 || groupsStr.length > 1000000) {
        throw new Error("Проект слишком большой для сохранения (лимит 1MB).");
      }

      await setDoc(doc(db, 'users', user.uid, 'projects', projectId), {
        userId: user.uid,
        name: projectName || 'Новый проект',
        threshold,
        parsedData: parsedDataStr,
        groups: groupsStr,
        createdAt: serverTimestamp(),
        updatedAt: serverTimestamp()
      });
      
      alert("Проект успешно сохранен!");
    } catch (err: any) {
      handleFirestoreError(err, OperationType.CREATE, `users/${user.uid}/projects`);
    } finally {
      setIsSaving(false);
    }
  };

  const handleLoadProject = (project: Project) => {
    setProjectName(project.name);
    setThreshold(project.threshold);
    setParsedData(project.parsedData);
    setGroups(project.groups);
    setExpandedGroups(new Set());
    setShowProjectsModal(false);
    setFile(null);
  };

  const handleDeleteProject = async (projectId: string, e: React.MouseEvent) => {
    e.stopPropagation();
    if (!user) return;
    
    try {
      const path = `users/${user.uid}/projects/${projectId}`;
      await deleteDoc(doc(db, 'users', user.uid, 'projects', projectId));
    } catch (err: any) {
      handleFirestoreError(err, OperationType.DELETE, `users/${user.uid}/projects/${projectId}`);
    }
  };

  const toggleGroup = (id: string) => {
    const newExpanded = new Set(expandedGroups);
    if (newExpanded.has(id)) {
      newExpanded.delete(id);
    } else {
      newExpanded.add(id);
    }
    setExpandedGroups(newExpanded);
  };

  if (!isAuthReady) {
    return <div className="min-h-screen flex items-center justify-center bg-neutral-50">Загрузка...</div>;
  }

  return (
    <div className="min-h-screen bg-neutral-50 text-neutral-900 font-sans p-6 md:p-12">
      <div className="max-w-6xl mx-auto space-y-8">
        
        {/* Header */}
        <header className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="space-y-2">
            <h1 className="text-3xl font-bold tracking-tight text-neutral-900">SEO Кластеризатор</h1>
            <p className="text-neutral-500">
              Группировка поисковых запросов на основе пересечения URL в ТОП-10 выдачи с AI-анализом интента.
            </p>
          </div>
          
          <div className="flex items-center gap-3">
            {user ? (
              <>
                <button 
                  onClick={() => setShowProjectsModal(true)}
                  className="flex items-center gap-2 px-4 py-2 bg-white border border-neutral-200 rounded-xl hover:bg-neutral-50 transition-colors font-medium text-sm shadow-sm"
                >
                  <FolderOpen className="w-4 h-4" />
                  Мои проекты
                </button>
                <button 
                  onClick={handleLogout}
                  className="flex items-center gap-2 px-4 py-2 bg-white border border-neutral-200 rounded-xl hover:bg-neutral-50 transition-colors font-medium text-sm shadow-sm text-red-600 hover:text-red-700"
                >
                  <LogOut className="w-4 h-4" />
                  Выйти
                </button>
              </>
            ) : (
              <button 
                onClick={handleLogin}
                className="flex items-center gap-2 px-6 py-2.5 bg-blue-600 text-white rounded-xl hover:bg-blue-700 transition-colors font-medium shadow-sm"
              >
                <LogIn className="w-4 h-4" />
                Войти через Google
              </button>
            )}
          </div>
        </header>

        {/* Main Controls */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          
          {/* Step 1: Upload */}
          <div className="bg-white p-6 rounded-2xl shadow-sm border border-neutral-200 space-y-4">
            <div className="flex items-center gap-2 text-lg font-semibold">
              <div className="bg-blue-100 text-blue-600 w-8 h-8 rounded-full flex items-center justify-center text-sm">1</div>
              Загрузка данных
            </div>
            <p className="text-sm text-neutral-500">Загрузите выгрузку из Arsenkin (CSV или XLSX)</p>
            
            <label 
              className={cn(
                "flex flex-col items-center justify-center w-full h-32 border-2 border-dashed rounded-xl cursor-pointer transition-colors",
                isDragging ? "border-blue-500 bg-blue-50" : "border-neutral-300 bg-neutral-50 hover:bg-neutral-100"
              )}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
            >
              <div className="flex flex-col items-center justify-center pt-5 pb-6 pointer-events-none">
                <UploadCloud className={cn("w-8 h-8 mb-3", isDragging ? "text-blue-500" : "text-neutral-400")} />
                <p className="mb-2 text-sm text-neutral-500 text-center px-4">
                  <span className="font-semibold">Нажмите для загрузки</span> или перетащите файл
                </p>
              </div>
              <input type="file" className="hidden" accept=".csv, .xlsx, .xls" onChange={handleFileUpload} />
            </label>
            
            {(file || projectName) && (
              <div className="space-y-2">
                <input 
                  type="text" 
                  value={projectName}
                  onChange={(e) => setProjectName(e.target.value)}
                  placeholder="Название проекта"
                  className="w-full px-3 py-2 border border-neutral-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
                <div className="flex items-center gap-2 text-sm text-green-600 bg-green-50 p-2 rounded-lg">
                  <FileSpreadsheet className="w-4 h-4" />
                  <span className="truncate">{file ? file.name : "Загружено из облака"}</span>
                  <span className="ml-auto font-medium">{parsedData.length} запросов</span>
                </div>
              </div>
            )}
          </div>

          {/* Step 2: Settings */}
          <div className="bg-white p-6 rounded-2xl shadow-sm border border-neutral-200 space-y-4">
            <div className="flex items-center gap-2 text-lg font-semibold">
              <div className="bg-blue-100 text-blue-600 w-8 h-8 rounded-full flex items-center justify-center text-sm">2</div>
              Настройки
            </div>
            <p className="text-sm text-neutral-500">Порог пересечения URL для объединения в группу</p>
            
            <div className="space-y-6 pt-4">
              <div className="flex items-center justify-between">
                <span className="text-sm font-medium">Количество совпадений:</span>
                <span className="text-lg font-bold text-blue-600 bg-blue-50 px-3 py-1 rounded-lg">{threshold}</span>
              </div>
              <input 
                type="range" 
                min="1" max="10" 
                value={threshold} 
                onChange={(e) => setThreshold(parseInt(e.target.value))}
                className="w-full h-2 bg-neutral-200 rounded-lg appearance-none cursor-pointer accent-blue-600"
              />
              <div className="flex justify-between text-xs text-neutral-400">
                <span>Слабая связь (1)</span>
                <span>Сильная связь (10)</span>
              </div>
            </div>
          </div>

          {/* Step 3: Actions */}
          <div className="bg-white p-6 rounded-2xl shadow-sm border border-neutral-200 space-y-4 flex flex-col">
            <div className="flex items-center gap-2 text-lg font-semibold">
              <div className="bg-blue-100 text-blue-600 w-8 h-8 rounded-full flex items-center justify-center text-sm">3</div>
              Анализ
            </div>
            <p className="text-sm text-neutral-500 mb-auto">Запустите кластеризацию и AI-анализ</p>
            
            <div className="space-y-3 mt-auto">
              <button 
                onClick={handleCluster}
                disabled={parsedData.length === 0 || isProcessing}
                className="w-full flex items-center justify-center gap-2 bg-neutral-900 hover:bg-neutral-800 text-white py-3 px-4 rounded-xl font-medium transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              >
                <Play className="w-4 h-4" />
                Сгруппировать запросы
              </button>
              
              <button 
                onClick={handleAIAnalysis}
                disabled={groups.length === 0 || isAnalyzing}
                className="w-full flex items-center justify-center gap-2 bg-blue-600 hover:bg-blue-700 text-white py-3 px-4 rounded-xl font-medium transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
              >
                <Brain className="w-4 h-4" />
                {isAnalyzing 
                  ? (analyzeProgress ? `Анализ... (${analyzeProgress.current} из ${analyzeProgress.total})` : "Нейросеть анализирует...") 
                  : "Анализ интента (AI)"}
              </button>

              {user && (
                <button 
                  onClick={handleSaveProject}
                  disabled={isSaving || (parsedData.length === 0 && groups.length === 0)}
                  className="w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white py-3 px-4 rounded-xl font-medium transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <Save className="w-4 h-4" />
                  {isSaving ? "Сохранение..." : "Сохранить проект"}
                </button>
              )}
            </div>
          </div>
        </div>

        {/* Domain Filter */}
        {parsedData.length > 0 && (
          <div className="bg-white rounded-2xl shadow-sm border border-neutral-200 overflow-hidden">
            <div 
              className="p-5 flex items-center justify-between cursor-pointer hover:bg-neutral-50 transition-colors"
              onClick={() => setShowDomainsFilter(!showDomainsFilter)}
            >
              <div className="flex items-center gap-3">
                <div className="bg-purple-100 text-purple-600 w-8 h-8 rounded-full flex items-center justify-center">
                  <Filter className="w-4 h-4" />
                </div>
                <div>
                  <h3 className="text-lg font-semibold">Фильтрация доменов</h3>
                  <p className="text-sm text-neutral-500">
                    Исключите неподходящие сайты (маркетплейсы, информационники) для более точной кластеризации
                  </p>
                </div>
              </div>
              <div className="flex items-center gap-3">
                <span className="text-sm font-medium bg-neutral-100 px-3 py-1 rounded-full">
                  {domainStats.length} доменов
                </span>
                {showDomainsFilter ? <ChevronUp className="w-5 h-5 text-neutral-400" /> : <ChevronDown className="w-5 h-5 text-neutral-400" />}
              </div>
            </div>

            {showDomainsFilter && (
              <div className="p-5 border-t border-neutral-100 bg-neutral-50/50">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-4">
                  <p className="text-sm text-neutral-600">
                    Отключите галочки у доменов, которые не должны учитываться при кластеризации.
                  </p>
                  <div className="flex gap-2">
                    <button 
                      onClick={() => setExcludedDomains(new Set())}
                      className="text-sm font-medium text-blue-600 hover:text-blue-700 px-3 py-1.5 bg-blue-50 hover:bg-blue-100 rounded-lg transition-colors"
                    >
                      Выбрать все
                    </button>
                    <button 
                      onClick={() => setExcludedDomains(new Set(domainStats.map(d => d[0])))}
                      className="text-sm font-medium text-neutral-600 hover:text-neutral-700 px-3 py-1.5 bg-neutral-100 hover:bg-neutral-200 rounded-lg transition-colors"
                    >
                      Снять выделение
                    </button>
                  </div>
                </div>
                <div className="max-h-80 overflow-y-auto custom-scrollbar border border-neutral-200 rounded-xl bg-white">
                  <table className="w-full text-left text-sm">
                    <thead className="bg-neutral-50 sticky top-0 border-b border-neutral-200 z-10">
                      <tr>
                        <th className="px-4 py-3 font-semibold text-neutral-600 w-12 text-center">Учитывать</th>
                        <th className="px-4 py-3 font-semibold text-neutral-600">Домен</th>
                        <th className="px-4 py-3 font-semibold text-neutral-600 w-24 text-center">Встречается</th>
                        <th className="px-4 py-3 font-semibold text-neutral-600 w-24 text-center">Ссылка</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-neutral-100">
                      {domainStats.map(([domain, count]) => {
                        const isExcluded = excludedDomains.has(domain);
                        return (
                          <tr 
                            key={domain} 
                            className={cn("hover:bg-neutral-50 transition-colors cursor-pointer", isExcluded && "bg-neutral-50/50")}
                            onClick={() => toggleDomain(domain)}
                          >
                            <td className="px-4 py-3 text-center">
                              {isExcluded ? (
                                <Square className="w-5 h-5 text-neutral-300 inline-block" />
                              ) : (
                                <CheckSquare className="w-5 h-5 text-blue-600 inline-block" />
                              )}
                            </td>
                            <td className={cn("px-4 py-3 font-medium break-all", isExcluded ? "text-neutral-400 line-through" : "text-neutral-800")}>
                              {domain}
                            </td>
                            <td className="px-4 py-3 text-center">
                              <span className={cn("px-2 py-0.5 rounded-full text-xs font-bold", isExcluded ? "bg-neutral-100 text-neutral-400" : "bg-neutral-100 text-neutral-600")}>
                                {count}
                              </span>
                            </td>
                            <td className="px-4 py-3 text-center" onClick={(e) => e.stopPropagation()}>
                              <a 
                                href={`https://${domain}`}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="inline-flex items-center justify-center p-1.5 text-neutral-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
                                title="Открыть сайт"
                              >
                                <ExternalLink className="w-5 h-5" />
                              </a>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Error Message */}
        {error && (
          <div className="bg-red-50 border border-red-200 text-red-700 p-4 rounded-xl flex items-start gap-3">
            <AlertCircle className="w-5 h-5 mt-0.5 flex-shrink-0" />
            <p>{error}</p>
          </div>
        )}

        {/* Results */}
        {groups.length > 0 && (
          <div className="space-y-6" ref={resultsRef}>
            <div className="flex items-center justify-between">
              <h2 className="text-2xl font-bold tracking-tight">Результаты кластеризации</h2>
              <div className="bg-white px-4 py-2 rounded-full shadow-sm border border-neutral-200 text-sm font-medium">
                Рекомендуется страниц: <span className="text-blue-600 font-bold">{groups.length}</span>
              </div>
            </div>

            <div className={cn("border p-4 rounded-xl flex items-start gap-3", 
              groups.length === 1 ? "bg-emerald-50 border-emerald-200 text-emerald-800" : "bg-blue-50 border-blue-200 text-blue-800"
            )}>
              <Brain className={cn("w-5 h-5 mt-0.5 flex-shrink-0", groups.length === 1 ? "text-emerald-600" : "text-blue-600")} />
              <div>
                <p className="font-medium text-lg mb-1">
                  {groups.length === 1 
                    ? "Все запросы можно продвигать на одной странице!" 
                    : `Запросы нужно разделить на ${groups.length} страниц.`}
                </p>
                <p className={cn("text-sm", groups.length === 1 ? "text-emerald-700" : "text-blue-700")}>
                  {groups.length === 1 
                    ? `Все ${parsedData.length} запросов имеют достаточно общих URL в ТОП-10 (≥${threshold}), поэтому поисковики считают их синонимичными.`
                    : `Выдача по этим запросам отличается. Мы сгруппировали их в ${groups.length} кластеров. Каждый кластер — это отдельная страница (статья, категория или услуга).`}
                </p>
                {!groups[0]?.intent && (
                  <p className={cn("text-sm mt-2 font-medium", groups.length === 1 ? "text-emerald-800" : "text-blue-800")}>
                    Нажмите кнопку «Анализ интента (AI)» выше, чтобы нейросеть дала рекомендации по созданию страниц.
                  </p>
                )}
              </div>
            </div>

            <div className="grid gap-4">
              {groups.map((group, index) => {
                const isExpanded = expandedGroups.has(group.id);
                return (
                  <div key={group.id} className="bg-white rounded-2xl shadow-sm border border-neutral-200 overflow-hidden transition-all">
                    {/* Group Header */}
                    <div 
                      className="p-5 flex items-center justify-between cursor-pointer hover:bg-neutral-50 transition-colors"
                      onClick={() => toggleGroup(group.id)}
                    >
                      <div className="flex-1 min-w-0 pr-4">
                        <div className="flex items-center gap-3 mb-1">
                          <span className="bg-blue-100 text-blue-700 text-xs font-bold px-2.5 py-1 rounded-md">
                            Страница {index + 1}
                          </span>
                          <h3 className="text-lg font-semibold truncate" title={group.queries[0]}>
                            {group.queries[0]}
                          </h3>
                          <span className="text-sm text-neutral-500 whitespace-nowrap">
                            +{group.queries.length - 1} запросов
                          </span>
                        </div>
                        
                        {group.intent && (
                          <div className={cn("flex gap-2 mt-2", isExpanded ? "flex-col items-start" : "items-center")}>
                            <span className={cn(
                              "text-xs font-medium px-2.5 py-1 rounded-full whitespace-nowrap",
                              group.intent?.toLowerCase().includes('коммерч') ? "bg-amber-100 text-amber-800" :
                              group.intent?.toLowerCase().includes('информ') ? "bg-blue-100 text-blue-800" :
                              group.intent?.toLowerCase().includes('транзакц') ? "bg-emerald-100 text-emerald-800" :
                              "bg-purple-100 text-purple-800"
                            )}>
                              {group.intent}
                            </span>
                            <p className={cn("text-sm text-neutral-600", isExpanded ? "whitespace-normal" : "truncate")}>{group.recommendation}</p>
                          </div>
                        )}
                      </div>
                      
                      <div className="flex items-center justify-center w-10 h-10 rounded-full bg-neutral-50 text-neutral-400">
                        {isExpanded ? <ChevronUp className="w-5 h-5" /> : <ChevronDown className="w-5 h-5" />}
                      </div>
                    </div>

                    {/* Group Content */}
                    {isExpanded && (
                      <div className="border-t border-neutral-100 p-5 bg-neutral-50/50 grid md:grid-cols-2 gap-6">
                        <div>
                          <h4 className="text-sm font-semibold text-neutral-900 mb-3 flex items-center justify-between">
                            Запросы в группе
                            <span className="bg-neutral-200 text-neutral-700 text-xs py-0.5 px-2 rounded-full">{group.queries.length}</span>
                          </h4>
                          <ul className="space-y-2 max-h-64 overflow-y-auto pr-2 custom-scrollbar">
                            {group.queries.map((q, idx) => (
                              <li key={idx} className="text-sm bg-white p-2 rounded-lg border border-neutral-200 shadow-sm">
                                {q}
                              </li>
                            ))}
                          </ul>
                        </div>
                        
                        <div>
                          <h4 className="text-sm font-semibold text-neutral-900 mb-3 flex items-center justify-between">
                            Пересекающиеся URL в ТОП-10
                            <span className="bg-neutral-200 text-neutral-700 text-xs py-0.5 px-2 rounded-full">{group.sharedUrls.length}</span>
                          </h4>
                          <ul className="space-y-2 max-h-64 overflow-y-auto pr-2 custom-scrollbar">
                            {group.sharedUrls.map((url, idx) => (
                              <li key={idx} className="text-xs bg-white p-2 rounded-lg border border-neutral-200 shadow-sm truncate" title={url}>
                                {url}
                              </li>
                            ))}
                            {group.sharedUrls.length === 0 && (
                              <li className="text-sm text-neutral-500 italic p-2">Нет общих URL (группа состоит из 1 запроса)</li>
                            )}
                          </ul>
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* Projects Modal */}
        {showProjectsModal && (
          <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
            <div className="bg-white rounded-2xl shadow-xl w-full max-w-2xl max-h-[80vh] flex flex-col">
              <div className="p-6 border-b border-neutral-100 flex items-center justify-between">
                <h2 className="text-xl font-bold">Мои проекты</h2>
                <button onClick={() => setShowProjectsModal(false)} className="text-neutral-500 hover:text-neutral-900">
                  Закрыть
                </button>
              </div>
              <div className="p-6 overflow-y-auto flex-1 custom-scrollbar">
                {projects.length === 0 ? (
                  <p className="text-center text-neutral-500 py-8">У вас пока нет сохраненных проектов.</p>
                ) : (
                  <div className="space-y-3">
                    {projects.map(project => (
                      <div key={project.id} className="flex items-center justify-between p-4 border border-neutral-200 rounded-xl hover:border-blue-300 transition-colors bg-neutral-50/50">
                        <div className="cursor-pointer flex-1" onClick={() => handleLoadProject(project)}>
                          <h3 className="font-semibold text-neutral-900">{project.name}</h3>
                          <p className="text-sm text-neutral-500">
                            {project.parsedData.length} запросов • Порог: {project.threshold} • {new Date(project.createdAt?.seconds * 1000).toLocaleDateString()}
                          </p>
                        </div>
                        <button 
                          onClick={(e) => handleDeleteProject(project.id, e)}
                          className="p-2 text-red-500 hover:bg-red-50 rounded-lg transition-colors"
                          title="Удалить проект"
                        >
                          <Trash2 className="w-5 h-5" />
                        </button>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export default function App() {
  return (
    <ErrorBoundary>
      <MainApp />
    </ErrorBoundary>
  );
}
