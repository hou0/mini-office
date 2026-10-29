/**
 * 插件文件存储：一个极简的虚拟文件系统。
 *
 * 为什么用 IndexedDB 而不是 `~/.mini-office/plugins/`：
 * 同一套代码在 Web 预览与 Tauri 桌面下都能工作，无需新增 Rust 命令与能力声明。
 * 代价是开发者无法直接往目录里丢文件做热重载——这留到 P2 引入原生目录实现，
 * 届时只需替换本文件的实现，上层接口不变。
 */

const DB_NAME = 'mini-office-plugins';
const DB_VERSION = 1;
const STORE = 'files';

export interface PluginFileStore {
  write(pluginId: string, path: string, bytes: Uint8Array): Promise<void>;
  read(pluginId: string, path: string): Promise<Uint8Array | null>;
  list(pluginId: string): Promise<string[]>;
  remove(pluginId: string): Promise<void>;
  listPluginIds(): Promise<string[]>;
}

interface FileRecord {
  key: string;
  pluginId: string;
  path: string;
  bytes: Uint8Array;
}

function recordKey(pluginId: string, path: string): string {
  return `${pluginId}::${path}`;
}

function openDatabase(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE)) {
        const store = db.createObjectStore(STORE, { keyPath: 'key' });
        store.createIndex('pluginId', 'pluginId', { unique: false });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error('无法打开插件存储'));
  });
}

function requestToPromise<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error('插件存储操作失败'));
  });
}

function transactionDone(tx: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error ?? new Error('插件存储事务失败'));
    tx.onabort = () => reject(tx.error ?? new Error('插件存储事务被中止'));
  });
}

/** IndexedDB 实现。 */
export function createIndexedDbStore(): PluginFileStore {
  let dbPromise: Promise<IDBDatabase> | null = null;
  const db = () => {
    if (!dbPromise) dbPromise = openDatabase();
    return dbPromise;
  };

  return {
    async write(pluginId, path, bytes) {
      const handle = await db();
      const tx = handle.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).put({ key: recordKey(pluginId, path), pluginId, path, bytes } satisfies FileRecord);
      await transactionDone(tx);
    },

    async read(pluginId, path) {
      const handle = await db();
      const tx = handle.transaction(STORE, 'readonly');
      const record = await requestToPromise<FileRecord | undefined>(
        tx.objectStore(STORE).get(recordKey(pluginId, path)),
      );
      await transactionDone(tx);
      return record ? new Uint8Array(record.bytes) : null;
    },

    async list(pluginId) {
      const handle = await db();
      const tx = handle.transaction(STORE, 'readonly');
      const records = await requestToPromise<FileRecord[]>(
        tx.objectStore(STORE).index('pluginId').getAll(pluginId),
      );
      await transactionDone(tx);
      return records.map((record) => record.path).sort();
    },

    async remove(pluginId) {
      const handle = await db();
      const tx = handle.transaction(STORE, 'readwrite');
      const store = tx.objectStore(STORE);
      const keys = await requestToPromise<IDBValidKey[]>(store.index('pluginId').getAllKeys(pluginId));
      for (const key of keys) store.delete(key);
      await transactionDone(tx);
    },

    async listPluginIds() {
      const handle = await db();
      const tx = handle.transaction(STORE, 'readonly');
      const records = await requestToPromise<FileRecord[]>(tx.objectStore(STORE).getAll());
      await transactionDone(tx);
      return [...new Set(records.map((record) => record.pluginId))];
    },
  };
}

/** 内存实现，用于测试环境或 IndexedDB 不可用时降级。 */
export function createMemoryStore(): PluginFileStore {
  const files = new Map<string, Uint8Array>();
  return {
    async write(pluginId, path, bytes) {
      files.set(recordKey(pluginId, path), new Uint8Array(bytes));
    },
    async read(pluginId, path) {
      const value = files.get(recordKey(pluginId, path));
      return value ? new Uint8Array(value) : null;
    },
    async list(pluginId) {
      const prefix = `${pluginId}::`;
      return [...files.keys()]
        .filter((key) => key.startsWith(prefix))
        .map((key) => key.slice(prefix.length))
        .sort();
    },
    async remove(pluginId) {
      const prefix = `${pluginId}::`;
      for (const key of [...files.keys()]) {
        if (key.startsWith(prefix)) files.delete(key);
      }
    },
    async listPluginIds() {
      const ids = new Set<string>();
      for (const key of files.keys()) {
        const index = key.indexOf('::');
        if (index > 0) ids.add(key.slice(0, index));
      }
      return [...ids];
    },
  };
}

export function createPluginFileStore(): PluginFileStore {
  if (typeof indexedDB === 'undefined') return createMemoryStore();
  return createIndexedDbStore();
}
