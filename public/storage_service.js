/**
 * STORAGE SERVICE (Desktop Version)
 * Adapts to use LocalStorage in Electron and Office.settings in Word.
 */

export const StorageService = {
    /**
     * Lưu trữ dữ liệu JSON
     */
    setProjectData: async (key, data) => {
        return new Promise((resolve) => {
            const jsonData = JSON.stringify(data);
            
            if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
                // Office Add-in Mode
                Office.context.document.settings.set(key, jsonData);
                Office.context.document.settings.saveAsync((result) => {
                    resolve(result.status === Office.AsyncResultStatus.Succeeded);
                });
            } else {
                // Desktop / Web Mode
                try {
                    localStorage.setItem(`hoso_${key}`, jsonData);
                    resolve(true);
                } catch (e) {
                    console.error("Storage Error:", e);
                    resolve(false);
                }
            }
        });
    },

    /**
     * Lấy dữ liệu JSON
     */
    getProjectData: async (key) => {
        let jsonData = null;

        if (typeof Office !== 'undefined' && Office.context && Office.context.document) {
            jsonData = Office.context.document.settings.get(key);
        } else {
            jsonData = localStorage.getItem(`hoso_${key}`);
        }

        if (!jsonData) return null;
        try {
            return JSON.parse(jsonData);
        } catch (e) {
            console.error("JSON Parse Error:", e);
            return null;
        }
    },

    /**
     * Lưu Directory Handle (Chỉ hỗ trợ trình duyệt hiện đại / Electron)
     */
    saveFolderHandle: async (handle) => {
        try {
            const db = await StorageService._openDB();
            const tx = db.transaction("handles", "readwrite");
            const store = tx.objectStore("handles");
            await store.put(handle, "exportFolder");
            return true;
        } catch (e) {
            console.error("IDB Save Error:", e);
            return false;
        }
    },

    /**
     * Lấy Directory Handle
     */
    getFolderHandle: async () => {
        try {
            const db = await StorageService._openDB();
            const tx = db.transaction("handles", "readonly");
            const store = tx.objectStore("handles");
            const request = store.get("exportFolder");
            return new Promise((resolve) => {
                request.onsuccess = () => resolve(request.result);
                request.onerror = () => resolve(null);
            });
        } catch (e) {
            return null;
        }
    },

    _openDB: () => {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open("HoSoAddinDB", 1);
            request.onupgradeneeded = (e) => {
                const db = e.target.result;
                if (!db.objectStoreNames.contains("handles")) {
                    db.createObjectStore("handles");
                }
            };
            request.onsuccess = (e) => resolve(e.target.result);
            request.onerror = (e) => reject(e.target.error);
        });
    }
};
