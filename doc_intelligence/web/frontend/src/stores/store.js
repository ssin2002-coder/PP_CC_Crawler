import { create } from 'zustand';
export const useStore = create((set, get) => ({
  documents: [],
  selectedDocId: null,
  parsedData: {},
  comStatus: 'connecting',
  envStatus: null,
  setDocuments: (docs) => set({ documents: docs }),
  selectDocument: (docId) => set({ selectedDocId: docId }),
  setComStatus: (status) => set({ comStatus: status }),
  setEnvStatus: (status) => set({ envStatus: status }),
  setParsedData: (docId, data) => set((state) => ({
    parsedData: { ...state.parsedData, [docId]: data },
  })),
  fetchParsed: async (docId) => {
    if (get().parsedData[docId]) return;
    try {
      const res = await fetch(`/api/documents/${docId}/parsed`);
      if (res.ok) {
        const data = await res.json();
        get().setParsedData(docId, data);
      }
    } catch (e) { console.error('Failed to fetch parsed data:', e); }
  },
  refetchDocuments: async () => {
    const res = await fetch('/api/documents');
    if (!res.ok) throw new Error(`documents fetch failed: ${res.status}`);
    set({ documents: await res.json() });
  },
  detectFiles: async () => {
    const res = await fetch('/api/documents/detect', { method: 'POST' });
    if (!res.ok) {
      let msg = `${res.status}`;
      try { const j = await res.json(); if (j?.message || j?.detail) msg = j.message || j.detail; } catch {}
      throw new Error(msg);
    }
    return await res.json();
  },
}));
