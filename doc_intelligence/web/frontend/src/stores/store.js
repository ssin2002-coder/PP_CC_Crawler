import { create } from 'zustand';
export const useStore = create((set, get) => ({
  documents: [],
  selectedDocId: null,
  parsedData: {},
  comStatus: 'connecting',
  setDocuments: (docs) => set({ documents: docs }),
  selectDocument: (docId) => set({ selectedDocId: docId }),
  setComStatus: (status) => set({ comStatus: status }),
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
}));
