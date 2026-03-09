"use client";
import { create } from "zustand";
import type { RegimenMaster } from "@/types/regimen";

export type AppStep = "upload" | "extracting" | "review-basic" | "review-regimen" | "preview" | "export";

interface AppState {
  step: AppStep;
  uploadedFiles: File[];
  regimenMaster: RegimenMaster | null;
  isLoading: boolean;
  error: string | null;

  setStep: (s: AppStep) => void;
  setUploadedFiles: (files: File[]) => void;
  setRegimenMaster: (data: RegimenMaster) => void;
  updateBasicInfo: (patch: Partial<RegimenMaster["basicInfo"]>) => void;
  setLoading: (v: boolean) => void;
  setError: (e: string | null) => void;
  reset: () => void;
}

export const useAppStore = create<AppState>((set, get) => ({
  step: "upload",
  uploadedFiles: [],
  regimenMaster: null,
  isLoading: false,
  error: null,

  setStep: (s) => set({ step: s }),
  setUploadedFiles: (files) => set({ uploadedFiles: files }),
  setRegimenMaster: (data) => set({ regimenMaster: data }),
  updateBasicInfo: (patch) => {
    const current = get().regimenMaster;
    if (!current) return;
    set({
      regimenMaster: {
        ...current,
        basicInfo: { ...current.basicInfo, ...patch },
      },
    });
  },
  setLoading: (v) => set({ isLoading: v }),
  setError: (e) => set({ error: e }),
  reset: () => set({ step: "upload", uploadedFiles: [], regimenMaster: null, error: null }),
}));
