import { useEffect, useCallback } from 'react';
import useStore from '../stores/store.js';
import * as api from '../utils/api.js';

// =============================================
// 규칙 관리 훅
// 규칙 CRUD 및 API 동기화
// =============================================

export function useRules() {
  const rules = useStore((s) => s.rules);
  const setRules = useStore((s) => s.setRules);
  const addRule = useStore((s) => s.addRule);
  const updateRuleInStore = useStore((s) => s.updateRule);
  const deleteRuleFromStore = useStore((s) => s.deleteRule);
  const toggleRuleInStore = useStore((s) => s.toggleRule);

  // 마운트 시 규칙 목록 로드
  useEffect(() => {
    async function load() {
      const { data, error } = await api.fetchRules();
      if (error) {
        console.error('[규칙] 로드 실패:', error);
        return;
      }
      if (data) {
        setRules(data.rules || []);
      }
    }
    load();
  }, [setRules]);

  // 규칙 생성 - API 호출 후 스토어 업데이트
  const createRule = useCallback(async (ruleData) => {
    const { data, error } = await api.createRule(ruleData);
    if (error) {
      console.error('[규칙] 생성 실패:', error);
      return { error };
    }
    if (data && data.rule) {
      addRule(data.rule);
    }
    return { data };
  }, [addRule]);

  // 규칙 수정 - API 호출 후 스토어 업데이트
  const updateRule = useCallback(async (id, updates) => {
    const { data, error } = await api.updateRule(id, updates);
    if (error) {
      console.error('[규칙] 수정 실패:', error);
      return { error };
    }
    if (data && data.rule) {
      updateRuleInStore(id, data.rule);
    } else {
      updateRuleInStore(id, updates);
    }
    return { data };
  }, [updateRuleInStore]);

  // 규칙 삭제 - API 호출 후 스토어 업데이트
  const deleteRule = useCallback(async (id) => {
    const { error } = await api.deleteRule(id);
    if (error) {
      console.error('[규칙] 삭제 실패:', error);
      return { error };
    }
    deleteRuleFromStore(id);
    return { data: { ok: true } };
  }, [deleteRuleFromStore]);

  // 규칙 토글 - API 호출 후 스토어 업데이트 (낙관적 업데이트)
  const toggleRule = useCallback(async (id) => {
    toggleRuleInStore(id); // 즉시 UI 반영
    const { error } = await api.toggleRule(id);
    if (error) {
      console.error('[규칙] 토글 실패, 롤백:', error);
      toggleRuleInStore(id); // 실패 시 롤백
    }
  }, [toggleRuleInStore]);

  return {
    rules,
    createRule,
    updateRule,
    deleteRule,
    toggleRule,
  };
}
