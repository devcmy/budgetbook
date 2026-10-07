"use strict";
window.LABEL_CATALOG = [{"id": "tab.home", "defaultName": "이번달", "group": "탭 이름", "maxLength": 12}, {"id": "tab.records", "defaultName": "기록", "group": "탭 이름", "maxLength": 12}, {"id": "tab.reserve", "defaultName": "예비비", "group": "탭 이름", "maxLength": 12}, {"id": "tab.balance", "defaultName": "잔액", "group": "탭 이름", "maxLength": 12}, {"id": "tab.settings", "defaultName": "설정", "group": "탭 이름", "maxLength": 12}, {"id": "monthly.summary", "defaultName": "이번달 결산", "group": "이번달 메뉴", "maxLength": 20}, {"id": "monthly.variable", "defaultName": "비고정비", "group": "이번달 메뉴", "maxLength": 20}, {"id": "monthly.fixed", "defaultName": "고정비", "group": "이번달 메뉴", "maxLength": 20}, {"id": "monthly.cash", "defaultName": "비예비비", "group": "이번달 메뉴", "maxLength": 20}, {"id": "monthly.reserveUse", "defaultName": "예비비 사용", "group": "이번달 메뉴", "maxLength": 20}, {"id": "monthly.calendar", "defaultName": "일별 지출", "group": "이번달 메뉴", "maxLength": 20}, {"id": "monthly.tetris", "defaultName": "테트리스 비교", "group": "이번달 메뉴", "maxLength": 20}, {"id": "monthly.remaining", "defaultName": "예상 남는 돈", "group": "이번달 메뉴", "maxLength": 20}, {"id": "monthly.income", "defaultName": "총수입", "group": "이번달 메뉴", "maxLength": 20}, {"id": "monthly.outflow", "defaultName": "지출·적립", "group": "이번달 메뉴", "maxLength": 20}, {"id": "records.list", "defaultName": "거래 내역", "group": "기록 메뉴", "maxLength": 20}, {"id": "records.calendar", "defaultName": "달력", "group": "기록 메뉴", "maxLength": 20}, {"id": "records.detail", "defaultName": "거래 상세", "group": "기록 메뉴", "maxLength": 20}, {"id": "entry.add", "defaultName": "기록 입력", "group": "기록 메뉴", "maxLength": 20}, {"id": "entry.edit", "defaultName": "기록 수정", "group": "기록 메뉴", "maxLength": 20}, {"id": "reserve.purposes", "defaultName": "목적별 예비비", "group": "예비비·잔액 메뉴", "maxLength": 20}, {"id": "reserve.remaining", "defaultName": "남은 자금", "group": "예비비·잔액 메뉴", "maxLength": 20}, {"id": "reserve.all", "defaultName": "전체 항목", "group": "예비비·잔액 메뉴", "maxLength": 20}, {"id": "balance.places", "defaultName": "돈의 자리", "group": "예비비·잔액 메뉴", "maxLength": 20}, {"id": "balance.living", "defaultName": "생활비", "group": "예비비·잔액 메뉴", "maxLength": 20}, {"id": "balance.purposes", "defaultName": "목적자금", "group": "예비비·잔액 메뉴", "maxLength": 20}, {"id": "balance.movement", "defaultName": "내부 이동", "group": "예비비·잔액 메뉴", "maxLength": 20}, {"id": "balance.recent", "defaultName": "최근 잔액 변동", "group": "예비비·잔액 메뉴", "maxLength": 20}, {"id": "common.month", "defaultName": "월 선택", "group": "공통 메뉴", "maxLength": 20}, {"id": "common.about", "defaultName": "시안 안내", "group": "공통 메뉴", "maxLength": 20}];
const SETTINGS_KEY = 'tetris-budget-ui-v2';
function validAlias(value, limit) {
  if (typeof value !== 'string') return null;
  const text = value.trim();
  return text && [...text].length <= limit && !/[\u0000-\u001f\u007f]/.test(text) ? text : null;
}
function loadPreferences() {
  const clean = {names:{}, compact:true, reduceTransparency:false};
  try {
    const stored = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
    if (typeof stored.compact === 'boolean') clean.compact = stored.compact;
    if (typeof stored.reduceTransparency === 'boolean') clean.reduceTransparency = stored.reduceTransparency;
    for (const item of window.LABEL_CATALOG) {
      const name = validAlias(stored.names?.[item.id], item.maxLength);
      if(name) clean.names[item.id] = name;
    }
  } catch (_) {}
  return clean;
}
const preferences = loadPreferences();
function uiName(id) {return preferences.names[id] || window.LABEL_CATALOG.find(x=>x.id===id)?.defaultName || id;}
function persistPreferences() {
  try {localStorage.setItem(SETTINGS_KEY, JSON.stringify(preferences)); return true;} catch (_) {return false;}
}
function applyAppearance() {
  document.body.classList.toggle('comfortable', !preferences.compact);
  document.body.classList.toggle('reduce-transparency', preferences.reduceTransparency);
}
applyAppearance();
