/**
 * 스마트 연차계산기 & 로컬 엑셀 DB 관리 애플리케이션
 * - 입사년도(입사일) 기준 (디폴트) / 회계년도(1.1) 기준 선택 지원
 * - 오늘 기준 근속연차(N년차), 근속기간(N개월차), 근속일수(N일) 자동 계산
 * - '이번 주기/이전 주기' 대신 직관적인 '1년차(N년차)' / '이전 연차 전체' 용어 적용
 * - 우측 상단 엑셀 DB 버튼 하단에 본인 PC 저장 경로 표시 지원
 */

const STORAGE_KEY = "ONSE_ANNUAL_LEAVE_CALCULATOR_V1";

// 기본 상태값
const state = {
  userName: "",
  hireDate: "", // YYYY-MM-DD
  calcBasis: "anniversary", // 'anniversary' (입사년도 기본값) | 'fiscal' (회계년도)
  firstYearMode: "accrued", // 'accrued' (매월 1개씩 발생) | 'max11' (1년차 총 11개 한도로 표시)
  fiscalIncludeFirstYearMonthly: true,
  manualAdjust: 0,
  autoDeductAdvance: true, // 이전 연차 초과 先사용분(당겨쓰기) 차기 연차에서 자동 차감
  autoDownloadExcel: false,
  historyFilter: "all", // 'all' | 'current' | 'past'
  lastExcelFileName: "",
  lastExcelActionTime: "",
  records: [] // [{ id, date, type, days, note }]
};

// DOM Elements
const els = {
  referenceDateInput: document.getElementById("referenceDateInput"),
  resetTodayBtn: document.getElementById("resetTodayBtn"),
  autoDownloadExcelChk: document.getElementById("autoDownloadExcelChk"),
  autoDeductAdvanceChk: document.getElementById("autoDeductAdvanceChk"),
  excelFileInput: document.getElementById("excelFileInput"),
  importExcelBtn: document.getElementById("importExcelBtn"),
  exportExcelBtn: document.getElementById("exportExcelBtn"),
  exportExcelBottomBtn: document.getElementById("exportExcelBottomBtn"),
  savedPathText: document.getElementById("savedPathText"),
  savedPathStatus: document.getElementById("savedPathStatus"),
  lastSavedIndicator: document.getElementById("lastSavedIndicator"),

  userNameInput: document.getElementById("userNameInput"),
  hireDateInput: document.getElementById("hireDateInput"),
  basisRadios: document.querySelectorAll('input[name="calcBasis"]'),
  basisOptionAnniversary: document.getElementById("basisOptionAnniversary"),
  basisOptionFiscal: document.getElementById("basisOptionFiscal"),
  currentBasisBadge: document.getElementById("currentBasisBadge"),

  firstYearOptionRow: document.getElementById("firstYearOptionRow"),
  firstYearModeSelect: document.getElementById("firstYearModeSelect"),
  fiscalSecondYearRow: document.getElementById("fiscalSecondYearRow"),
  fiscalIncludeFirstYearMonthlyChk: document.getElementById("fiscalIncludeFirstYearMonthlyChk"),
  manualAdjustInput: document.getElementById("manualAdjustInput"),

  currentCyclePeriodText: document.getElementById("currentCyclePeriodText"),
  nextResetDdayBadge: document.getElementById("nextResetDdayBadge"),

  tenureYearCount: document.getElementById("tenureYearCount"),
  tenureExactYears: document.getElementById("tenureExactYears"),
  tenureMonthCount: document.getElementById("tenureMonthCount"),
  tenureExactMonths: document.getElementById("tenureExactMonths"),
  tenureDayCount: document.getElementById("tenureDayCount"),
  tenureExactDays: document.getElementById("tenureExactDays"),

  generatedLeaveLabel: document.getElementById("generatedLeaveLabel"),
  generatedLeaveVal: document.getElementById("generatedLeaveVal"),
  generatedLeaveDesc: document.getElementById("generatedLeaveDesc"),
  usedLeaveLabel: document.getElementById("usedLeaveLabel"),
  usedLeaveVal: document.getElementById("usedLeaveVal"),
  usedLeaveDesc: document.getElementById("usedLeaveDesc"),
  remainingLeaveLabel: document.getElementById("remainingLeaveLabel"),
  remainingLeaveVal: document.getElementById("remainingLeaveVal"),
  remainingLeaveDesc: document.getElementById("remainingLeaveDesc"),
  remainingStatCard: document.querySelector(".stat-remaining"),

  usageRateLabel: document.getElementById("usageRateLabel"),
  usageRateText: document.getElementById("usageRateText"),
  usageProgressBar: document.getElementById("usageProgressBar"),

  leaveRecordForm: document.getElementById("leaveRecordForm"),
  entryRowsContainer: document.getElementById("entryRowsContainer"),
  addInputRowTopBtn: document.getElementById("addInputRowTopBtn"),
  addInputRowBtn: document.getElementById("addInputRowBtn"),

  historySubtitleText: document.getElementById("historySubtitleText"),
  filterTabs: document.querySelectorAll(".filter-tab"),
  currentTabLabel: document.getElementById("currentTabLabel"),
  countAllBadge: document.getElementById("countAllBadge"),
  countCurrentBadge: document.getElementById("countCurrentBadge"),
  countPastBadge: document.getElementById("countPastBadge"),

  stripCurrentLabel: document.getElementById("stripCurrentLabel"),
  stripCurrentUsed: document.getElementById("stripCurrentUsed"),
  stripPastUsed: document.getElementById("stripPastUsed"),
  stripAdvanceUsed: document.getElementById("stripAdvanceUsed"),
  stripTotalUsed: document.getElementById("stripTotalUsed"),

  leaveHistoryTbody: document.getElementById("leaveHistoryTbody"),
  emptyHistoryState: document.getElementById("emptyHistoryState"),
  clearAllRecordsBtn: document.getElementById("clearAllRecordsBtn"),
  lawReferenceTbody: document.getElementById("lawReferenceTbody"),
  toastContainer: document.getElementById("toastContainer")
};

// ============================================================================
// 1. 날짜 및 숫자 유틸리티
// ============================================================================

function getTodayStr() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function parseDate(dateStr) {
  if (!dateStr) return null;
  const parts = String(dateStr).trim().split(/[-./]/);
  if (parts.length !== 3) return null;
  const y = parseInt(parts[0], 10);
  const m = parseInt(parts[1], 10) - 1;
  const d = parseInt(parts[2], 10);
  if (isNaN(y) || isNaN(m) || isNaN(d)) return null;
  return new Date(y, m, d, 0, 0, 0, 0);
}

function formatDate(dateObj) {
  if (!dateObj || isNaN(dateObj.getTime())) return "";
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, "0");
  const d = String(dateObj.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function formatDateDot(dateObj) {
  if (!dateObj || isNaN(dateObj.getTime())) return "-";
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, "0");
  const d = String(dateObj.getDate()).padStart(2, "0");
  return `${y}.${m}.${d}`;
}

function addDays(dateObj, days) {
  const res = new Date(dateObj.getTime());
  res.setDate(res.getDate() + days);
  return res;
}

function diffDaysBetween(startObj, endObj) {
  const msPerDay = 24 * 60 * 60 * 1000;
  const utc1 = Date.UTC(startObj.getFullYear(), startObj.getMonth(), startObj.getDate());
  const utc2 = Date.UTC(endObj.getFullYear(), endObj.getMonth(), endObj.getDate());
  return Math.floor((utc2 - utc1) / msPerDay);
}

function getSafeAnniversary(year, monthIndex, day) {
  const lastDayOfMonth = new Date(year, monthIndex + 1, 0).getDate();
  const safeDay = Math.min(day, lastDayOfMonth);
  return new Date(year, monthIndex, safeDay, 0, 0, 0, 0);
}

function formatNumber(num) {
  if (num === null || num === undefined || isNaN(num)) return "0";
  const rounded = Math.round(num * 100) / 100;
  return rounded.toLocaleString("ko-KR", { maximumFractionDigits: 2 });
}

// ============================================================================
// 2. 근속 기간 및 연차 산정 엔진 (입사년도 기준 / 회계년도 기준)
// ============================================================================

/**
 * 입사일(hireDate)로부터 기준일(refDate)까지의 근속연차, 근속개월차, 근속일수를 계산합니다.
 */
function calculateTenure(hireDate, refDate) {
  if (!hireDate || !refDate || refDate < hireDate) {
    return null;
  }

  const elapsedDays = diffDaysBetween(hireDate, refDate);
  const totalDaysInclusive = elapsedDays + 1;

  let years = refDate.getFullYear() - hireDate.getFullYear();
  let months = refDate.getMonth() - hireDate.getMonth();
  let days = refDate.getDate() - hireDate.getDate();

  if (days < 0) {
    months -= 1;
    const prevMonthLastDay = new Date(refDate.getFullYear(), refDate.getMonth(), 0).getDate();
    const hireDayCapped = Math.min(hireDate.getDate(), prevMonthLastDay);
    days = prevMonthLastDay - hireDayCapped + refDate.getDate();
  }

  if (months < 0) {
    years -= 1;
    months += 12;
  }

  const exactFullMonths = years * 12 + months;
  const tenureYearOrdinal = years + 1; // 1년차, 2년차, 3년차...
  const tenureMonthOrdinal = exactFullMonths + 1; // 1개월차, 2개월차...

  return {
    exactYears: years,
    remMonths: months,
    remDays: days,
    exactFullMonths,
    elapsedDays,
    totalDaysInclusive,
    tenureYearOrdinal,
    tenureMonthOrdinal
  };
}

/**
 * 만 근속년수(fullYears, 1 이상)에 따른 연차 일수 계산
 * - 2년차 (만 1년 근속): 기본 15일
 * - 3년차부터 매년 1일씩 증가 (만 2년 16일, 만 3년 17일, ...)
 * - 최대 25일 상한(cap) 적용
 */
function getStatutoryLeaveForFullYears(fullYears) {
  if (fullYears < 1) return 0;
  // fullYears = 1 (2년차): 15일, fullYears = 2 (3년차): 16일 ...
  const extraDays = fullYears - 1;
  return Math.min(25, 15 + extraDays);
}

/**
 * 현재 연차 범위의 시작일·종료일 및 발생 연차를 계산합니다.
 */
function calculateCurrentCycleAndLeave(hireDate, refDate, options) {
  const {
    calcBasis,
    firstYearMode,
    fiscalIncludeFirstYearMonthly,
    manualAdjust
  } = options;

  const tenure = calculateTenure(hireDate, refDate);
  if (!tenure) return null;

  const hireY = hireDate.getFullYear();
  const hireM = hireDate.getMonth();
  const hireD = hireDate.getDate();
  const refY = refDate.getFullYear();

  let cycleStart = null;
  let cycleEnd = null;
  let baseLeave = 0;
  let formulaDesc = "";
  let cycleLabel = `${tenure.tenureYearOrdinal}년차`;

  if (calcBasis === "anniversary") {
    // 1) 입사년도(입사일) 기준 (디폴트)
    const annivThisYear = getSafeAnniversary(refY, hireM, hireD);
    let startYear = refY;
    if (refDate < annivThisYear) {
      startYear = refY - 1;
    }

    cycleStart = getSafeAnniversary(startYear, hireM, hireD);
    const nextAnniv = getSafeAnniversary(startYear + 1, hireM, hireD);
    cycleEnd = addDays(nextAnniv, -1);

    const fullYearsAtCycleStart = startYear - hireY;
    const yearOrdinal = fullYearsAtCycleStart + 1;
    cycleLabel = `${yearOrdinal}년차`;

    if (fullYearsAtCycleStart === 0) {
      // 1년차 (만 1년 미만): 1개월 개근 시 1일씩 발생 (최대 11일)
      const accruedMonthly = Math.min(11, tenure.exactFullMonths);
      if (firstYearMode === "max11") {
        baseLeave = 11;
        formulaDesc = `1년차 총 발생 한도 11일 기준 (현재 만 ${tenure.exactFullMonths}개월 근속)`;
      } else {
        baseLeave = accruedMonthly;
        formulaDesc = `1년차 매월 1일씩 발생: 만 ${tenure.exactFullMonths}개월 개근 (${accruedMonthly}일 발생 / 1년차 최대 11일)`;
      }
    } else {
      // 2년차 이상 (만 1년 이상): 기본 15일 + 만 3년 이상부터 2년마다 1일 가산 (최대 25일)
      baseLeave = getStatutoryLeaveForFullYears(fullYearsAtCycleStart);
      const extra = baseLeave - 15;
      formulaDesc =
        extra > 0
          ? `${yearOrdinal}년차 (만 ${fullYearsAtCycleStart}년 근속): 기본 15일 + 가산 ${extra}일`
          : `${yearOrdinal}년차 (만 ${fullYearsAtCycleStart}년 근속): 법정 기본 15일`;
    }
  } else {
    // 2) 회계년도(1월 1일) 기준
    const fiscalOrdinal = refY - hireY + 1;
    cycleLabel = `${tenure.tenureYearOrdinal}년차 (${refY}년 회계기준)`;

    if (refY === hireY) {
      // 입사 당해 연도 (입사일 ~ 12/31)
      cycleStart = new Date(hireDate.getTime());
      cycleEnd = new Date(refY, 11, 31, 0, 0, 0, 0);

      const maxMonthsInFirstFiscalYear = 11 - hireM;
      const accruedMonthly = Math.min(11, tenure.exactFullMonths);

      if (firstYearMode === "max11") {
        baseLeave = Math.max(0, maxMonthsInFirstFiscalYear);
        formulaDesc = `1년차 (${refY}년 연말까지 발생하는 월차 총 ${baseLeave}일 기준)`;
      } else {
        baseLeave = accruedMonthly;
        formulaDesc = `1년차 매월 1일씩 발생: 만 ${tenure.exactFullMonths}개월 개근분`;
      }
    } else if (refY === hireY + 1) {
      // 입사 2차년도 (다음 해 1월 1일 ~ 12월 31일)
      cycleStart = new Date(refY, 0, 1, 0, 0, 0, 0);
      cycleEnd = new Date(refY, 11, 31, 0, 0, 0, 0);

      if (hireM === 0 && hireD === 1) {
        baseLeave = 15;
        formulaDesc = `2년차 (1.1 입사 만 1년): 기본 15일`;
      } else {
        const startOfHireYear = new Date(hireY, 0, 1);
        const endOfHireYear = new Date(hireY, 11, 31);
        const totalDaysInHireYear = diffDaysBetween(startOfHireYear, endOfHireYear) + 1;
        const workedDaysInHireYear = diffDaysBetween(hireDate, endOfHireYear) + 1;
        const proportionalLeave = Math.round(15 * (workedDaysInHireYear / totalDaysInHireYear) * 10) / 10;

        let secondYearMonthly = 0;
        if (fiscalIncludeFirstYearMonthly) {
          const totalPossibleMonthly = 11;
          const firstYearMonths = Math.max(0, 11 - hireM);
          const remainingMonthlyCap = Math.max(0, totalPossibleMonthly - firstYearMonths);

          if (firstYearMode === "max11") {
            secondYearMonthly = remainingMonthlyCap;
          } else {
            const totalAccruedSoFar = Math.min(11, tenure.exactFullMonths);
            secondYearMonthly = Math.max(0, totalAccruedSoFar - firstYearMonths);
          }
        }

        baseLeave = Math.round((proportionalLeave + secondYearMonthly) * 100) / 100;
        formulaDesc =
          secondYearMonthly > 0
            ? `회계 ${fiscalOrdinal}년차 비례연차 ${proportionalLeave}일 + 1년 미만 월차 ${secondYearMonthly}일`
            : `회계 ${fiscalOrdinal}년차 비례연차 (${workedDaysInHireYear}/${totalDaysInHireYear}일 비례)`;
      }
    } else {
      // 입사 3차년도 이후 (refY - hireY >= 2)
      cycleStart = new Date(refY, 0, 1, 0, 0, 0, 0);
      cycleEnd = new Date(refY, 11, 31, 0, 0, 0, 0);

      const isJanFirstHire = hireM === 0 && hireD === 1;
      const fullYearsAtJan1 = isJanFirstHire ? refY - hireY : refY - hireY - 1;
      baseLeave = getStatutoryLeaveForFullYears(fullYearsAtJan1);
      const extra = baseLeave - 15;
      formulaDesc =
        extra > 0
          ? `${tenure.tenureYearOrdinal}년차 (${refY}년 회계기준): 기본 15일 + 가산 ${extra}일`
          : `${tenure.tenureYearOrdinal}년차 (${refY}년 회계기준): 법정 기본 15일`;
    }
  }

  const adj = parseFloat(manualAdjust) || 0;
  const totalGenerated = Math.max(0, Math.round((baseLeave + adj) * 100) / 100);
  if (adj !== 0) {
    const sign = adj > 0 ? `+${adj}` : `${adj}`;
    formulaDesc += ` (수동보정 ${sign}일 반영)`;
  }

  const daysUntilReset = diffDaysBetween(refDate, cycleEnd) + 1;

  return {
    tenure,
    cycleStart,
    cycleEnd,
    cycleLabel,
    shortYearLabel: `${tenure.tenureYearOrdinal}년차`,
    baseLeave,
    manualAdjust: adj,
    totalGenerated,
    formulaDesc,
    daysUntilReset
  };
}

/**
 * 특정 차수(1년차, 2년차, ...)의 유효 기간과 법정 발생 한도(cap)를 계산합니다.
 */
function getCycleBoundsAndCap(ordinal, hireDate, calcBasis, options = {}) {
  if (!hireDate || ordinal < 1) return null;
  const hireY = hireDate.getFullYear();
  const hireM = hireDate.getMonth();
  const hireD = hireDate.getDate();

  if (calcBasis === "anniversary") {
    const cycleStart = getSafeAnniversary(hireY + ordinal - 1, hireM, hireD);
    const nextAnniv = getSafeAnniversary(hireY + ordinal, hireM, hireD);
    const cycleEnd = addDays(nextAnniv, -1);

    let cap = 0;
    if (ordinal === 1) {
      cap = 11; // 1년차 법정 최대 한도
    } else {
      cap = getStatutoryLeaveForFullYears(ordinal - 1);
    }

    return { ordinal, cycleStart, cycleEnd, cap };
  } else {
    // 회계년도 기준
    const cycleYear = hireY + ordinal - 1;
    let cycleStart, cycleEnd, cap;

    if (ordinal === 1) {
      cycleStart = new Date(hireDate.getTime());
      cycleEnd = new Date(cycleYear, 11, 31, 0, 0, 0, 0);
      cap = Math.max(0, 11 - hireM);
    } else if (ordinal === 2) {
      cycleStart = new Date(cycleYear, 0, 1, 0, 0, 0, 0);
      cycleEnd = new Date(cycleYear, 11, 31, 0, 0, 0, 0);
      if (hireM === 0 && hireD === 1) {
        cap = 15;
      } else {
        const startOfHireYear = new Date(hireY, 0, 1);
        const endOfHireYear = new Date(hireY, 11, 31);
        const totalDaysInHireYear = diffDaysBetween(startOfHireYear, endOfHireYear) + 1;
        const workedDaysInHireYear = diffDaysBetween(hireDate, endOfHireYear) + 1;
        const proportionalLeave = Math.round(15 * (workedDaysInHireYear / totalDaysInHireYear) * 10) / 10;
        const firstYearMonths = Math.max(0, 11 - hireM);
        const secondYearMonthly = Math.max(0, 11 - firstYearMonths);
        cap = Math.round((proportionalLeave + secondYearMonthly) * 100) / 100;
      }
    } else {
      cycleStart = new Date(cycleYear, 0, 1, 0, 0, 0, 0);
      cycleEnd = new Date(cycleYear, 11, 31, 0, 0, 0, 0);
      const isJanFirstHire = hireM === 0 && hireD === 1;
      const fullYears = isJanFirstHire ? ordinal - 1 : ordinal - 2;
      cap = getStatutoryLeaveForFullYears(fullYears);
    }

    return { ordinal, cycleStart, cycleEnd, cap };
  }
}

/**
 * 특정 사용일자가 입사일 기준 몇 년차에 속하는지 계산합니다.
 */
function getRecordCycleOrdinal(recDate, hireDate, calcBasis) {
  if (!recDate || !hireDate) return 1;
  const hireY = hireDate.getFullYear();
  const hireM = hireDate.getMonth();
  const hireD = hireDate.getDate();
  const recY = recDate.getFullYear();

  if (calcBasis === "anniversary") {
    const annivInRecYear = getSafeAnniversary(recY, hireM, hireD);
    const cycleStartYear = recDate < annivInRecYear ? recY - 1 : recY;
    const ord = cycleStartYear - hireY + 1;
    return Math.max(1, ord);
  } else {
    const ord = recY - hireY + 1;
    return Math.max(1, ord);
  }
}

// ============================================================================
// 3. 화면 렌더링 및 로컬 스토리지 동기화
// ============================================================================

function saveToLocalStorage() {
  try {
    const payload = {
      userName: state.userName,
      hireDate: state.hireDate,
      calcBasis: state.calcBasis,
      firstYearMode: state.firstYearMode,
      fiscalIncludeFirstYearMonthly: state.fiscalIncludeFirstYearMonthly,
      manualAdjust: state.manualAdjust,
      autoDeductAdvance: state.autoDeductAdvance,
      autoDownloadExcel: state.autoDownloadExcel,
      lastExcelFileName: state.lastExcelFileName,
      lastExcelActionTime: state.lastExcelActionTime,
      records: state.records,
      updatedAt: new Date().toISOString()
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    const now = new Date();
    const hh = String(now.getHours()).padStart(2, "0");
    const mm = String(now.getMinutes()).padStart(2, "0");
    const ss = String(now.getSeconds()).padStart(2, "0");
    els.lastSavedIndicator.textContent = `자동 저장됨 (${hh}:${mm}:${ss})`;
  } catch (err) {
    console.error("로컬 스토리지 저장 실패:", err);
  }
}

function loadFromLocalStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return false;
    const data = JSON.parse(raw);
    if (data && typeof data === "object") {
      state.userName = data.userName || "";
      state.hireDate = data.hireDate || "";
      state.calcBasis = data.calcBasis === "fiscal" ? "fiscal" : "anniversary";
      state.firstYearMode = data.firstYearMode === "max11" ? "max11" : "accrued";
      state.fiscalIncludeFirstYearMonthly =
        data.fiscalIncludeFirstYearMonthly !== undefined ? Boolean(data.fiscalIncludeFirstYearMonthly) : true;
      state.manualAdjust = parseFloat(data.manualAdjust) || 0;
      state.autoDeductAdvance =
        data.autoDeductAdvance !== undefined ? Boolean(data.autoDeductAdvance) : true;
      state.autoDownloadExcel = Boolean(data.autoDownloadExcel);
      state.lastExcelFileName = data.lastExcelFileName || "";
      state.lastExcelActionTime = data.lastExcelActionTime || "";
      state.records = Array.isArray(data.records) ? data.records : [];
      return true;
    }
  } catch (err) {
    console.error("로컬 스토리지 로드 실패:", err);
  }
  return false;
}

function renderSavedPathBar() {
  if (state.lastExcelFileName) {
    els.savedPathText.textContent = `C:\\Users\\[본인PC계정]\\Downloads\\${state.lastExcelFileName}`;
    els.savedPathStatus.textContent = state.lastExcelActionTime
      ? `(${state.lastExcelActionTime} 기준)`
      : "(내 PC 다운로드 폴더 보관)";
  } else {
    els.savedPathText.textContent = `C:\\Users\\[본인PC계정]\\Downloads\\연차 사용이력 기록용.xlsx`;
    els.savedPathStatus.textContent = "(현재 브라우저 내부 자동저장 중)";
  }
}

function syncControlsFromState() {
  els.userNameInput.value = state.userName;
  els.hireDateInput.value = state.hireDate;
  els.firstYearModeSelect.value = state.firstYearMode;
  els.fiscalIncludeFirstYearMonthlyChk.checked = state.fiscalIncludeFirstYearMonthly;
  els.manualAdjustInput.value = state.manualAdjust;
  if (els.autoDeductAdvanceChk) {
    els.autoDeductAdvanceChk.checked = state.autoDeductAdvance;
  }
  els.autoDownloadExcelChk.checked = state.autoDownloadExcel;

  els.basisRadios.forEach((radio) => {
    radio.checked = radio.value === state.calcBasis;
  });

  if (state.calcBasis === "anniversary") {
    els.basisOptionAnniversary.classList.add("selected");
    els.basisOptionFiscal.classList.remove("selected");
    els.currentBasisBadge.textContent = "입사년도 기준 (기본)";
    els.fiscalSecondYearRow.style.display = "none";
  } else {
    els.basisOptionAnniversary.classList.remove("selected");
    els.basisOptionFiscal.classList.add("selected");
    els.currentBasisBadge.textContent = "회계년도(1.1) 기준";
    els.fiscalSecondYearRow.style.display = "flex";
  }

  renderSavedPathBar();
}

function renderAll() {
  syncControlsFromState();

  const hireDateObj = parseDate(state.hireDate);
  const refDateObj = parseDate(els.referenceDateInput.value) || parseDate(getTodayStr());

  // 입사일이 없거나 기준일이 입사일보다 빠른 경우 초기화 화면 표시
  if (!hireDateObj || refDateObj < hireDateObj) {
    els.currentCyclePeriodText.textContent = !hireDateObj
      ? "좌측 상단에서 입사일을 먼저 선택해 주세요."
      : "기준일(오늘)이 입사일보다 빠릅니다. 날짜를 확인해 주세요.";
    els.nextResetDdayBadge.textContent = "대기 중";

    els.tenureYearCount.textContent = "-년차";
    els.tenureExactYears.textContent = "만 -년 -개월";
    els.tenureMonthCount.textContent = "-개월차";
    els.tenureExactMonths.textContent = "만 -개월 경과";
    els.tenureDayCount.textContent = "-일";
    els.tenureExactDays.textContent = "입사일 포함 재직일수";

    els.generatedLeaveLabel.textContent = "현재 연차 발생일수";
    els.generatedLeaveVal.textContent = "0";
    els.generatedLeaveDesc.textContent = "입사일을 입력하면 자동 계산됩니다";

    els.usedLeaveLabel.textContent = "현재 연차 사용일수";
    els.usedLeaveVal.textContent = "0";
    els.usedLeaveDesc.textContent = "현재 연차 사용 0건";

    els.remainingLeaveVal.textContent = "0";
    els.remainingLeaveDesc.textContent = "사용 가능 연차";
    els.usageRateLabel.textContent = "현재 연차 소진율";
    els.usageRateText.textContent = "0.0%";
    els.usageProgressBar.style.width = "0%";

    els.currentTabLabel.textContent = "1년차 (현재)";
    els.stripCurrentLabel.textContent = "현재 연차 사용 합계";
    if (els.stripAdvanceUsed) els.stripAdvanceUsed.textContent = "0일 (0건)";
    els.historySubtitleText.innerHTML =
      "전체 기록은 연차가 리셋되어도 계속 이어지며, <strong>현재 연차</strong> 내역 및 <strong>이전 연차 초과분(끌어쓰기) 차감</strong>이 잔여 연차 계산에 반영됩니다.";

    renderHistoryTable(null);
    renderLawReferenceTable(null);
    return;
  }

  const calcResult = calculateCurrentCycleAndLeave(hireDateObj, refDateObj, {
    calcBasis: state.calcBasis,
    firstYearMode: state.firstYearMode,
    fiscalIncludeFirstYearMonthly: state.fiscalIncludeFirstYearMonthly,
    manualAdjust: state.manualAdjust
  });

  const {
    tenure,
    cycleStart,
    cycleEnd,
    cycleLabel,
    shortYearLabel,
    totalGenerated,
    formulaDesc,
    daysUntilReset
  } = calcResult;

  // 1) 상단 현재 적용 연차 및 리셋 D-Day 표시
  els.currentCyclePeriodText.textContent = `현재 적용 연차: ${cycleLabel} (${formatDateDot(cycleStart)} – ${formatDateDot(cycleEnd)})`;
  els.nextResetDdayBadge.textContent = `다음 연차 갱신까지 ${daysUntilReset.toLocaleString("ko-KR")}일`;

  // 2) 근속연차 / 근속개월차 / 근속일수 표시
  els.tenureYearCount.textContent = `${tenure.tenureYearOrdinal.toLocaleString("ko-KR")}년차`;
  els.tenureExactYears.textContent = `만 ${tenure.exactYears}년 ${tenure.remMonths}개월 (${tenure.remDays}일)`;

  els.tenureMonthCount.textContent = `${tenure.tenureMonthOrdinal.toLocaleString("ko-KR")}개월차`;
  els.tenureExactMonths.textContent = `만 ${tenure.exactFullMonths.toLocaleString("ko-KR")}개월 경과`;

  els.tenureDayCount.textContent = `${tenure.totalDaysInclusive.toLocaleString("ko-KR")}일`;
  els.tenureExactDays.textContent = `입사일 포함 ${tenure.totalDaysInclusive.toLocaleString("ko-KR")}일째 (만 ${tenure.elapsedDays.toLocaleString("ko-KR")}일)`;

  // 3) 연차 사용 기록 분류 및 차수별 先사용(당겨쓰기) 분석
  let currentUsedDays = 0;
  let currentUsedCount = 0;
  let pastUsedDays = 0;
  let pastUsedCount = 0;
  let totalUsedDays = 0;

  const sortedRecords = [...state.records].sort((a, b) => a.date.localeCompare(b.date));

  // 차수별 누적 사용일수 맵
  const cycleAccumMap = {};

  const enrichedRecords = sortedRecords.map((rec) => {
    const daysNum = parseFloat(rec.days) || 0;
    const recDate = parseDate(rec.date);
    totalUsedDays += daysNum;

    if (!recDate) {
      return {
        ...rec,
        daysNum,
        cycleStatus: "unknown",
        cycleBadgeLabel: "미분류",
        isAdvance: false,
        advancePortion: 0
      };
    }

    const recCycleOrd = getRecordCycleOrdinal(recDate, hireDateObj, state.calcBasis);
    const cycleDef = getCycleBoundsAndCap(recCycleOrd, hireDateObj, state.calcBasis, {
      firstYearMode: state.firstYearMode,
      fiscalIncludeFirstYearMonthly: state.fiscalIncludeFirstYearMonthly
    });
    const cap = cycleDef ? cycleDef.cap : 999;

    const prevAcc = cycleAccumMap[recCycleOrd] || 0;
    const newAcc = Math.round((prevAcc + daysNum) * 100) / 100;
    cycleAccumMap[recCycleOrd] = newAcc;

    let isAdvance = false;
    let advancePortion = 0;
    const targetNextOrd = recCycleOrd + 1;

    if (prevAcc >= cap) {
      // 이미 발생 한도를 초과한 상태에서의 추가 사용 -> 해당 건 전체가 先사용
      isAdvance = true;
      advancePortion = daysNum;
    } else if (newAcc > cap) {
      // 이번 사용으로 발생 한도를 돌파함 -> 초과분만큼 先사용
      isAdvance = true;
      advancePortion = Math.round((newAcc - cap) * 100) / 100;
    }

    // 현재 차수 vs 이전 차수 vs 미래 차수 상태 판별
    let status = "past";
    if (recCycleOrd === tenure.tenureYearOrdinal) {
      status = "current";
      currentUsedDays += daysNum;
      currentUsedCount += 1;
    } else if (recCycleOrd > tenure.tenureYearOrdinal) {
      status = "future";
    } else {
      pastUsedDays += daysNum;
      pastUsedCount += 1;
    }

    // 배지 라벨 구성 (ONSE 사내 규정: 모든 차수 N년차(N+1년차 끌) 통일)
    let badgeLabel = "";
    if (isAdvance) {
      badgeLabel = `${recCycleOrd}년차(${targetNextOrd}년차 끌)`;
    } else {
      if (status === "current") {
        badgeLabel = `${tenure.tenureYearOrdinal}년차 (현재)`;
      } else if (status === "future") {
        badgeLabel = `${recCycleOrd}년차 (예정)`;
      } else {
        badgeLabel =
          state.calcBasis === "fiscal"
            ? `이전 (${recCycleOrd}년차·${recDate.getFullYear()}년)`
            : `이전 (${recCycleOrd}년차)`;
      }
    }

    return {
      ...rec,
      daysNum,
      recCycleOrd,
      cycleStatus: status,
      isAdvance,
      advancePortion,
      targetNextOrd,
      cycleBadgeLabel: badgeLabel
    };
  });

  currentUsedDays = Math.round(currentUsedDays * 100) / 100;
  pastUsedDays = Math.round(pastUsedDays * 100) / 100;
  totalUsedDays = Math.round(totalUsedDays * 100) / 100;

  // 전체 등록된 기록 중 발생한 총 先사용 통계 (요약 스트립용)
  let totalAdvanceDays = 0;
  let totalAdvanceCount = 0;
  enrichedRecords.forEach((r) => {
    if (r.isAdvance) {
      totalAdvanceDays += r.advancePortion;
      totalAdvanceCount += 1;
    }
  });
  totalAdvanceDays = Math.round(totalAdvanceDays * 100) / 100;

  // 이전 연차 초과분(끌어쓰기) 순차 차감 계산 (ONSE 사내 규정 반영)
  // 1) 1년차에 11일을 초과하여 사용한 일수도 2년차 연차 당겨쓰기로 처리되어 2년차 발생 연차에서 차감됨.
  // 2) 2년차 이후 발생한 초과분 또한 차기 연차 발생 일수에서 순차적으로 차감됨.
  let totalPastAdvanceDeficit = 0;
  let carryoverDeficit = 0;
  for (let ord = 1; ord <= tenure.tenureYearOrdinal; ord++) {
    if (ord === tenure.tenureYearOrdinal) {
      totalPastAdvanceDeficit = carryoverDeficit;
      break;
    }
    const cycleDef = getCycleBoundsAndCap(ord, hireDateObj, state.calcBasis, {
      firstYearMode: state.firstYearMode,
      fiscalIncludeFirstYearMonthly: state.fiscalIncludeFirstYearMonthly
    });
    const baseCap = cycleDef ? cycleDef.cap : (ord === 1 ? 11 : 15);
    const availableInCycle = Math.max(0, baseCap - carryoverDeficit);
    const usedInCycle = cycleAccumMap[ord] || 0;
    if (usedInCycle > availableInCycle) {
      carryoverDeficit = Math.round((usedInCycle - availableInCycle) * 100) / 100;
    } else {
      carryoverDeficit = 0;
    }
  }
  totalPastAdvanceDeficit = Math.round(totalPastAdvanceDeficit * 100) / 100;

  // 이전 연차 초과분 자동 차감 적용
  const advanceDeductedDays = state.autoDeductAdvance ? totalPastAdvanceDeficit : 0;
  const netGenerated = Math.max(
    0,
    Math.round((calcResult.baseLeave - advanceDeductedDays + calcResult.manualAdjust) * 100) / 100
  );

  const remainingDays = Math.round(
    (calcResult.baseLeave - advanceDeductedDays + calcResult.manualAdjust - currentUsedDays) * 100
  ) / 100;

  // 4) 직관적인 N년차 용어로 좌측/우측 라벨 동기화
  els.generatedLeaveLabel.textContent = `${shortYearLabel} 발생 연차`;
  els.generatedLeaveVal.textContent = formatNumber(netGenerated);

  let updatedFormulaDesc = calcResult.formulaDesc;
  if (advanceDeductedDays > 0) {
    const pastCycleName = tenure.tenureYearOrdinal > 1 ? `${tenure.tenureYearOrdinal - 1}년차` : "이전 연차";
    updatedFormulaDesc = `기본 ${formatNumber(calcResult.baseLeave)}일 – ${pastCycleName} 초과(끌어쓰기) ${formatNumber(advanceDeductedDays)}일 차감`;
    if (calcResult.manualAdjust !== 0) {
      const sign = calcResult.manualAdjust > 0 ? `+${calcResult.manualAdjust}` : `${calcResult.manualAdjust}`;
      updatedFormulaDesc += ` (수동보정 ${sign}일)`;
    }
  }
  els.generatedLeaveDesc.textContent = updatedFormulaDesc;

  els.usedLeaveLabel.textContent = `${shortYearLabel} 사용 연차`;
  els.usedLeaveVal.textContent = formatNumber(currentUsedDays);
  els.usedLeaveDesc.textContent = `${shortYearLabel} 사용 ${currentUsedCount}건 합산`;

  els.remainingLeaveVal.textContent = formatNumber(remainingDays);
  if (remainingDays < 0) {
    els.remainingStatCard.classList.add("negative-remaining");
    els.remainingLeaveDesc.textContent = `가용 연차 대비 ${formatNumber(Math.abs(remainingDays))}일 초과 사용`;
  } else {
    els.remainingStatCard.classList.remove("negative-remaining");
    els.remainingLeaveDesc.textContent = `${formatDateDot(cycleEnd)}까지 사용 가능`;
  }

  // 소진율 프로그레스 바
  els.usageRateLabel.textContent = `${shortYearLabel} 연차 소진율`;
  let usageRate = 0;
  if (netGenerated > 0) {
    usageRate = (currentUsedDays / netGenerated) * 100;
  } else if (currentUsedDays > 0) {
    usageRate = 100;
  }
  els.usageRateText.textContent = `${usageRate.toFixed(1)}% (${formatNumber(currentUsedDays)} / ${formatNumber(netGenerated)}일)`;
  els.usageProgressBar.style.width = `${Math.min(100, Math.max(0, usageRate))}%`;
  if (usageRate > 100) {
    els.usageProgressBar.classList.add("over-limit");
  } else {
    els.usageProgressBar.classList.remove("over-limit");
  }

  // 5) 우측 연차 사용 기록 탭 및 요약 스트립 업데이트
  els.currentTabLabel.textContent = `${shortYearLabel} (현재)`;
  els.stripCurrentLabel.textContent = `${shortYearLabel} 사용 합계`;
  els.historySubtitleText.innerHTML = `전체 기록은 연차가 리셋되어도 계속 이어지며, <strong>${shortYearLabel} (현재)</strong> 내역 및 <strong>이전 연차 초과분(끌어쓰기) 차감</strong>이 잔여 연차 계산에 반영됩니다.`;

  els.countAllBadge.textContent = enrichedRecords.length;
  els.countCurrentBadge.textContent = currentUsedCount;
  els.countPastBadge.textContent = pastUsedCount;

  els.stripCurrentUsed.textContent = `${formatNumber(currentUsedDays)}일 (${currentUsedCount}건)`;
  els.stripPastUsed.textContent = `${formatNumber(pastUsedDays)}일 (${pastUsedCount}건)`;
  if (els.stripAdvanceUsed) {
    els.stripAdvanceUsed.textContent = `${formatNumber(totalAdvanceDays)}일 (${totalAdvanceCount}건)`;
  }
  els.stripTotalUsed.textContent = `${formatNumber(totalUsedDays)}일 (${enrichedRecords.length}건)`;

  renderHistoryTable(enrichedRecords);
  renderLawReferenceTable(tenure.tenureYearOrdinal);
}

function renderHistoryTable(enrichedRecords) {
  els.leaveHistoryTbody.innerHTML = "";

  if (!enrichedRecords || enrichedRecords.length === 0) {
    els.emptyHistoryState.style.display = "block";
    return;
  }

  const filtered = enrichedRecords.filter((rec) => {
    if (state.historyFilter === "current") return rec.cycleStatus === "current";
    if (state.historyFilter === "past") return rec.cycleStatus === "past";
    return true;
  });

  if (filtered.length === 0) {
    els.emptyHistoryState.style.display = "block";
    return;
  }

  els.emptyHistoryState.style.display = "none";

  filtered.forEach((rec, idx) => {
    const tr = document.createElement("tr");
    if (rec.cycleStatus === "past") {
      tr.classList.add("past-cycle-row");
    }
    if (rec.isAdvance) {
      tr.classList.add("row-advance");
    }

    const badgeClass = rec.isAdvance
      ? "advance"
      : rec.cycleStatus === "current"
      ? "current"
      : rec.cycleStatus === "future"
      ? "future"
      : "past";

    tr.innerHTML = `
      <td class="col-no">${idx + 1}</td>
      <td class="col-cycle">
        <span class="cycle-badge ${badgeClass}">${escapeHtml(rec.cycleBadgeLabel)}</span>
      </td>
      <td class="col-date">${escapeHtml(rec.date)}</td>
      <td class="col-type"><span class="type-tag">${escapeHtml(rec.type || "연차")}</span></td>
      <td class="col-days">${formatNumber(rec.daysNum)}일</td>
      <td class="col-note">${escapeHtml(rec.note || "-")}</td>
      <td class="col-action">
        <button type="button" class="btn-row-del" data-id="${escapeHtml(rec.id)}">삭제</button>
      </td>
    `;
    els.leaveHistoryTbody.appendChild(tr);
  });
}

function renderLawReferenceTable(currentYearOrdinal) {
  els.lawReferenceTbody.innerHTML = "";
  const rows = [
    { ord: 1, full: "1년 미만", days: "최대 11일", desc: "1개월 개근 시 1일씩 발생 (초과 시 2년차 연차에서 차감)" },
    { ord: 2, full: "만 1년", days: "15일", desc: "기본 15일 부여 (1년차 초과 당겨쓰기분 차감 반영)" },
    { ord: 3, full: "만 2년", days: "16일", desc: "기본 15일 + 가산 1일 (매년 1일 증가)" },
    { ord: 4, full: "만 3년", days: "17일", desc: "기본 15일 + 가산 2일" },
    { ord: 5, full: "만 4년", days: "18일", desc: "기본 15일 + 가산 3일" },
    { ord: 6, full: "만 5년", days: "19일", desc: "기본 15일 + 가산 4일" },
    { ord: 7, full: "만 6년", days: "20일", desc: "기본 15일 + 가산 5일" },
    { ord: 8, full: "만 7년", days: "21일", desc: "기본 15일 + 가산 6일" },
    { ord: 9, full: "만 8년", days: "22일", desc: "기본 15일 + 가산 7일" },
    { ord: 10, full: "만 9년", days: "23일", desc: "기본 15일 + 가산 8일" },
    { ord: 11, full: "만 10년", days: "24일", desc: "기본 15일 + 가산 9일" },
    { ord: 12, full: "만 11년 이상", days: "25일", desc: "최대 상한 (25일)" }
  ];

  rows.forEach((r) => {
    const tr = document.createElement("tr");
    if (currentYearOrdinal && (currentYearOrdinal === r.ord || (r.ord === 12 && currentYearOrdinal >= 12))) {
      tr.classList.add("active-tenure-row");
    }
    tr.innerHTML = `
      <td>${r.ord === 12 ? "12년차 이상" : r.ord + "년차"}</td>
      <td>${r.full}</td>
      <td>${r.days}</td>
      <td>${r.desc}</td>
    `;
    els.lawReferenceTbody.appendChild(tr);
  });
}

function escapeHtml(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function showToast(msg) {
  const div = document.createElement("div");
  div.className = "toast";
  div.textContent = msg;
  els.toastContainer.appendChild(div);
  setTimeout(() => {
    div.remove();
  }, 3200);
}

// ============================================================================
// 4. 로컬 PC 엑셀 저장 ("연차 사용이력 기록용.xlsx") 및 불러오기
// ============================================================================

function buildExcelFileName() {
  return "연차 사용이력 기록용.xlsx";
}

function exportToExcel() {
  if (typeof XLSX === "undefined") {
    alert("엑셀 라이브러리를 불러오는 중입니다. 잠시 후 다시 시도해 주세요.");
    return;
  }

  const hireDateObj = parseDate(state.hireDate);
  const refDateObj = parseDate(els.referenceDateInput.value) || parseDate(getTodayStr());
  const calcResult =
    hireDateObj && refDateObj >= hireDateObj
      ? calculateCurrentCycleAndLeave(hireDateObj, refDateObj, {
          calcBasis: state.calcBasis,
          firstYearMode: state.firstYearMode,
          fiscalIncludeFirstYearMonthly: state.fiscalIncludeFirstYearMonthly,
          manualAdjust: state.manualAdjust
        })
      : null;

  const sortedRecords = [...state.records].sort((a, b) => a.date.localeCompare(b.date));

  const cycleAccumMap = {};
  let currentCycleUsed = 0;

  const historyRows = sortedRecords.map((rec, index) => {
    const daysNum = parseFloat(rec.days) || 0;
    const recDate = parseDate(rec.date);

    let cycleLabel = "-";
    let isCurrent = "N";
    let isAdvance = false;
    let advancePortion = 0;

    if (calcResult && hireDateObj && recDate) {
      const recCycleOrd = getRecordCycleOrdinal(recDate, hireDateObj, state.calcBasis);
      const cycleDef = getCycleBoundsAndCap(recCycleOrd, hireDateObj, state.calcBasis, {
        firstYearMode: state.firstYearMode,
        fiscalIncludeFirstYearMonthly: state.fiscalIncludeFirstYearMonthly
      });
      const cap = cycleDef ? cycleDef.cap : 999;

      const prevAcc = cycleAccumMap[recCycleOrd] || 0;
      const newAcc = Math.round((prevAcc + daysNum) * 100) / 100;
      cycleAccumMap[recCycleOrd] = newAcc;

      const targetNextOrd = recCycleOrd + 1;
      if (prevAcc >= cap) {
        isAdvance = true;
        advancePortion = daysNum;
      } else if (newAcc > cap) {
        isAdvance = true;
        advancePortion = Math.round((newAcc - cap) * 100) / 100;
      }

      if (recCycleOrd === calcResult.tenure.tenureYearOrdinal) {
        isCurrent = "Y";
        currentCycleUsed += daysNum;
      }

      if (isAdvance) {
        cycleLabel = `${recCycleOrd}년차(${targetNextOrd}년차 끌)`;
      } else {
        if (isCurrent === "Y") {
          cycleLabel = `${calcResult.tenure.tenureYearOrdinal}년차 (현재)`;
        } else if (recCycleOrd > calcResult.tenure.tenureYearOrdinal) {
          cycleLabel = `${recCycleOrd}년차 (예정)`;
        } else {
          cycleLabel =
            state.calcBasis === "fiscal"
              ? `이전 (${recCycleOrd}년차·${recDate.getFullYear()}년)`
              : `이전 (${recCycleOrd}년차)`;
        }
      }
    }

    return {
      "번호": index + 1,
      "사용일자": rec.date,
      "일수": daysNum,
      "비고": rec.note || "",
      "휴가구분": rec.type || "연차",
      "연차구분": cycleLabel,
      "현재연차포함": isCurrent
    };
  });

  const wsHistory =
    historyRows.length > 0
      ? XLSX.utils.json_to_sheet(historyRows)
      : XLSX.utils.aoa_to_sheet([["번호", "사용일자", "일수", "비고", "휴가구분", "연차구분", "현재연차포함"]]);

  wsHistory["!cols"] = [
    { wch: 8 },
    { wch: 15 },
    { wch: 10 },
    { wch: 30 },
    { wch: 14 },
    { wch: 24 },
    { wch: 15 }
  ];

  // 이전 연차 초과분(끌어쓰기) 순차 차감 계산 (ONSE 사내 규정 반영: 1년차 포함)
  let totalPastAdvanceDeficit = 0;
  if (calcResult && hireDateObj) {
    let carryoverDeficit = 0;
    for (let ord = 1; ord <= calcResult.tenure.tenureYearOrdinal; ord++) {
      if (ord === calcResult.tenure.tenureYearOrdinal) {
        totalPastAdvanceDeficit = carryoverDeficit;
        break;
      }
      const cycleDef = getCycleBoundsAndCap(ord, hireDateObj, state.calcBasis, {
        firstYearMode: state.firstYearMode,
        fiscalIncludeFirstYearMonthly: state.fiscalIncludeFirstYearMonthly
      });
      const baseCap = cycleDef ? cycleDef.cap : (ord === 1 ? 11 : 15);
      const availableInCycle = Math.max(0, baseCap - carryoverDeficit);
      const usedInCycle = cycleAccumMap[ord] || 0;
      if (usedInCycle > availableInCycle) {
        carryoverDeficit = Math.round((usedInCycle - availableInCycle) * 100) / 100;
      } else {
        carryoverDeficit = 0;
      }
    }
  }
  totalPastAdvanceDeficit = Math.round(totalPastAdvanceDeficit * 100) / 100;

  const advanceDeductedDays = state.autoDeductAdvance ? totalPastAdvanceDeficit : 0;
  const baseLeave = calcResult ? calcResult.baseLeave : 0;
  const manualAdjust = calcResult ? calcResult.manualAdjust : 0;
  const netGenerated = Math.max(0, Math.round((baseLeave - advanceDeductedDays + manualAdjust) * 100) / 100);
  const remainingLeave = Math.round(
    (baseLeave - advanceDeductedDays + manualAdjust - currentCycleUsed) * 100
  ) / 100;

  const summaryAoA = [
    ["항목", "값", "설명"],
    ["사용자명", state.userName || "-", "사용자 이름/직급"],
    ["입사일", state.hireDate || "", "YYYY-MM-DD 형식"],
    ["산정기준", state.calcBasis === "fiscal" ? "회계년도 기준" : "입사년도 기준", "입사년도 기준(디폴트) / 회계년도 기준"],
    ["산정기준코드", state.calcBasis, "시스템 복원용 코드 (anniversary / fiscal)"],
    ["1년차표시방식", state.firstYearMode, "accrued(매월 1개씩 발생) / max11(11개 한도 표시)"],
    ["회계2년차월차합산", state.fiscalIncludeFirstYearMonthly ? "Y" : "N", "회계년도 2년차에 1년 미만 월차 합산 여부"],
    ["이전연차_끌어쓰기자동차감", state.autoDeductAdvance ? "Y" : "N", "이전 연차 초과분(끌어쓰기) 차기 연차에서 자동 차감 여부 (1년차는 리셋)"],
    ["이전연차_끌어쓰기차감일수", advanceDeductedDays, "차기 연차에서 차감된 일수"],
    ["수동보정일수", state.manualAdjust || 0, "발생 연차 수동 가감 일수"],
    ["기준일자(오늘)", formatDate(refDateObj), "근속 및 연차 산정 기준일"],
    ["근속연차", calcResult ? `${calcResult.tenure.tenureYearOrdinal}년차` : "-", calcResult ? `만 ${calcResult.tenure.exactYears}년 ${calcResult.tenure.remMonths}개월` : "-"],
    ["근속기간(개월차)", calcResult ? `${calcResult.tenure.tenureMonthOrdinal}개월차` : "-", calcResult ? `만 ${calcResult.tenure.exactFullMonths}개월 경과` : "-"],
    ["근속기간(일수)", calcResult ? `${calcResult.tenure.totalDaysInclusive}일` : "-", "입사일 포함 총 재직일수"],
    [
      "현재적용연차기간",
      calcResult ? `${formatDate(calcResult.cycleStart)} ~ ${formatDate(calcResult.cycleEnd)}` : "-",
      calcResult ? calcResult.cycleLabel : "-"
    ],
    ["기본발생_일수", baseLeave, calcResult ? calcResult.formulaDesc : "-"],
    ["현재연차_실질발생일수", netGenerated, "기본 발생일수 - 이전연차 끌어쓰기 차감일수 + 수동보정"],
    ["현재연차_사용일수", Math.round(currentCycleUsed * 100) / 100, "현재 연차 범위 내 사용일수 합계"],
    ["현재연차_잔여일수", remainingLeave, "실질발생일수 - 현재연차 사용일수"]
  ];

  const wsSummary = XLSX.utils.aoa_to_sheet(summaryAoA);
  wsSummary["!cols"] = [{ wch: 22 }, { wch: 28 }, { wch: 44 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsHistory, "연차사용이력");
  XLSX.utils.book_append_sheet(wb, wsSummary, "설정및연차요약");

  const fileName = buildExcelFileName();
  XLSX.writeFile(wb, fileName);

  const now = new Date();
  const timeStr = `${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`;
  state.lastExcelFileName = fileName;
  state.lastExcelActionTime = `저장 완료 ${timeStr}`;
  saveToLocalStorage();
  renderSavedPathBar();

  showToast(`💾 '${fileName}' 파일이 내 PC 다운로드 폴더에 저장되었습니다.`);
}

function normalizeExcelDateValue(val) {
  if (!val) return "";
  if (typeof val === "number") {
    const parsed = XLSX.SSF.parse_date_code(val);
    if (parsed && parsed.y && parsed.m && parsed.d) {
      return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
  }
  const str = String(val).trim();
  const dObj = parseDate(str);
  return dObj ? formatDate(dObj) : "";
}

function importFromExcel(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: "array" });

      if (wb.SheetNames.includes("설정및연차요약")) {
        const wsSum = wb.Sheets["설정및연차요약"];
        const sumRows = XLSX.utils.sheet_to_json(wsSum, { header: 1 });
        const kv = {};
        sumRows.forEach((row) => {
          if (row && row.length >= 2 && row[0]) {
            kv[String(row[0]).trim()] = row[1];
          }
        });

        if (kv["사용자명"] && kv["사용자명"] !== "-") {
          state.userName = String(kv["사용자명"]).trim();
        }
        if (kv["입사일"]) {
          const normalizedHire = normalizeExcelDateValue(kv["입사일"]);
          if (normalizedHire) state.hireDate = normalizedHire;
        }
        if (kv["산정기준코드"]) {
          state.calcBasis = String(kv["산정기준코드"]).trim() === "fiscal" ? "fiscal" : "anniversary";
        } else if (kv["산정기준"]) {
          state.calcBasis = String(kv["산정기준"]).includes("회계") ? "fiscal" : "anniversary";
        }
        if (kv["1년차표시방식"]) {
          state.firstYearMode = String(kv["1년차표시방식"]).trim() === "max11" ? "max11" : "accrued";
        }
        if (kv["회계2년차월차합산"]) {
          state.fiscalIncludeFirstYearMonthly = String(kv["회계2년차월차합산"]).trim() !== "N";
        }
        const autoDeductKey = kv["이전연차_끌어쓰기자동차감"] || kv["이전연차_先사용자동차감"];
        if (autoDeductKey) {
          state.autoDeductAdvance = String(autoDeductKey).trim() !== "N";
        }
        if (kv["수동보정일수"] !== undefined) {
          state.manualAdjust = parseFloat(kv["수동보정일수"]) || 0;
        }
      }

      const historySheetName = wb.SheetNames.includes("연차사용이력") ? "연차사용이력" : wb.SheetNames[0];
      const wsHist = wb.Sheets[historySheetName];
      const rawJson = XLSX.utils.sheet_to_json(wsHist, { defval: "" });

      const importedRecords = [];
      rawJson.forEach((row, idx) => {
        const rawDate = row["사용일자"] || row["날짜"] || row["일자"] || row["Date"] || "";
        const normDate = normalizeExcelDateValue(rawDate);
        if (!normDate) return;

        const rawDays = row["일수"] !== undefined && row["일수"] !== "" ? row["일수"] : row["사용일수"];
        const daysNum = parseFloat(rawDays);
        if (isNaN(daysNum) || daysNum <= 0) return;

        const note = String(row["비고"] || row["사유"] || row["메모"] || "").trim();
        const type = String(
          row["휴가구분"] || row["구분"] || (daysNum === 0.5 ? "반차" : daysNum === 0.25 ? "반반차" : "연차")
        ).trim();

        importedRecords.push({
          id: `imp_${Date.now()}_${idx}`,
          date: normDate,
          type,
          days: daysNum,
          note
        });
      });

      const now = new Date();
      const timeStr = `${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`;
      state.records = importedRecords;
      state.lastExcelFileName = file.name;
      state.lastExcelActionTime = `불러오기 완료 ${timeStr}`;

      saveToLocalStorage();
      renderAll();
      showToast(`📂 '${file.name}' 파일에서 연차 이력 ${importedRecords.length}건을 불러왔습니다.`);
    } catch (err) {
      console.error("엑셀 파일 읽기 오류:", err);
      alert("엑셀 파일을 읽는 중 오류가 발생했습니다. 올바른 연차 사용 이력 엑셀 파일인지 확인해 주세요.");
    } finally {
      els.excelFileInput.value = "";
    }
  };
  reader.readAsArrayBuffer(file);
}

// ============================================================================
// 5. 다중 행(Batch) 연차 입력 관리
// ============================================================================

function createEntryRowElement(initialData = {}) {
  const rowDiv = document.createElement("div");
  rowDiv.className = "entry-row-item";

  const defaultDate = initialData.date || getTodayStr();
  const defaultType = initialData.type || "연차|1";
  const defaultDays = initialData.days !== undefined ? initialData.days : 1;
  const defaultNote = initialData.note || "";

  rowDiv.innerHTML = `
    <input type="date" class="form-control-compact input-row-date" value="${defaultDate}" required />
    <select class="form-control-compact input-row-type">
      <option value="연차|1" ${defaultType === "연차|1" ? "selected" : ""}>연차 (1.0일)</option>
      <option value="오전반차|0.5" ${defaultType === "오전반차|0.5" ? "selected" : ""}>오전반차 (0.5일)</option>
      <option value="오후반차|0.5" ${defaultType === "오후반차|0.5" ? "selected" : ""}>오후반차 (0.5일)</option>
      <option value="반반차|0.25" ${defaultType === "반반차|0.25" ? "selected" : ""}>반반차 (0.25일)</option>
      <option value="직접입력|custom" ${defaultType.includes("custom") ? "selected" : ""}>직접 입력</option>
    </select>
    <input type="number" class="form-control-compact input-days input-row-days" step="0.25" min="0.25" max="365" value="${defaultDays}" required />
    <input type="text" class="form-control-compact input-row-note" placeholder="사유/메모" value="${defaultNote}" />
    <button type="button" class="btn-remove-row" title="행 삭제">✕</button>
  `;

  const typeSelect = rowDiv.querySelector(".input-row-type");
  const daysInput = rowDiv.querySelector(".input-row-days");
  const removeBtn = rowDiv.querySelector(".btn-remove-row");

  typeSelect.addEventListener("change", (e) => {
    const [, defaultDays] = e.target.value.split("|");
    if (defaultDays !== "custom") {
      daysInput.value = parseFloat(defaultDays);
    } else {
      daysInput.focus();
      daysInput.select();
    }
  });

  removeBtn.addEventListener("click", () => {
    if (!els.entryRowsContainer) return;
    const allRows = els.entryRowsContainer.querySelectorAll(".entry-row-item");
    if (allRows.length <= 1) {
      // 1개 행만 남아있을 때는 행을 삭제하지 않고 기본값으로 리셋
      const dateInp = rowDiv.querySelector(".input-row-date");
      const noteInp = rowDiv.querySelector(".input-row-note");
      dateInp.value = getTodayStr();
      typeSelect.value = "연차|1";
      daysInput.value = 1;
      noteInp.value = "";
    } else {
      rowDiv.remove();
    }
  });

  return rowDiv;
}

function addEntryRow(initialData = {}) {
  if (!els.entryRowsContainer) return null;
  // 이전 행의 사용일자가 있으면 디폴트 날짜로 계승하여 연속 입력 지원
  if (!initialData.date) {
    const rows = els.entryRowsContainer.querySelectorAll(".entry-row-item");
    if (rows.length > 0) {
      const lastRow = rows[rows.length - 1];
      const lastDate = lastRow.querySelector(".input-row-date").value;
      if (lastDate) initialData.date = lastDate;
    }
  }

  const rowEl = createEntryRowElement(initialData);
  els.entryRowsContainer.appendChild(rowEl);
  els.entryRowsContainer.scrollTop = els.entryRowsContainer.scrollHeight;
  return rowEl;
}

function initEntryRows() {
  if (!els.entryRowsContainer) return;
  els.entryRowsContainer.innerHTML = "";
  addEntryRow({ date: getTodayStr(), type: "연차|1", days: 1, note: "" });
}

// ============================================================================
// 6. 이벤트 리스너 바인딩 및 초기화
// ============================================================================

function initEvents() {
  els.referenceDateInput.addEventListener("change", () => {
    renderAll();
  });

  els.resetTodayBtn.addEventListener("click", () => {
    els.referenceDateInput.value = getTodayStr();
    renderAll();
    showToast("기준일이 오늘 날짜로 설정되었습니다.");
  });

  els.userNameInput.addEventListener("input", (e) => {
    state.userName = e.target.value.trim();
    saveToLocalStorage();
  });

  els.hireDateInput.addEventListener("change", (e) => {
    state.hireDate = e.target.value;
    saveToLocalStorage();
    renderAll();
  });

  els.basisRadios.forEach((radio) => {
    radio.addEventListener("change", (e) => {
      if (e.target.checked) {
        state.calcBasis = e.target.value;
        saveToLocalStorage();
        renderAll();
      }
    });
  });

  els.firstYearModeSelect.addEventListener("change", (e) => {
    state.firstYearMode = e.target.value;
    saveToLocalStorage();
    renderAll();
  });

  els.fiscalIncludeFirstYearMonthlyChk.addEventListener("change", (e) => {
    state.fiscalIncludeFirstYearMonthly = e.target.checked;
    saveToLocalStorage();
    renderAll();
  });

  els.manualAdjustInput.addEventListener("input", (e) => {
    state.manualAdjust = parseFloat(e.target.value) || 0;
    saveToLocalStorage();
    renderAll();
  });

  els.autoDownloadExcelChk.addEventListener("change", (e) => {
    state.autoDownloadExcel = e.target.checked;
    saveToLocalStorage();
  });

  if (els.autoDeductAdvanceChk) {
    els.autoDeductAdvanceChk.addEventListener("change", (e) => {
      state.autoDeductAdvance = e.target.checked;
      saveToLocalStorage();
      renderAll();
      showToast(
        state.autoDeductAdvance
          ? "이전 연차 초과분(끌어쓰기) 자동 차감이 활성화되었습니다."
          : "이전 연차 초과분(끌어쓰기) 자동 차감이 비활성화되었습니다."
      );
    });
  }

  if (els.addInputRowTopBtn) {
    els.addInputRowTopBtn.addEventListener("click", () => addEntryRow());
  }

  if (els.addInputRowBtn) {
    els.addInputRowBtn.addEventListener("click", () => addEntryRow());
  }

  els.leaveRecordForm.addEventListener("submit", (e) => {
    e.preventDefault();
    if (!els.entryRowsContainer) return;

    const rowEls = els.entryRowsContainer.querySelectorAll(".entry-row-item");
    const validRecords = [];

    for (let i = 0; i < rowEls.length; i++) {
      const row = rowEls[i];
      const dateVal = row.querySelector(".input-row-date").value.trim();
      const typeVal = row.querySelector(".input-row-type").value;
      const daysVal = parseFloat(row.querySelector(".input-row-days").value);
      const noteVal = row.querySelector(".input-row-note").value.trim();

      // 전체가 빈 행이면 스킵
      if (!dateVal && !noteVal && isNaN(daysVal)) {
        continue;
      }

      if (!dateVal) {
        alert(`${i + 1}번째 행의 사용일자를 선택해 주세요.`);
        row.querySelector(".input-row-date").focus();
        return;
      }

      if (isNaN(daysVal) || daysVal <= 0) {
        alert(`${i + 1}번째 행의 올바른 사용 일수를 입력해 주세요.`);
        row.querySelector(".input-row-days").focus();
        return;
      }

      const [typeName] = typeVal.split("|");
      validRecords.push({
        id: `rec_${Date.now()}_${i}_${Math.random().toString(36).slice(2, 7)}`,
        date: dateVal,
        type: typeName,
        days: daysVal,
        note: noteVal
      });
    }

    if (validRecords.length === 0) {
      alert("저장할 연차 사용 내역을 입력해 주세요.");
      return;
    }

    state.records.push(...validRecords);
    saveToLocalStorage();
    renderAll();

    initEntryRows();
    showToast(`💾 총 ${validRecords.length}건의 연차 사용 내역이 저장되었습니다.`);

    if (state.autoDownloadExcel) {
      exportToExcel();
    }
  });

  els.leaveHistoryTbody.addEventListener("click", (e) => {
    const btn = e.target.closest(".btn-row-del");
    if (!btn) return;
    const targetId = btn.getAttribute("data-id");
    const targetRec = state.records.find((r) => r.id === targetId);
    if (!targetRec) return;

    state.records = state.records.filter((r) => r.id !== targetId);
    saveToLocalStorage();
    renderAll();
    showToast(`🗑️ ${targetRec.date} 연차 기록이 삭제되었습니다.`);

    if (state.autoDownloadExcel) {
      exportToExcel();
    }
  });

  els.filterTabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      els.filterTabs.forEach((t) => t.classList.remove("active"));
      tab.classList.add("active");
      state.historyFilter = tab.getAttribute("data-filter") || "all";
      renderAll();
    });
  });

  els.clearAllRecordsBtn.addEventListener("click", () => {
    if (state.records.length === 0) {
      showToast("삭제할 연차 사용 기록이 없습니다.");
      return;
    }
    if (
      confirm(
        "등록된 모든 연차 사용 기록을 초기화하시겠습니까?\n(초기화 전 '엑셀 DB 저장하기'로 백업해 두시는 것을 권장합니다.)"
      )
    ) {
      state.records = [];
      saveToLocalStorage();
      renderAll();
      showToast("모든 연차 사용 기록이 초기화되었습니다.");
    }
  });

  els.exportExcelBtn.addEventListener("click", exportToExcel);
  els.exportExcelBottomBtn.addEventListener("click", exportToExcel);

  els.importExcelBtn.addEventListener("click", () => {
    els.excelFileInput.click();
  });

  els.excelFileInput.addEventListener("change", (e) => {
    const file = e.target.files && e.target.files[0];
    if (file) {
      importFromExcel(file);
    }
  });
}

function initApp() {
  const todayStr = getTodayStr();
  els.referenceDateInput.value = todayStr;

  const hasSaved = loadFromLocalStorage();
  if (!hasSaved) {
    state.calcBasis = "anniversary";
  }

  initEntryRows();
  initEvents();
  renderAll();
}

document.addEventListener("DOMContentLoaded", initApp);
