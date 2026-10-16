/**
 * 스마트 연차계산기 & 로컬 엑셀 DB 관리 애플리케이션
 * - 입사년도(입사일) 기준 (디폴트) / 회계년도(1.1) 기준 선택 지원
 * - 오늘 기준 근속연차(N년차), 근속기간(N개월차), 근속일수(N일) 자동 계산
 * - 전체 연차 사용 기록은 쭉 누적하되, 이번 주기 범위 내 기록만 필터링하여 잔여 연차 계산
 * - 로컬 PC 엑셀 파일("연차 사용 이력_(업데이트 날짜, 시간).xlsx") 저장 및 불러오기 지원
 */

const STORAGE_KEY = "ONSE_ANNUAL_LEAVE_CALCULATOR_V1";

// 기본 상태값
const state = {
  userName: "",
  hireDate: "", // YYYY-MM-DD
  calcBasis: "anniversary", // 'anniversary' (입사년도 기본값) | 'fiscal' (회계년도)
  firstYearMode: "accrued", // 'accrued' (경과 개월수 기준) | 'max11' (1년차 최대 11개 선반영)
  fiscalIncludeFirstYearMonthly: true,
  manualAdjust: 0,
  autoDownloadExcel: false,
  historyFilter: "all", // 'all' | 'current' | 'past'
  records: [] // [{ id, date, type, days, note }]
};

// DOM Elements
const els = {
  referenceDateInput: document.getElementById("referenceDateInput"),
  resetTodayBtn: document.getElementById("resetTodayBtn"),
  autoDownloadExcelChk: document.getElementById("autoDownloadExcelChk"),
  excelFileInput: document.getElementById("excelFileInput"),
  importExcelBtn: document.getElementById("importExcelBtn"),
  exportExcelBtn: document.getElementById("exportExcelBtn"),
  exportExcelBottomBtn: document.getElementById("exportExcelBottomBtn"),
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

  generatedLeaveVal: document.getElementById("generatedLeaveVal"),
  generatedLeaveDesc: document.getElementById("generatedLeaveDesc"),
  usedLeaveVal: document.getElementById("usedLeaveVal"),
  usedLeaveDesc: document.getElementById("usedLeaveDesc"),
  remainingLeaveVal: document.getElementById("remainingLeaveVal"),
  remainingLeaveDesc: document.getElementById("remainingLeaveDesc"),
  remainingStatCard: document.querySelector(".stat-remaining"),

  usageRateText: document.getElementById("usageRateText"),
  usageProgressBar: document.getElementById("usageProgressBar"),

  leaveRecordForm: document.getElementById("leaveRecordForm"),
  leaveUseDate: document.getElementById("leaveUseDate"),
  leaveTypeSelect: document.getElementById("leaveTypeSelect"),
  leaveDaysInput: document.getElementById("leaveDaysInput"),
  leaveNoteInput: document.getElementById("leaveNoteInput"),

  filterTabs: document.querySelectorAll(".filter-tab"),
  countAllBadge: document.getElementById("countAllBadge"),
  countCurrentBadge: document.getElementById("countCurrentBadge"),
  countPastBadge: document.getElementById("countPastBadge"),

  stripCurrentUsed: document.getElementById("stripCurrentUsed"),
  stripPastUsed: document.getElementById("stripPastUsed"),
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
  // 윤년(2월 29일) 입사자가 평년을 맞이할 경우 2월 28일로 안전하게 매핑
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

  // 총 일수 (입사 당일을 1일째로 산정)
  const elapsedDays = diffDaysBetween(hireDate, refDate);
  const totalDaysInclusive = elapsedDays + 1;

  // 만 개월수 및 만 년수 계산
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
 * 만 근속년수(fullYears, 1 이상)에 따른 법정 연차 일수 계산 (15일 ~ 최대 25일)
 */
function getStatutoryLeaveForFullYears(fullYears) {
  if (fullYears < 1) return 0;
  const extraDays = Math.floor((fullYears - 1) / 2);
  return Math.min(25, 15 + extraDays);
}

/**
 * 현재 산정 주기(Cycle)의 시작일·종료일 및 발생 연차를 계산합니다.
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
  let cycleLabel = "";

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
    cycleLabel = `${yearOrdinal}년차 주기`;

    if (fullYearsAtCycleStart === 0) {
      // 1년차 (만 1년 미만): 1개월 개근 시 1일씩 발생 (최대 11일)
      const accruedMonthly = Math.min(11, tenure.exactFullMonths);
      if (firstYearMode === "max11") {
        baseLeave = 11;
        formulaDesc = `1년차 최대 월차 11일 선반영 (현재 만 ${tenure.exactFullMonths}개월 경과)`;
      } else {
        baseLeave = accruedMonthly;
        formulaDesc = `1년차 월차: 만 ${tenure.exactFullMonths}개월 경과분 (최대 11일 중 ${accruedMonthly}일 발생)`;
      }
    } else {
      // 2년차 이상 (만 1년 이상): 기본 15일 + 만 3년 이상부터 2년마다 1일 가산 (최대 25일)
      baseLeave = getStatutoryLeaveForFullYears(fullYearsAtCycleStart);
      const extra = baseLeave - 15;
      formulaDesc =
        extra > 0
          ? `${yearOrdinal}년차 (만 ${fullYearsAtCycleStart}년): 기본 15일 + 가산 ${extra}일`
          : `${yearOrdinal}년차 (만 ${fullYearsAtCycleStart}년): 법정 기본 15일`;
    }
  } else {
    // 2) 회계년도(1월 1일) 기준
    cycleLabel = `${refY}년 회계주기`;

    if (refY === hireY) {
      // 입사 당해 연도 (입사일 ~ 12/31)
      cycleStart = new Date(hireDate.getTime());
      cycleEnd = new Date(refY, 11, 31, 0, 0, 0, 0);

      const maxMonthsInFirstFiscalYear = 11 - hireM;
      const accruedMonthly = Math.min(11, tenure.exactFullMonths);

      if (firstYearMode === "max11") {
        baseLeave = Math.max(0, maxMonthsInFirstFiscalYear);
        formulaDesc = `입사 첫해 회계기준 월차 (연말까지 최대 ${baseLeave}일)`;
      } else {
        baseLeave = accruedMonthly;
        formulaDesc = `입사 첫해 월차: 만 ${tenure.exactFullMonths}개월 경과분`;
      }
    } else if (refY === hireY + 1) {
      // 입사 2차년도 (다음 해 1월 1일 ~ 12월 31일)
      cycleStart = new Date(refY, 0, 1, 0, 0, 0, 0);
      cycleEnd = new Date(refY, 11, 31, 0, 0, 0, 0);

      // 1월 1일 입사자는 만 1년을 채웠으므로 15일 온전 부여
      if (hireM === 0 && hireD === 1) {
        baseLeave = 15;
        formulaDesc = `회계 2년차 (1.1 입사 만 1년): 기본 15일`;
      } else {
        // 전년도 재직일수 비례 부여: 15 * (전년도 재직일수 / 전년도 총일수)
        const startOfHireYear = new Date(hireY, 0, 1);
        const endOfHireYear = new Date(hireY, 11, 31);
        const totalDaysInHireYear = diffDaysBetween(startOfHireYear, endOfHireYear) + 1;
        const workedDaysInHireYear = diffDaysBetween(hireDate, endOfHireYear) + 1;
        const proportionalLeave = Math.round(15 * (workedDaysInHireYear / totalDaysInHireYear) * 10) / 10;

        // 입사 후 1년 되기 전까지 당해 연도(2차년도)에 발생하는 월차 개수
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
            ? `회계 2년차 비례연차 ${proportionalLeave}일 + 당해 발생 월차 ${secondYearMonthly}일`
            : `회계 2년차 비례연차 (${workedDaysInHireYear}/${totalDaysInHireYear}일 비례)`;
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
          ? `${refY}년 회계기준 (만 ${fullYearsAtJan1}년+): 기본 15일 + 가산 ${extra}일`
          : `${refY}년 회계기준 (만 ${fullYearsAtJan1}년): 법정 기본 15일`;
    }
  }

  const adj = parseFloat(manualAdjust) || 0;
  const totalGenerated = Math.max(0, Math.round((baseLeave + adj) * 100) / 100);
  if (adj !== 0) {
    const sign = adj > 0 ? `+${adj}` : `${adj}`;
    formulaDesc += ` (수동보정 ${sign}일 반영)`;
  }

  // 다음 주기 리셋까지 남은 일수
  const daysUntilReset = diffDaysBetween(refDate, cycleEnd) + 1;

  return {
    tenure,
    cycleStart,
    cycleEnd,
    cycleLabel,
    baseLeave,
    manualAdjust: adj,
    totalGenerated,
    formulaDesc,
    daysUntilReset
  };
}

/**
 * 개별 연차 사용 기록이 어느 주기에 속하는지 판별합니다.
 */
function classifyRecordCycle(recordDateStr, hireDate, cycleStart, cycleEnd, calcBasis) {
  const recDate = parseDate(recordDateStr);
  if (!recDate || !hireDate || !cycleStart || !cycleEnd) {
    return { status: "unknown", label: "미분류" };
  }

  if (recDate >= cycleStart && recDate <= cycleEnd) {
    return { status: "current", label: "이번 주기" };
  }

  if (recDate > cycleEnd) {
    return { status: "future", label: "다음 주기(예정)" };
  }

  // 이전 주기인 경우 몇 년차/몇 년도 주기였는지 표시
  if (calcBasis === "fiscal") {
    return { status: "past", label: `이전 (${recDate.getFullYear()}년)` };
  } else {
    const hireY = hireDate.getFullYear();
    const hireM = hireDate.getMonth();
    const hireD = hireDate.getDate();
    const recY = recDate.getFullYear();
    const annivInRecYear = getSafeAnniversary(recY, hireM, hireD);
    const cycleStartYear = recDate < annivInRecYear ? recY - 1 : recY;
    const ord = cycleStartYear - hireY + 1;
    return {
      status: "past",
      label: ord >= 1 ? `이전 (${ord}년차)` : "입사 전"
    };
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
      autoDownloadExcel: state.autoDownloadExcel,
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
      state.autoDownloadExcel = Boolean(data.autoDownloadExcel);
      state.records = Array.isArray(data.records) ? data.records : [];
      return true;
    }
  } catch (err) {
    console.error("로컬 스토리지 로드 실패:", err);
  }
  return false;
}

function syncControlsFromState() {
  els.userNameInput.value = state.userName;
  els.hireDateInput.value = state.hireDate;
  els.firstYearModeSelect.value = state.firstYearMode;
  els.fiscalIncludeFirstYearMonthlyChk.checked = state.fiscalIncludeFirstYearMonthly;
  els.manualAdjustInput.value = state.manualAdjust;
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

    els.generatedLeaveVal.textContent = "0";
    els.generatedLeaveDesc.textContent = "입사일을 입력하면 자동 계산됩니다";
    els.usedLeaveVal.textContent = "0";
    els.usedLeaveDesc.textContent = "이번 주기 내 사용 0건";
    els.remainingLeaveVal.textContent = "0";
    els.remainingLeaveDesc.textContent = "사용 가능 연차";
    els.usageRateText.textContent = "0.0%";
    els.usageProgressBar.style.width = "0%";

    renderHistoryTable(null, null, null);
    renderLawReferenceTable(null);
    return;
  }

  const calcResult = calculateCurrentCycleAndLeave(hireDateObj, refDateObj, {
    calcBasis: state.calcBasis,
    firstYearMode: state.firstYearMode,
    fiscalIncludeFirstYearMonthly: state.fiscalIncludeFirstYearMonthly,
    manualAdjust: state.manualAdjust
  });

  const { tenure, cycleStart, cycleEnd, cycleLabel, totalGenerated, formulaDesc, daysUntilReset } = calcResult;

  // 1) 상단 산정 주기 및 리셋 D-Day 표시
  els.currentCyclePeriodText.textContent = `이번 산정 주기 (${cycleLabel}): ${formatDateDot(cycleStart)} – ${formatDateDot(cycleEnd)}`;
  els.nextResetDdayBadge.textContent = `다음 연차 갱신까지 ${daysUntilReset.toLocaleString("ko-KR")}일`;

  // 2) 근속연차 / 근속개월차 / 근속일수 표시
  els.tenureYearCount.textContent = `${tenure.tenureYearOrdinal.toLocaleString("ko-KR")}년차`;
  els.tenureExactYears.textContent = `만 ${tenure.exactYears}년 ${tenure.remMonths}개월 (${tenure.remDays}일)`;

  els.tenureMonthCount.textContent = `${tenure.tenureMonthOrdinal.toLocaleString("ko-KR")}개월차`;
  els.tenureExactMonths.textContent = `만 ${tenure.exactFullMonths.toLocaleString("ko-KR")}개월 경과`;

  els.tenureDayCount.textContent = `${tenure.totalDaysInclusive.toLocaleString("ko-KR")}일`;
  els.tenureExactDays.textContent = `입사일 포함 ${tenure.totalDaysInclusive.toLocaleString("ko-KR")}일째 (만 ${tenure.elapsedDays.toLocaleString("ko-KR")}일)`;

  // 3) 연차 사용 기록 주기 분류 및 합산
  let currentUsedDays = 0;
  let currentUsedCount = 0;
  let pastUsedDays = 0;
  let pastUsedCount = 0;
  let totalUsedDays = 0;

  // 날짜 오름차순(또는 최신순) 정렬: 표에서 흐름을 보기 좋게 사용일자 오름차순 정렬
  const sortedRecords = [...state.records].sort((a, b) => a.date.localeCompare(b.date));

  const enrichedRecords = sortedRecords.map((rec) => {
    const daysNum = parseFloat(rec.days) || 0;
    const cycleInfo = classifyRecordCycle(rec.date, hireDateObj, cycleStart, cycleEnd, state.calcBasis);

    totalUsedDays += daysNum;
    if (cycleInfo.status === "current") {
      currentUsedDays += daysNum;
      currentUsedCount += 1;
    } else if (cycleInfo.status === "past") {
      pastUsedDays += daysNum;
      pastUsedCount += 1;
    }

    return {
      ...rec,
      daysNum,
      cycleStatus: cycleInfo.status,
      cycleBadgeLabel: cycleInfo.label
    };
  });

  currentUsedDays = Math.round(currentUsedDays * 100) / 100;
  pastUsedDays = Math.round(pastUsedDays * 100) / 100;
  totalUsedDays = Math.round(totalUsedDays * 100) / 100;

  const remainingDays = Math.round((totalGenerated - currentUsedDays) * 100) / 100;

  // 4) 발생 / 사용 / 잔여 연차 카드 업데이트
  els.generatedLeaveVal.textContent = formatNumber(totalGenerated);
  els.generatedLeaveDesc.textContent = formulaDesc;

  els.usedLeaveVal.textContent = formatNumber(currentUsedDays);
  els.usedLeaveDesc.textContent = `이번 주기 내 사용 ${currentUsedCount}건 합산`;

  els.remainingLeaveVal.textContent = formatNumber(remainingDays);
  if (remainingDays < 0) {
    els.remainingStatCard.classList.add("negative-remaining");
    els.remainingLeaveDesc.textContent = `발생 연차 대비 ${formatNumber(Math.abs(remainingDays))}일 초과 사용`;
  } else {
    els.remainingStatCard.classList.remove("negative-remaining");
    els.remainingLeaveDesc.textContent = `${formatDateDot(cycleEnd)}까지 사용 가능`;
  }

  // 소진율 프로그레스 바
  let usageRate = 0;
  if (totalGenerated > 0) {
    usageRate = (currentUsedDays / totalGenerated) * 100;
  } else if (currentUsedDays > 0) {
    usageRate = 100;
  }
  els.usageRateText.textContent = `${usageRate.toFixed(1)}% (${formatNumber(currentUsedDays)} / ${formatNumber(totalGenerated)}일)`;
  els.usageProgressBar.style.width = `${Math.min(100, Math.max(0, usageRate))}%`;
  if (usageRate > 100) {
    els.usageProgressBar.classList.add("over-limit");
  } else {
    els.usageProgressBar.classList.remove("over-limit");
  }

  // 5) 우측 연차 사용 기록 요약 스트립 & 테이블 렌더링
  els.countAllBadge.textContent = enrichedRecords.length;
  els.countCurrentBadge.textContent = currentUsedCount;
  els.countPastBadge.textContent = pastUsedCount;

  els.stripCurrentUsed.textContent = `${formatNumber(currentUsedDays)}일 (${currentUsedCount}건)`;
  els.stripPastUsed.textContent = `${formatNumber(pastUsedDays)}일 (${pastUsedCount}건)`;
  els.stripTotalUsed.textContent = `${formatNumber(totalUsedDays)}일 (${enrichedRecords.length}건)`;

  renderHistoryTable(enrichedRecords, cycleStart, cycleEnd);
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

    const badgeClass =
      rec.cycleStatus === "current"
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
    { ord: 1, full: "1년 미만", days: "최대 11일", desc: "1개월 개근 시 1일씩 발생" },
    { ord: 2, full: "만 1년", days: "15일", desc: "법정 기본 연차 발생" },
    { ord: 3, full: "만 2년", days: "15일", desc: "기본 15일" },
    { ord: 4, full: "만 3년", days: "16일", desc: "기본 15일 + 가산 1일" },
    { ord: 5, full: "만 4년", days: "16일", desc: "기본 15일 + 가산 1일" },
    { ord: 6, full: "만 5년", days: "17일", desc: "기본 15일 + 가산 2일" },
    { ord: 7, full: "만 6년", days: "17일", desc: "기본 15일 + 가산 2일" },
    { ord: 8, full: "만 7년", days: "18일", desc: "기본 15일 + 가산 3일" },
    { ord: 9, full: "만 8년", days: "18일", desc: "기본 15일 + 가산 3일" },
    { ord: 10, full: "만 9년", days: "19일", desc: "기본 15일 + 가산 4일" },
    { ord: 12, full: "만 11년", days: "20일", desc: "기본 15일 + 가산 5일" },
    { ord: 14, full: "만 13년", days: "21일", desc: "기본 15일 + 가산 6일" },
    { ord: 16, full: "만 15년", days: "22일", desc: "기본 15일 + 가산 7일" },
    { ord: 18, full: "만 17년", days: "23일", desc: "기본 15일 + 가산 8일" },
    { ord: 20, full: "만 19년", days: "24일", desc: "기본 15일 + 가산 9일" },
    { ord: 22, full: "만 21년 이상", days: "25일", desc: "법정 최대 한도 (25일)" }
  ];

  rows.forEach((r) => {
    const tr = document.createElement("tr");
    if (currentYearOrdinal && (currentYearOrdinal === r.ord || (r.ord === 22 && currentYearOrdinal >= 22))) {
      tr.classList.add("active-tenure-row");
    }
    tr.innerHTML = `
      <td>${r.ord === 22 ? "22년차 이상" : r.ord + "년차"}</td>
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
// 4. 로컬 PC 엑셀 저장 ("연차 사용 이력_(업데이트 날짜, 시간).xlsx") 및 불러오기
// ============================================================================

function buildExcelFileName() {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  const hh = String(now.getHours()).padStart(2, "0");
  const min = String(now.getMinutes()).padStart(2, "0");
  return `연차 사용 이력_${yyyy}-${mm}-${dd}_${hh}시${min}분.xlsx`;
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

  // 날짜순 정렬된 전체 기록
  const sortedRecords = [...state.records].sort((a, b) => a.date.localeCompare(b.date));

  let currentCycleUsed = 0;
  const historyRows = sortedRecords.map((rec, index) => {
    const daysNum = parseFloat(rec.days) || 0;
    const cycleInfo = calcResult
      ? classifyRecordCycle(rec.date, hireDateObj, calcResult.cycleStart, calcResult.cycleEnd, state.calcBasis)
      : { status: "unknown", label: "-" };

    if (cycleInfo.status === "current") {
      currentCycleUsed += daysNum;
    }

    return {
      "번호": index + 1,
      "사용일자": rec.date,
      "일수": daysNum,
      "비고": rec.note || "",
      "휴가구분": rec.type || "연차",
      "적용주기": cycleInfo.label,
      "이번주기계산포함": cycleInfo.status === "current" ? "Y" : "N"
    };
  });

  // 기록이 하나도 없을 때도 컬럼 헤더가 유지되도록 빈 행 템플릿 준비
  const wsHistory =
    historyRows.length > 0
      ? XLSX.utils.json_to_sheet(historyRows)
      : XLSX.utils.aoa_to_sheet([["번호", "사용일자", "일수", "비고", "휴가구분", "적용주기", "이번주기계산포함"]]);

  wsHistory["!cols"] = [
    { wch: 8 },
    { wch: 15 },
    { wch: 10 },
    { wch: 30 },
    { wch: 14 },
    { wch: 16 },
    { wch: 18 }
  ];

  // 시트 2: 설정 및 현재 주기 연차 현황 요약 (불러오기 시 메타데이터 복원 겸용)
  const totalGenerated = calcResult ? calcResult.totalGenerated : 0;
  const remainingLeave = Math.round((totalGenerated - currentCycleUsed) * 100) / 100;

  const summaryAoA = [
    ["항목", "값", "설명"],
    ["사용자명", state.userName || "-", "사용자 이름/직급"],
    ["입사일", state.hireDate || "", "YYYY-MM-DD 형식"],
    ["산정기준", state.calcBasis === "fiscal" ? "회계년도 기준" : "입사년도 기준", "입사년도 기준(디폴트) / 회계년도 기준"],
    ["산정기준코드", state.calcBasis, "시스템 복원용 코드 (anniversary / fiscal)"],
    ["1년차표시방식", state.firstYearMode, "accrued(실시간 개월) / max11(11일 선반영)"],
    ["회계2년차월차합산", state.fiscalIncludeFirstYearMonthly ? "Y" : "N", "회계년도 2년차에 1년 미만 월차 합산 여부"],
    ["수동보정일수", state.manualAdjust || 0, "발생 연차 수동 가감 일수"],
    ["기준일자(오늘)", formatDate(refDateObj), "근속 및 연차 산정 기준일"],
    ["근속연차", calcResult ? `${calcResult.tenure.tenureYearOrdinal}년차` : "-", calcResult ? `만 ${calcResult.tenure.exactYears}년 ${calcResult.tenure.remMonths}개월` : "-"],
    ["근속기간(개월차)", calcResult ? `${calcResult.tenure.tenureMonthOrdinal}개월차` : "-", calcResult ? `만 ${calcResult.tenure.exactFullMonths}개월 경과` : "-"],
    ["근속기간(일수)", calcResult ? `${calcResult.tenure.totalDaysInclusive}일` : "-", "입사일 포함 총 재직일수"],
    [
      "이번산정주기",
      calcResult ? `${formatDate(calcResult.cycleStart)} ~ ${formatDate(calcResult.cycleEnd)}` : "-",
      calcResult ? calcResult.cycleLabel : "-"
    ],
    ["이번주기_발생연차", totalGenerated, calcResult ? calcResult.formulaDesc : "-"],
    ["이번주기_사용연차", Math.round(currentCycleUsed * 100) / 100, "이번 주기 범위 내 사용일수 합계"],
    ["이번주기_잔여연차", remainingLeave, "이번주기 발생연차 - 이번주기 사용연차"]
  ];

  const wsSummary = XLSX.utils.aoa_to_sheet(summaryAoA);
  wsSummary["!cols"] = [{ wch: 22 }, { wch: 28 }, { wch: 44 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsHistory, "연차사용이력");
  XLSX.utils.book_append_sheet(wb, wsSummary, "설정및연차요약");

  const fileName = buildExcelFileName();
  XLSX.writeFile(wb, fileName);
  showToast(`💾 '${fileName}' 파일이 PC에 저장되었습니다.`);
}

function normalizeExcelDateValue(val) {
  if (!val) return "";
  if (typeof val === "number") {
    // 엑셀 시리얼 날짜 변환
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

      // 1) '설정및연차요약' 시트가 있으면 설정 복원
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
        if (kv["수동보정일수"] !== undefined) {
          state.manualAdjust = parseFloat(kv["수동보정일수"]) || 0;
        }
      }

      // 2) '연차사용이력' 또는 첫 번째 시트에서 사용 기록 복원
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
        const type = String(row["휴가구분"] || row["구분"] || (daysNum === 0.5 ? "반차" : daysNum === 0.25 ? "반반차" : "연차")).trim();

        importedRecords.push({
          id: `imp_${Date.now()}_${idx}`,
          date: normDate,
          type,
          days: daysNum,
          note
        });
      });

      state.records = importedRecords;
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
// 5. 이벤트 리스너 바인딩 및 초기화
// ============================================================================

function initEvents() {
  // 기준일(오늘) 변경 및 리셋
  els.referenceDateInput.addEventListener("change", () => {
    renderAll();
  });

  els.resetTodayBtn.addEventListener("click", () => {
    els.referenceDateInput.value = getTodayStr();
    renderAll();
    showToast("기준일이 오늘 날짜로 설정되었습니다.");
  });

  // 기본 설정 변경 이벤트
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

  // 휴가 종류 선택 시 일수 자동 세팅
  els.leaveTypeSelect.addEventListener("change", (e) => {
    const [, defaultDays] = e.target.value.split("|");
    if (defaultDays !== "custom") {
      els.leaveDaysInput.value = parseFloat(defaultDays);
    } else {
      els.leaveDaysInput.focus();
    }
  });

  // 연차 사용 내역 추가 폼 제출
  els.leaveRecordForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const useDate = els.leaveUseDate.value;
    if (!useDate) {
      alert("사용일자를 선택해 주세요.");
      return;
    }

    const daysVal = parseFloat(els.leaveDaysInput.value);
    if (isNaN(daysVal) || daysVal <= 0) {
      alert("올바른 사용 일수를 입력해 주세요.");
      return;
    }

    const [typeName] = els.leaveTypeSelect.value.split("|");
    const noteText = els.leaveNoteInput.value.trim();

    const newRecord = {
      id: `rec_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      date: useDate,
      type: typeName,
      days: daysVal,
      note: noteText
    };

    state.records.push(newRecord);
    saveToLocalStorage();
    renderAll();

    els.leaveNoteInput.value = "";
    showToast(`✅ ${useDate} (${daysVal}일) 연차 사용 내역이 추가되었습니다.`);

    if (state.autoDownloadExcel) {
      exportToExcel();
    }
  });

  // 우측 테이블 개별 삭제 (이벤트 위임)
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

  // 우측 기록 필터 탭
  els.filterTabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      els.filterTabs.forEach((t) => t.classList.remove("active"));
      tab.classList.add("active");
      state.historyFilter = tab.getAttribute("data-filter") || "all";
      renderAll();
    });
  });

  // 전체 기록 초기화
  els.clearAllRecordsBtn.addEventListener("click", () => {
    if (state.records.length === 0) {
      showToast("삭제할 연차 사용 기록이 없습니다.");
      return;
    }
    if (confirm("등록된 모든 연차 사용 기록을 초기화하시겠습니까?\n(초기화 전 '엑셀 DB 저장하기'로 백업해 두시는 것을 권장합니다.)")) {
      state.records = [];
      saveToLocalStorage();
      renderAll();
      showToast("모든 연차 사용 기록이 초기화되었습니다.");
    }
  });

  // 엑셀 저장 / 불러오기 버튼
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
  els.leaveUseDate.value = todayStr;

  const hasSaved = loadFromLocalStorage();
  if (!hasSaved) {
    // 첫 방문 시 디폴트값: 입사년도 기준('anniversary'), 예시 입사일(오늘로부터 약 1년 6개월 전)을 세팅하지 않고
    // 바로 체험해볼 수 있도록 기본 입사일을 작년 3월 1일로 넣어주되 언제든 수정 가능하게 제공
    state.calcBasis = "anniversary";
  }

  initEvents();
  renderAll();
}

document.addEventListener("DOMContentLoaded", initApp);
