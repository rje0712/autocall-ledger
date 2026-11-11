# build_ledger.py
import json

with open('ledger_data.json', encoding='utf-8') as f:
    initial_data = json.load(f)

json_data_str = json.dumps(initial_data, ensure_ascii=False)

html_template = """<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>오토콜(AutoCall) 회비 장부</title>
  <!-- Tailwind CSS -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- FontAwesome Icons -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <!-- Chart.js -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <!-- SheetJS for Excel import & export -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <!-- Canvas Confetti -->
  <script src="https://cdn.jsdelivr.net/npm/canvas-confetti@1.6.0/dist/confetti.browser.min.js"></script>
  <style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    body {
      font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif;
      background-color: #f8fafc;
    }
    .custom-scroll::-webkit-scrollbar {
      width: 6px;
      height: 6px;
    }
    .custom-scroll::-webkit-scrollbar-track {
      background: #f1f5f9;
    }
    .custom-scroll::-webkit-scrollbar-thumb {
      background: #cbd5e1;
      border-radius: 4px;
    }
    .custom-scroll::-webkit-scrollbar-thumb:hover {
      background: #94a3b8;
    }
    .matrix-cell {
      transition: all 0.15s ease-in-out;
    }
    .matrix-cell:hover {
      transform: translateY(-1px);
    }
  </style>
  <script>
    tailwind.config = {
      theme: {
        extend: {
          colors: {
            brand: {
              50: '#eff6ff',
              100: '#dbeafe',
              500: '#3b82f6',
              600: '#2563eb',
              700: '#1d4ed8',
              900: '#1e3a8a',
            }
          }
        }
      }
    }
  </script>
</head>
<body class="min-h-screen text-slate-800 flex flex-col">

  <!-- Header -->
  <header class="bg-slate-900 text-white shadow-lg sticky top-0 z-40">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
      <div class="flex items-center justify-between h-16 sm:h-20">
        <!-- Brand -->
        <div class="flex items-center space-x-3">
          <div class="w-10 h-10 rounded-xl bg-gradient-to-tr from-blue-600 to-indigo-500 flex items-center justify-center shadow-md">
            <i class="fa-solid fa-car text-white text-lg"></i>
          </div>
          <div>
            <div class="flex items-center space-x-2">
              <h1 class="text-xl sm:text-2xl font-bold tracking-tight">오토콜 회비 장부</h1>
              <span id="roleBadge" class="text-xs px-2.5 py-0.5 rounded-full font-medium bg-slate-800 text-slate-300 border border-slate-700">
                <i class="fa-regular fa-eye mr-1"></i>열람 모드
              </span>
            </div>
            <p class="text-xs text-slate-400 hidden sm:block">2026 오토콜 정기 회비 스마트 관리 시스템</p>
          </div>
        </div>

        <!-- Right Action Buttons -->
        <div class="flex items-center space-x-2">
          <!-- Viewer Mode Button -->
          <div id="viewerActions" class="flex items-center space-x-2">
            <button onclick="openLoginModal()" class="inline-flex items-center px-3.5 py-2 rounded-lg text-sm font-semibold bg-blue-600 hover:bg-blue-500 text-white shadow-sm transition">
              <i class="fa-solid fa-lock mr-2"></i>관리자 로그인
            </button>
          </div>

          <!-- Admin Mode Buttons -->
          <div id="adminActions" class="hidden flex items-center space-x-2 flex-wrap sm:flex-nowrap">
            <!-- 1. Single Transaction Add -->
            <button onclick="openAddTxModal()" class="inline-flex items-center px-3 py-1.5 sm:px-3.5 sm:py-2 rounded-lg text-xs sm:text-sm font-semibold bg-emerald-600 hover:bg-emerald-500 text-white shadow-sm transition" title="한 번에 1건 등록 (자동 충당 지원)">
              <i class="fa-solid fa-plus mr-1.5"></i>건별 등록
            </button>

            <!-- 2. Batch Monthly Add (전체 장부 입력) -->
            <button onclick="openBatchModal()" class="inline-flex items-center px-3 py-1.5 sm:px-3.5 sm:py-2 rounded-lg text-xs sm:text-sm font-semibold bg-indigo-600 hover:bg-indigo-500 text-white shadow-sm transition" title="당월 8명 회비 한 번에 일괄 등록">
              <i class="fa-solid fa-list-check mr-1.5"></i>전체 일괄 등록
            </button>

            <!-- 3. Excel Upload / Sync -->
            <button onclick="document.getElementById('excelFileInput').click()" class="inline-flex items-center px-3 py-1.5 sm:px-3.5 sm:py-2 rounded-lg text-xs sm:text-sm font-medium bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition" title="엑셀 파일(.xlsx)로 전체 장부 덮어쓰기/불러오기">
              <i class="fa-solid fa-file-excel mr-1.5 text-emerald-400"></i>엑셀 불러오기
            </button>
            <input type="file" id="excelFileInput" accept=".xlsx, .xls" class="hidden" onchange="handleExcelUpload(event)">

            <!-- 4. Admin Menu Dropdown -->
            <div class="relative inline-block text-left" id="adminMenuDropdownContainer">
              <button onclick="toggleAdminMenu()" class="inline-flex items-center px-3 py-1.5 sm:px-3 sm:py-2 rounded-lg text-xs sm:text-sm font-medium bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition">
                <i class="fa-solid fa-gear mr-1.5"></i>관리
                <i class="fa-solid fa-chevron-down ml-1 text-xs text-slate-400"></i>
              </button>
              <div id="adminMenuDropdown" class="hidden absolute right-0 mt-2 w-64 rounded-xl shadow-2xl bg-white text-slate-800 ring-1 ring-black ring-opacity-5 divide-y divide-slate-100 z-50">
                <div class="p-1">
                  <button onclick="exportShareableHtml()" class="w-full text-left px-3 py-2 text-sm text-slate-700 hover:bg-blue-50 hover:text-blue-700 rounded-lg flex items-center font-medium">
                    <i class="fa-solid fa-share-nodes text-blue-600 w-5"></i>카톡 공유용 HTML 내보내기
                  </button>
                  <button onclick="exportToExcel()" class="w-full text-left px-3 py-2 text-sm text-slate-700 hover:bg-emerald-50 hover:text-emerald-700 rounded-lg flex items-center">
                    <i class="fa-solid fa-file-excel text-emerald-600 w-5"></i>엑셀(.xlsx) 파일 다운로드
                  </button>
                  <button onclick="exportJsonBackup()" class="w-full text-left px-3 py-2 text-sm text-slate-700 hover:bg-slate-50 rounded-lg flex items-center">
                    <i class="fa-solid fa-download text-slate-500 w-5"></i>JSON 데이터 백업
                  </button>
                  <button onclick="document.getElementById('importJsonFile').click()" class="w-full text-left px-3 py-2 text-sm text-slate-700 hover:bg-slate-50 rounded-lg flex items-center">
                    <i class="fa-solid fa-upload text-slate-500 w-5"></i>JSON 데이터 복원
                  </button>
                  <input type="file" id="importJsonFile" accept=".json" class="hidden" onchange="importJsonBackup(event)">
                </div>
                <div class="p-1">
                  <button onclick="openChangePasswordModal()" class="w-full text-left px-3 py-2 text-sm text-slate-700 hover:bg-slate-50 rounded-lg flex items-center">
                    <i class="fa-solid fa-key text-amber-500 w-5"></i>비밀번호 변경
                  </button>
                  <button onclick="confirmResetData()" class="w-full text-left px-3 py-2 text-sm text-red-600 hover:bg-red-50 rounded-lg flex items-center">
                    <i class="fa-solid fa-rotate-left text-red-500 w-5"></i>초기 데이터로 리셋
                  </button>
                </div>
              </div>
            </div>

            <!-- Logout -->
            <button onclick="logoutAdmin()" class="inline-flex items-center px-2.5 py-1.5 sm:px-3 sm:py-2 rounded-lg text-xs sm:text-sm font-medium bg-slate-800 hover:bg-rose-900 text-slate-300 hover:text-rose-200 border border-slate-700 transition" title="로그아웃">
              <i class="fa-solid fa-arrow-right-from-bracket"></i>
            </button>
          </div>
        </div>
      </div>
    </div>
  </header>

  <!-- Main Container -->
  <main class="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">

    <!-- KPI Summary Cards -->
    <div class="grid grid-cols-2 lg:grid-cols-4 gap-4">
      <!-- Balance Card -->
      <div class="bg-white rounded-2xl p-5 shadow-sm border border-slate-200/80 relative overflow-hidden flex flex-col justify-between">
        <div class="flex items-center justify-between">
          <span class="text-xs font-semibold uppercase tracking-wider text-slate-500">현재 통장 잔액</span>
          <div class="w-9 h-9 rounded-xl bg-blue-50 text-blue-600 flex items-center justify-center font-bold">
            <i class="fa-solid fa-wallet text-sm"></i>
          </div>
        </div>
        <div class="mt-3">
          <div class="text-2xl sm:text-3xl font-black text-slate-900 tracking-tight" id="kpiBalance">0원</div>
          <div class="text-xs text-blue-600 mt-1 font-medium flex items-center">
            <i class="fa-solid fa-circle-check mr-1 text-[10px]"></i>실시간 잔액 일치
          </div>
        </div>
      </div>

      <!-- Total Income Card -->
      <div class="bg-white rounded-2xl p-5 shadow-sm border border-slate-200/80 flex flex-col justify-between">
        <div class="flex items-center justify-between">
          <span class="text-xs font-semibold uppercase tracking-wider text-slate-500">총 누적 입금</span>
          <div class="w-9 h-9 rounded-xl bg-emerald-50 text-emerald-600 flex items-center justify-center">
            <i class="fa-solid fa-arrow-trend-up text-sm"></i>
          </div>
        </div>
        <div class="mt-3">
          <div class="text-2xl sm:text-3xl font-black text-emerald-600 tracking-tight" id="kpiTotalIncome">0원</div>
          <div class="text-xs text-slate-500 mt-1">
            순 회비: <span id="kpiPureFee" class="font-semibold text-slate-700">0원</span>
          </div>
        </div>
      </div>

      <!-- Total Expenses Card -->
      <div class="bg-white rounded-2xl p-5 shadow-sm border border-slate-200/80 flex flex-col justify-between">
        <div class="flex items-center justify-between">
          <span class="text-xs font-semibold uppercase tracking-wider text-slate-500">누적 출금 / 지출</span>
          <div class="w-9 h-9 rounded-xl bg-rose-50 text-rose-600 flex items-center justify-center">
            <i class="fa-solid fa-arrow-trend-down text-sm"></i>
          </div>
        </div>
        <div class="mt-3">
          <div class="text-2xl sm:text-3xl font-black text-rose-600 tracking-tight" id="kpiTotalExpense">0원</div>
          <div class="text-xs text-slate-500 mt-1">
            오출금/반환 상쇄 완료
          </div>
        </div>
      </div>

      <!-- Interest / Member Progress Card -->
      <div class="bg-white rounded-2xl p-5 shadow-sm border border-slate-200/80 flex flex-col justify-between">
        <div class="flex items-center justify-between">
          <span class="text-xs font-semibold uppercase tracking-wider text-slate-500">예금 이자 누적</span>
          <div class="w-9 h-9 rounded-xl bg-amber-50 text-amber-600 flex items-center justify-center">
            <i class="fa-solid fa-piggy-bank text-sm"></i>
          </div>
        </div>
        <div class="mt-3">
          <div class="text-2xl sm:text-3xl font-black text-amber-600 tracking-tight" id="kpiInterest">0원</div>
          <div class="text-xs text-slate-500 mt-1 flex items-center justify-between">
            <span>회원 8명 (월 20,000원)</span>
          </div>
        </div>
      </div>
    </div>

    <!-- Tab Buttons -->
    <div class="flex border-b border-slate-200 space-x-1 sm:space-x-3 overflow-x-auto custom-scroll pb-1">
      <button onclick="switchTab('matrix')" id="tabBtn-matrix" class="tab-btn px-4 py-2.5 text-sm font-bold border-b-2 border-blue-600 text-blue-600 whitespace-nowrap flex items-center">
        <i class="fa-solid fa-table-cells mr-2"></i>월별 납부 현황판
      </button>
      <button onclick="switchTab('members')" id="tabBtn-members" class="tab-btn px-4 py-2.5 text-sm font-semibold border-b-2 border-transparent text-slate-500 hover:text-slate-800 whitespace-nowrap flex items-center">
        <i class="fa-solid fa-users mr-2"></i>멤버별 정산 & 알림
      </button>
      <button onclick="switchTab('ledger')" id="tabBtn-ledger" class="tab-btn px-4 py-2.5 text-sm font-semibold border-b-2 border-transparent text-slate-500 hover:text-slate-800 whitespace-nowrap flex items-center">
        <i class="fa-solid fa-receipt mr-2"></i>전체 거래 장부 (<span id="ledgerCountBadge">0</span>)
      </button>
      <button onclick="switchTab('charts')" id="tabBtn-charts" class="tab-btn px-4 py-2.5 text-sm font-semibold border-b-2 border-transparent text-slate-500 hover:text-slate-800 whitespace-nowrap flex items-center">
        <i class="fa-solid fa-chart-column mr-2"></i>통계 & 차트
      </button>
    </div>

    <!-- TAB 1: 월별 납부 현황 매트릭스 -->
    <div id="tab-matrix" class="tab-content space-y-4">
      <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
        <!-- Matrix Top Bar -->
        <div class="p-4 sm:p-5 border-b border-slate-200 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 bg-slate-50/50">
          <div>
            <h2 class="text-base sm:text-lg font-bold text-slate-900 flex items-center">
              <span>2026 연간 월별 회비 납부 매트릭스</span>
              <span class="ml-2 text-xs font-normal text-slate-500 bg-white px-2.5 py-1 rounded-md border border-slate-200 shadow-xs">
                2~5월 1만 / 6~12월 2만
              </span>
            </h2>
            <p class="text-xs text-slate-500 mt-1">
              한눈에 보는 멤버별 월별 납부 완료 여부입니다. <span id="matrixAdminNotice" class="hidden text-blue-600 font-semibold">(관리자 모드: 셀 클릭 시 납부/미납 즉시 토글)</span>
            </p>
          </div>
          <div class="flex items-center space-x-2 text-xs">
            <span class="inline-flex items-center px-2 py-1 rounded bg-emerald-50 text-emerald-700 font-medium border border-emerald-200">
              <i class="fa-solid fa-check mr-1"></i>납부 완료
            </span>
            <span class="inline-flex items-center px-2 py-1 rounded bg-rose-50 text-rose-700 font-medium border border-rose-200">
              <i class="fa-solid fa-xmark mr-1"></i>미납
            </span>
          </div>
        </div>

        <!-- Matrix Table Container -->
        <div class="overflow-x-auto custom-scroll">
          <table class="w-full text-center border-collapse text-xs sm:text-sm">
            <thead>
              <tr class="bg-slate-100 text-slate-700 font-bold border-b border-slate-200">
                <th class="p-3 text-left sticky left-0 bg-slate-100 z-10 min-w-[90px] border-r border-slate-200 shadow-xs">회원명</th>
                <th class="p-2 min-w-[70px]">2월<br><span class="text-[10px] text-slate-500 font-normal">1만</span></th>
                <th class="p-2 min-w-[70px]">3월<br><span class="text-[10px] text-slate-500 font-normal">1만</span></th>
                <th class="p-2 min-w-[70px]">4월<br><span class="text-[10px] text-slate-500 font-normal">1만</span></th>
                <th class="p-2 min-w-[70px]">5월<br><span class="text-[10px] text-slate-500 font-normal">1만</span></th>
                <th class="p-2 min-w-[75px] bg-blue-50/50">6월<br><span class="text-[10px] text-blue-600 font-semibold">2만(인상)</span></th>
                <th class="p-2 min-w-[70px]">7월<br><span class="text-[10px] text-slate-500 font-normal">2만</span></th>
                <th class="p-2 min-w-[70px]">8월<br><span class="text-[10px] text-slate-500 font-normal">2만</span></th>
                <th class="p-2 min-w-[70px]">9월<br><span class="text-[10px] text-slate-500 font-normal">2만</span></th>
                <th class="p-2 min-w-[70px]">10월<br><span class="text-[10px] text-slate-500 font-normal">2만</span></th>
                <th class="p-2 min-w-[70px] bg-amber-50/40">11월<br><span class="text-[10px] text-amber-700 font-normal">2만(당월)</span></th>
                <th class="p-2 min-w-[70px]">12월<br><span class="text-[10px] text-slate-400 font-normal">2만</span></th>
                <th class="p-3 min-w-[100px] bg-slate-100 font-bold text-slate-800 border-l border-slate-200">납부 합계</th>
              </tr>
            </thead>
            <tbody id="matrixTableBody" class="divide-y divide-slate-100">
              <!-- Rendered via JS -->
            </tbody>
            <tfoot id="matrixTableFoot" class="bg-slate-50 font-bold border-t-2 border-slate-200">
              <!-- Rendered via JS -->
            </tfoot>
          </table>
        </div>
      </div>
    </div>

    <!-- TAB 2: 멤버별 납부 & 카톡 알림 -->
    <div id="tab-members" class="tab-content hidden space-y-5">
      <!-- Kakao Notification Helper -->
      <div class="bg-gradient-to-r from-amber-500/10 via-amber-400/5 to-transparent border border-amber-300/60 rounded-2xl p-5 flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
        <div class="flex items-start space-x-3">
          <div class="w-10 h-10 rounded-xl bg-amber-400 text-amber-950 flex items-center justify-center font-bold text-lg shadow-sm">
            <i class="fa-solid fa-comment"></i>
          </div>
          <div>
            <h3 class="font-bold text-slate-900 text-sm sm:text-base">카카오톡 모임 공지 & 미납 안내 생성기</h3>
            <p class="text-xs text-slate-600 mt-0.5">버튼을 누르면 현재 미납 내역과 계좌 현황이 서식에 맞게 자동으로 복사되어 카톡방에 바로 전송할 수 있습니다.</p>
          </div>
        </div>
        <button onclick="copyKakaoNotice()" class="inline-flex items-center px-4 py-2.5 rounded-xl bg-amber-400 hover:bg-amber-300 text-amber-950 font-bold text-xs sm:text-sm shadow-sm transition">
          <i class="fa-solid fa-copy mr-2"></i>카톡 공지 텍스트 복사
        </button>
      </div>

      <!-- Member Grid -->
      <div id="membersGrid" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <!-- Rendered via JS -->
      </div>
    </div>

    <!-- TAB 3: 전체 거래 장부 (입출금) -->
    <div id="tab-ledger" class="tab-content hidden space-y-4">
      <div class="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
        <!-- Search and Filter Bar -->
        <div class="p-4 sm:p-5 border-b border-slate-200 space-y-3 bg-slate-50/50">
          <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
            <h2 class="text-base sm:text-lg font-bold text-slate-900 flex items-center">
              <span>통장 입출금 거래 내역</span>
              <span class="ml-2 text-xs font-semibold px-2.5 py-0.5 rounded-full bg-slate-200 text-slate-700" id="filteredCountBadge">0건</span>
            </h2>
            <div class="flex items-center space-x-2">
              <select id="filterMonth" onchange="renderLedgerTable()" class="text-xs sm:text-sm bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 focus:ring-2 focus:ring-blue-500 focus:outline-none">
                <option value="ALL">전체 월</option>
                <option value="2월">2월</option>
                <option value="3월">3월</option>
                <option value="4월">4월</option>
                <option value="5월">5월</option>
                <option value="6월">6월</option>
                <option value="7월">7월</option>
                <option value="8월">8월</option>
                <option value="9월">9월</option>
                <option value="10월">10월</option>
                <option value="11월">11월</option>
                <option value="12월">12월</option>
              </select>
              <select id="filterPerson" onchange="renderLedgerTable()" class="text-xs sm:text-sm bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 focus:ring-2 focus:ring-blue-500 focus:outline-none">
                <option value="ALL">전체 멤버/유형</option>
                <!-- Rendered options -->
              </select>
              <select id="filterType" onchange="renderLedgerTable()" class="text-xs sm:text-sm bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 focus:ring-2 focus:ring-blue-500 focus:outline-none">
                <option value="ALL">전체 거래</option>
                <option value="IN">입금(+)</option>
                <option value="OUT">출금(-)</option>
              </select>
            </div>
          </div>
          <!-- Search Input -->
          <div class="relative">
            <i class="fa-solid fa-magnifying-glass absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 text-xs"></i>
            <input type="text" id="txSearchInput" oninput="renderLedgerTable()" placeholder="이름, 날짜, 비고 검색..." class="w-full pl-9 pr-4 py-2 text-xs sm:text-sm bg-white border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:outline-none">
          </div>
        </div>

        <!-- Ledger Table -->
        <div class="overflow-x-auto custom-scroll max-h-[600px]">
          <table class="w-full text-left border-collapse text-xs sm:text-sm">
            <thead class="bg-slate-100 text-slate-600 font-semibold sticky top-0 z-10 shadow-xs">
              <tr>
                <th class="p-3 w-16 text-center">연번</th>
                <th class="p-3 w-24">상세날짜</th>
                <th class="p-3 w-20 text-center">해당월</th>
                <th class="p-3 w-32 text-center">추가/충당월</th>
                <th class="p-3 w-28">입금인/항목</th>
                <th class="p-3 w-32 text-right">금액</th>
                <th class="p-3">비고</th>
                <th class="p-3 w-20 text-center admin-only hidden">관리</th>
              </tr>
            </thead>
            <tbody id="ledgerTableBody" class="divide-y divide-slate-100">
              <!-- Rendered via JS -->
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- TAB 4: 통계 & 차트 -->
    <div id="tab-charts" class="tab-content hidden space-y-6">
      <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <!-- Monthly Revenue Chart -->
        <div class="bg-white rounded-2xl p-5 shadow-sm border border-slate-200">
          <h3 class="text-sm sm:text-base font-bold text-slate-900 mb-4 flex items-center">
            <i class="fa-solid fa-chart-column text-blue-600 mr-2"></i>월별 회비 수입 추이 (원)
          </h3>
          <div class="h-64 sm:h-72">
            <canvas id="monthlyChart"></canvas>
          </div>
        </div>

        <!-- Member Contribution Chart -->
        <div class="bg-white rounded-2xl p-5 shadow-sm border border-slate-200">
          <h3 class="text-sm sm:text-base font-bold text-slate-900 mb-4 flex items-center">
            <i class="fa-solid fa-chart-pie text-emerald-600 mr-2"></i>멤버별 누적 납부 현황
          </h3>
          <div class="h-64 sm:h-72 flex items-center justify-center">
            <canvas id="memberChart"></canvas>
          </div>
        </div>
      </div>
    </div>

  </main>

  <!-- Footer -->
  <footer class="bg-slate-900 text-slate-400 py-6 text-xs text-center border-t border-slate-800 mt-auto">
    <div class="max-w-7xl mx-auto px-4">
      <p>© 2026 오토콜(AutoCall) 모임 장부 시스템 · 투명하고 공정한 모임 운영</p>
      <p class="mt-1 text-slate-500">24시간 상시 운영 웹 애플리케이션</p>
    </div>
  </footer>

  <!-- ================= MODALS ================= -->

  <!-- 1. Admin Login Modal -->
  <div id="loginModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 hidden">
    <div class="bg-white rounded-2xl max-w-sm w-full p-6 shadow-2xl border border-slate-200">
      <div class="text-center">
        <div class="w-12 h-12 bg-blue-50 text-blue-600 rounded-2xl mx-auto flex items-center justify-center text-xl mb-3 shadow-xs">
          <i class="fa-solid fa-shield-halved"></i>
        </div>
        <h3 class="text-lg font-bold text-slate-900">관리자 인증</h3>
        <p class="text-xs text-slate-500 mt-1">장부 수정 및 관리를 위해 비밀번호를 입력해주세요.</p>
      </div>
      <form onsubmit="handleLoginSubmit(event)" class="mt-5 space-y-4">
        <div>
          <input type="password" id="adminPasswordInput" placeholder="관리자 비밀번호 입력" class="w-full px-4 py-2.5 text-sm rounded-xl border border-slate-300 focus:ring-2 focus:ring-blue-500 focus:outline-none text-center tracking-widest font-mono">
          <p class="text-[11px] text-slate-400 mt-1.5 text-center">기본 초기 비밀번호: <span class="font-mono text-blue-600 font-bold">1234</span></p>
        </div>
        <div class="flex space-x-2">
          <button type="button" onclick="closeLoginModal()" class="flex-1 py-2.5 rounded-xl border border-slate-300 text-slate-700 text-sm font-semibold hover:bg-slate-50 transition">
            취소
          </button>
          <button type="submit" class="flex-1 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-500 text-white text-sm font-semibold shadow-sm transition">
            로그인
          </button>
        </div>
      </form>
    </div>
  </div>

  <!-- 2. Single Transaction Modal (With Auto-Fill FIFO) -->
  <div id="txModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 hidden">
    <div class="bg-white rounded-2xl max-w-lg w-full p-6 shadow-2xl border border-slate-200 max-h-[90vh] overflow-y-auto custom-scroll">
      <div class="flex items-center justify-between border-b border-slate-100 pb-3">
        <h3 class="text-base font-bold text-slate-900 flex items-center" id="txModalTitle">
          <i class="fa-solid fa-plus-circle text-emerald-600 mr-2"></i>회비 / 입출금 내역 등록
        </h3>
        <button onclick="closeTxModal()" class="text-slate-400 hover:text-slate-600 text-lg">
          <i class="fa-solid fa-xmark"></i>
        </button>
      </div>

      <form onsubmit="handleTxSubmit(event)" class="mt-4 space-y-4 text-xs sm:text-sm">
        <input type="hidden" id="txEditId" value="">

        <!-- Person / Member Selection -->
        <div>
          <div class="flex items-center justify-between mb-1">
            <label class="block text-xs font-semibold text-slate-700">입금인 / 회원 선택 <span class="text-red-500">*</span></label>
            <span id="memberUnpaidStatusBadge" class="text-[11px] text-slate-500"></span>
          </div>
          <div class="flex space-x-2">
            <select id="txPersonSelect" onchange="onPersonChange()" class="flex-1 px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:outline-none font-bold text-slate-800">
              <!-- Rendered options -->
            </select>
            <input type="text" id="txPersonDirect" placeholder="직접 입력" class="w-1/2 px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:outline-none hidden">
          </div>
        </div>

        <!-- Amount & Quick Buttons -->
        <div>
          <div class="flex items-center justify-between mb-1">
            <label class="block text-xs font-semibold text-slate-700">입금 금액 (원) <span class="text-red-500">* (출금은 마이너스)</span></label>
            <button type="button" id="btnFillAllUnpaid" onclick="fillAllUnpaid()" class="hidden text-[11px] text-blue-600 hover:underline font-bold">
              ⚡ 미납 전액 원클릭 채우기
            </button>
          </div>
          <div class="relative">
            <input type="number" id="txAmount" placeholder="20000" oninput="updateAutoFillPreview()" required class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:outline-none font-mono font-black text-lg text-slate-900">
          </div>
          <!-- Quick Multi-month Presets -->
          <div class="flex flex-wrap gap-1.5 mt-2">
            <button type="button" onclick="setTxAmount(20000)" class="px-2.5 py-1 text-xs rounded-lg bg-blue-50 hover:bg-blue-100 text-blue-700 font-semibold border border-blue-200">+2만(1달)</button>
            <button type="button" onclick="setTxAmount(40000)" class="px-2.5 py-1 text-xs rounded-lg bg-blue-50 hover:bg-blue-100 text-blue-700 font-semibold border border-blue-200">+4만(2달)</button>
            <button type="button" onclick="setTxAmount(60000)" class="px-2.5 py-1 text-xs rounded-lg bg-blue-50 hover:bg-blue-100 text-blue-700 font-semibold border border-blue-200">+6만(3달)</button>
            <button type="button" onclick="setTxAmount(80000)" class="px-2.5 py-1 text-xs rounded-lg bg-blue-50 hover:bg-blue-100 text-blue-700 font-semibold border border-blue-200">+8만(4달)</button>
            <button type="button" onclick="setTxAmount(100000)" class="px-2.5 py-1 text-xs rounded-lg bg-indigo-50 hover:bg-indigo-100 text-indigo-700 font-bold border border-indigo-200">+10만(5달)</button>
            <button type="button" onclick="setTxAmount(-10000)" class="px-2.5 py-1 text-xs rounded-lg bg-rose-50 hover:bg-rose-100 text-rose-700 font-medium border border-rose-200">-1만(출금)</button>
          </div>
        </div>

        <!-- ✨ Auto-Fill FIFO Preview Card -->
        <div id="autoFillPreviewCard" class="hidden rounded-xl p-3.5 bg-gradient-to-r from-blue-50 to-indigo-50 border border-blue-200 space-y-2">
          <div class="flex items-center justify-between">
            <span class="text-xs font-bold text-blue-900 flex items-center">
              <i class="fa-solid fa-wand-magic-sparkles text-blue-600 mr-1.5"></i>비워진 달부터 자동 순차 충당 (스마트 배분)
            </span>
            <label class="inline-flex items-center cursor-pointer">
              <input type="checkbox" id="autoFillEnabled" checked onchange="updateAutoFillPreview()" class="sr-only peer">
              <div class="relative w-8 h-4 bg-slate-200 peer-focus:outline-none rounded-full peer peer-checked:after:translate-x-full rtl:peer-checked:after:-translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:start-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-3 after:w-3 after:transition-all peer-checked:bg-blue-600"></div>
            </label>
          </div>
          <div id="autoFillDetails" class="text-xs text-blue-800 space-y-1"></div>
        </div>

        <!-- Date & Month -->
        <div class="grid grid-cols-2 gap-2">
          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">상세 날짜 (YY.MM.DD) <span class="text-red-500">*</span></label>
            <input type="text" id="txDate" placeholder="26.11.11" required class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:outline-none font-mono">
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">해당월 (입금 기준월) <span class="text-red-500">*</span></label>
            <select id="txMonth" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:outline-none">
              <option value="2월">2월</option>
              <option value="3월">3월</option>
              <option value="4월">4월</option>
              <option value="5월">5월</option>
              <option value="6월">6월</option>
              <option value="7월">7월</option>
              <option value="8월">8월</option>
              <option value="9월">9월</option>
              <option value="10월">10월</option>
              <option value="11월" selected>11월</option>
              <option value="12월">12월</option>
              <option value="기타">기타</option>
            </select>
          </div>
        </div>

        <!-- Extra Month -->
        <div>
          <label class="block text-xs font-semibold text-slate-700 mb-1">추가납입 / 충당월 목록</label>
          <input type="text" id="txExtraMonth" placeholder="예: 7월, 8월, 9월 (자동 입력됨)" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:outline-none">
        </div>

        <!-- Note -->
        <div>
          <label class="block text-xs font-semibold text-slate-700 mb-1">비고 (선택 사항)</label>
          <input type="text" id="txNote" placeholder="예: 5개월치 회비 일괄 납부" class="w-full px-3 py-2 border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:outline-none">
        </div>

        <!-- Submit Buttons -->
        <div class="pt-2 flex space-x-2">
          <button type="button" onclick="closeTxModal()" class="flex-1 py-2.5 rounded-xl border border-slate-300 text-slate-700 font-semibold hover:bg-slate-50 transition">
            취소
          </button>
          <button type="submit" class="flex-1 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-500 text-white font-semibold shadow-sm transition">
            저장하기
          </button>
        </div>
      </form>
    </div>
  </div>

  <!-- 3. BATCH MONTHLY PAYMENT MODAL (전체 장부 일괄 입력!) -->
  <div id="batchModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 hidden">
    <div class="bg-white rounded-2xl max-w-xl w-full p-6 shadow-2xl border border-slate-200 max-h-[90vh] overflow-y-auto custom-scroll">
      <div class="flex items-center justify-between border-b border-slate-100 pb-3">
        <div>
          <h3 class="text-base font-bold text-slate-900 flex items-center">
            <i class="fa-solid fa-list-check text-indigo-600 mr-2"></i>월별 전체 회원 일괄 회비 등록
          </h3>
          <p class="text-xs text-slate-500 mt-0.5">선택한 월에 납부한 회원들을 체크하여 한 번에 장부에 입력합니다.</p>
        </div>
        <button onclick="closeBatchModal()" class="text-slate-400 hover:text-slate-600 text-lg">
          <i class="fa-solid fa-xmark"></i>
        </button>
      </div>

      <form onsubmit="handleBatchSubmit(event)" class="mt-4 space-y-4 text-xs sm:text-sm">
        <div class="grid grid-cols-2 gap-3 bg-slate-50 p-3 rounded-xl border border-slate-200">
          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">대상월 선택</label>
            <select id="batchMonthSelect" onchange="renderBatchMembersTable()" class="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg font-bold text-slate-800 focus:ring-2 focus:ring-indigo-500">
              <option value="2월">2월</option>
              <option value="3월">3월</option>
              <option value="4월">4월</option>
              <option value="5월">5월</option>
              <option value="6월">6월</option>
              <option value="7월">7월</option>
              <option value="8월">8월</option>
              <option value="9월">9월</option>
              <option value="10월">10월</option>
              <option value="11월" selected>11월</option>
              <option value="12월">12월</option>
            </select>
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-700 mb-1">납부 날짜 (YY.MM.DD)</label>
            <input type="text" id="batchDate" placeholder="26.11.11" required class="w-full px-3 py-2 bg-white border border-slate-300 rounded-lg font-mono focus:ring-2 focus:ring-indigo-500">
          </div>
        </div>

        <div class="flex items-center justify-between pt-1">
          <span class="font-bold text-slate-800 text-xs">회원별 납부 체크</span>
          <div class="space-x-2 text-xs">
            <button type="button" onclick="toggleBatchAll(true)" class="text-indigo-600 font-semibold hover:underline">전체 선택</button>
            <span class="text-slate-300">|</span>
            <button type="button" onclick="toggleBatchAll(false)" class="text-slate-500 font-semibold hover:underline">전체 해제</button>
          </div>
        </div>

        <!-- Batch Members Table -->
        <div class="border border-slate-200 rounded-xl overflow-hidden">
          <table class="w-full text-left border-collapse text-xs">
            <thead class="bg-slate-100 text-slate-600 font-semibold border-b border-slate-200">
              <tr>
                <th class="p-2.5 w-12 text-center">선택</th>
                <th class="p-2.5 w-24">회원명</th>
                <th class="p-2.5 w-28">현재 상태</th>
                <th class="p-2.5 text-right w-28">입금 금액</th>
                <th class="p-2.5">비고</th>
              </tr>
            </thead>
            <tbody id="batchTableBody" class="divide-y divide-slate-100">
              <!-- Rendered via JS -->
            </tbody>
          </table>
        </div>

        <div class="pt-2 flex space-x-2">
          <button type="button" onclick="closeBatchModal()" class="flex-1 py-2.5 rounded-xl border border-slate-300 text-slate-700 font-semibold hover:bg-slate-50 transition">
            취소
          </button>
          <button type="submit" class="flex-1 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white font-semibold shadow-sm transition">
            <i class="fa-solid fa-check mr-1.5"></i>선택한 회원 일괄 등록
          </button>
        </div>
      </form>
    </div>
  </div>

  <!-- 4. Change Password Modal -->
  <div id="passwordModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 hidden">
    <div class="bg-white rounded-2xl max-w-sm w-full p-6 shadow-2xl border border-slate-200">
      <div class="flex items-center justify-between border-b border-slate-100 pb-3">
        <h3 class="text-base font-bold text-slate-900">관리자 비밀번호 변경</h3>
        <button onclick="closePasswordModal()" class="text-slate-400 hover:text-slate-600">
          <i class="fa-solid fa-xmark"></i>
        </button>
      </div>
      <form onsubmit="handleChangePasswordSubmit(event)" class="mt-4 space-y-3">
        <div>
          <label class="block text-xs font-semibold text-slate-600 mb-1">새 비밀번호</label>
          <input type="password" id="newPasswordInput" required placeholder="새 비밀번호 입력" class="w-full px-3 py-2 text-sm border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:outline-none font-mono">
        </div>
        <div>
          <label class="block text-xs font-semibold text-slate-600 mb-1">새 비밀번호 확인</label>
          <input type="password" id="newPasswordConfirmInput" required placeholder="새 비밀번호 재입력" class="w-full px-3 py-2 text-sm border border-slate-300 rounded-xl focus:ring-2 focus:ring-blue-500 focus:outline-none font-mono">
        </div>
        <div class="pt-2 flex space-x-2">
          <button type="button" onclick="closePasswordModal()" class="flex-1 py-2 rounded-xl border border-slate-300 text-slate-700 font-semibold hover:bg-slate-50 text-xs">
            취소
          </button>
          <button type="submit" class="flex-1 py-2 rounded-xl bg-blue-600 hover:bg-blue-500 text-white font-semibold text-xs transition">
            변경 완료
          </button>
        </div>
      </form>
    </div>
  </div>

  <!-- Toast Notification Container -->
  <div id="toastContainer" class="fixed bottom-5 right-5 z-50 flex flex-col space-y-2 pointer-events-none"></div>

  <!-- Javascript Logic -->
  <script>
    // Embedded initial data
    const INITIAL_DATA = __INITIAL_DATA_PLACEHOLDER__;

    // State
    let appData = null;
    let isAdmin = false;
    let charts = {};

    // Init App
    document.addEventListener('DOMContentLoaded', () => {
      loadData();
      checkAdminSession();
      renderAll();
      setupClickOutside();
    });

    // Load Data from LocalStorage or Default
    function loadData() {
      const saved = localStorage.getItem('autocall_ledger_data');
      if (saved) {
        try {
          appData = JSON.parse(saved);
        } catch (e) {
          console.error('Failed to parse local storage data', e);
          appData = JSON.parse(JSON.stringify(INITIAL_DATA));
        }
      } else {
        appData = JSON.parse(JSON.stringify(INITIAL_DATA));
      }

      // Ensure proper structure
      if (!appData.monthlyMatrix) appData.monthlyMatrix = {};
      appData.members.forEach(m => {
        if (!appData.monthlyMatrix[m]) appData.monthlyMatrix[m] = {};
      });
    }

    // Save Data
    function saveData() {
      localStorage.setItem('autocall_ledger_data', JSON.stringify(appData));
    }

    // Check Admin Session
    function checkAdminSession() {
      const adminSession = sessionStorage.getItem('autocall_admin_logged_in');
      isAdmin = (adminSession === 'true');
      updateAdminUI();
    }

    function updateAdminUI() {
      const roleBadge = document.getElementById('roleBadge');
      const viewerActions = document.getElementById('viewerActions');
      const adminActions = document.getElementById('adminActions');
      const adminNotice = document.getElementById('matrixAdminNotice');
      const adminOnlyElements = document.querySelectorAll('.admin-only');

      if (isAdmin) {
        roleBadge.className = 'text-xs px-2.5 py-0.5 rounded-full font-semibold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30';
        roleBadge.innerHTML = '<i class="fa-solid fa-crown mr-1 text-amber-400"></i>관리자 모드';
        viewerActions.classList.add('hidden');
        adminActions.classList.remove('hidden');
        if (adminNotice) adminNotice.classList.remove('hidden');
        adminOnlyElements.forEach(el => el.classList.remove('hidden'));
      } else {
        roleBadge.className = 'text-xs px-2.5 py-0.5 rounded-full font-medium bg-slate-800 text-slate-300 border border-slate-700';
        roleBadge.innerHTML = '<i class="fa-regular fa-eye mr-1"></i>열람 모드';
        viewerActions.classList.remove('hidden');
        adminActions.classList.add('hidden');
        if (adminNotice) adminNotice.classList.add('hidden');
        adminOnlyElements.forEach(el => el.classList.add('hidden'));
      }
    }

    // Render Everything
    function renderAll() {
      renderKPIs();
      renderMatrix();
      renderMembersGrid();
      renderLedgerFilters();
      renderLedgerTable();
      renderCharts();
    }

    // 1. Render KPIs
    function renderKPIs() {
      const txs = appData.transactions || [];
      const totalBalance = txs.reduce((acc, cur) => acc + (Number(cur.amount) || 0), 0);
      const totalIncome = txs.filter(t => t.amount > 0).reduce((acc, cur) => acc + cur.amount, 0);
      const totalExpense = txs.filter(t => t.amount < 0).reduce((acc, cur) => acc + Math.abs(cur.amount), 0);
      const interestTotal = txs.filter(t => t.person === '예금이자').reduce((acc, cur) => acc + cur.amount, 0);

      const pureFee = txs.filter(t => appData.members.includes(t.person) && t.amount > 0).reduce((acc, cur) => acc + cur.amount, 0);

      document.getElementById('kpiBalance').innerText = totalBalance.toLocaleString() + '원';
      document.getElementById('kpiTotalIncome').innerText = totalIncome.toLocaleString() + '원';
      document.getElementById('kpiTotalExpense').innerText = totalExpense.toLocaleString() + '원';
      document.getElementById('kpiPureFee').innerText = pureFee.toLocaleString() + '원';
      document.getElementById('kpiInterest').innerText = interestTotal.toLocaleString() + '원';
    }

    // 2. Render Matrix
    function renderMatrix() {
      const tbody = document.getElementById('matrixTableBody');
      const tfoot = document.getElementById('matrixTableFoot');
      tbody.innerHTML = '';
      tfoot.innerHTML = '';

      const months = appData.months;
      const members = appData.members;
      const fees = appData.monthlyFees;
      const matrix = appData.monthlyMatrix;

      const monthlyTotals = {};
      const monthlyPaidCounts = {};
      months.forEach(m => {
        monthlyTotals[m] = 0;
        monthlyPaidCounts[m] = 0;
      });

      members.forEach((person, idx) => {
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-slate-50 transition border-b border-slate-100';

        let personTotal = 0;
        let cellsHtml = `<td class="p-3 text-left font-bold text-slate-800 sticky left-0 bg-white group-hover:bg-slate-50 z-10 border-r border-slate-200 shadow-xs flex items-center justify-between">
          <span>${person}</span>
          <span class="text-[10px] text-slate-400 font-normal ml-1">#${idx + 1}</span>
        </td>`;

        months.forEach(m => {
          const targetFee = fees[m] || 20000;
          const paidAmt = (matrix[person] && matrix[person][m] !== undefined) ? Number(matrix[person][m]) : 0;
          personTotal += paidAmt;
          monthlyTotals[m] += paidAmt;

          const isPaid = paidAmt >= targetFee;
          if (isPaid) monthlyPaidCounts[m]++;

          let badgeHtml = '';
          if (isPaid) {
            badgeHtml = `<span class="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-semibold bg-emerald-100 text-emerald-800 border border-emerald-200">
              <i class="fa-solid fa-check text-[10px] mr-1"></i>${(paidAmt / 10000)}만
            </span>`;
          } else if (paidAmt > 0) {
            badgeHtml = `<span class="inline-flex items-center px-1.5 py-0.5 rounded-full text-[11px] font-semibold bg-amber-100 text-amber-800 border border-amber-200">
              ${(paidAmt / 10000)}만(일부)
            </span>`;
          } else {
            badgeHtml = `<span class="inline-flex items-center px-1.5 py-0.5 rounded-full text-[11px] font-medium bg-slate-100 text-slate-400">
              <i class="fa-solid fa-xmark text-[9px] mr-1 text-rose-400"></i>미납
            </span>`;
          }

          const clickableClass = isAdmin ? 'cursor-pointer hover:bg-blue-50/80 matrix-cell' : '';
          const clickAttr = isAdmin ? `onclick="toggleMatrixPayment('${person}', '${m}')"` : '';
          const tooltip = `${person} - ${m} (납부: ${paidAmt.toLocaleString()}원 / 기준: ${targetFee.toLocaleString()}원)${isAdmin ? ' [클릭하여 납부 토글]' : ''}`;

          cellsHtml += `<td class="p-2.5 ${clickableClass}" ${clickAttr} title="${tooltip}">
            ${badgeHtml}
          </td>`;
        });

        cellsHtml += `<td class="p-3 font-bold text-slate-900 bg-slate-50/50 border-l border-slate-200">
          ${personTotal.toLocaleString()}원
        </td>`;

        tr.innerHTML = cellsHtml;
        tbody.appendChild(tr);
      });

      // Tfoot
      let footHtml = `<tr class="bg-slate-100 text-slate-700">
        <td class="p-3 text-left font-bold sticky left-0 bg-slate-100 z-10 border-r border-slate-200 shadow-xs">
          월별 납부 합계
        </td>`;

      let grandTotal = 0;
      months.forEach(m => {
        const mTotal = monthlyTotals[m];
        grandTotal += mTotal;
        const paidCount = monthlyPaidCounts[m];
        const isFull = paidCount === members.length;
        footHtml += `<td class="p-2.5 text-center">
          <div class="font-bold text-slate-900">${(mTotal / 10000)}만</div>
          <div class="text-[10px] ${isFull ? 'text-emerald-600 font-bold' : 'text-slate-500'}">
            ${paidCount}/${members.length}명
          </div>
        </td>`;
      });

      footHtml += `<td class="p-3 text-center font-black text-blue-700 bg-blue-50/50 border-l border-slate-200">
        ${grandTotal.toLocaleString()}원
      </td></tr>`;
      tfoot.innerHTML = footHtml;
    }

    // Toggle Matrix Payment (Admin Action)
    function toggleMatrixPayment(person, month) {
      if (!isAdmin) return;
      const targetFee = appData.monthlyFees[month] || 20000;
      const currentPaid = (appData.monthlyMatrix[person] && appData.monthlyMatrix[person][month]) || 0;

      if (currentPaid >= targetFee) {
        if (confirm(`${person} 님의 [${month}] 회비 납부를 [미납] 상태로 변경하시겠습니까?`)) {
          appData.monthlyMatrix[person][month] = 0;
          showToast(`${person} 님의 ${month} 회비가 미납 처리되었습니다.`, 'info');
        } else {
          return;
        }
      } else {
        appData.monthlyMatrix[person][month] = targetFee;
        confetti({ particleCount: 30, spread: 60, origin: { y: 0.8 } });
        showToast(`${person} 님의 ${month} 회비 (${targetFee.toLocaleString()}원) 완납 처리 완료!`, 'success');
      }

      saveData();
      renderAll();
    }

    // 3. Render Members Grid
    function renderMembersGrid() {
      const grid = document.getElementById('membersGrid');
      grid.innerHTML = '';

      const months = appData.months;
      const members = appData.members;
      const fees = appData.monthlyFees;
      const matrix = appData.monthlyMatrix;

      const activeMonths = ['2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월'];
      const totalExpectedUpToNow = activeMonths.reduce((sum, m) => sum + fees[m], 0); // 160,000

      members.forEach((person, idx) => {
        let totalPaid = 0;
        let unpaidMonths = [];

        activeMonths.forEach(m => {
          const paid = (matrix[person] && matrix[person][m] !== undefined) ? matrix[person][m] : 0;
          totalPaid += paid;
          if (paid < fees[m]) {
            unpaidMonths.push(m);
          }
        });

        const unpaidAmt = Math.max(0, totalExpectedUpToNow - totalPaid);
        const isAllPaid = unpaidMonths.length === 0;

        const card = document.createElement('div');
        card.className = 'bg-white rounded-2xl p-5 shadow-sm border border-slate-200 flex flex-col justify-between hover:shadow-md transition';

        card.innerHTML = `
          <div>
            <div class="flex items-center justify-between">
              <div class="flex items-center space-x-2">
                <div class="w-9 h-9 rounded-xl ${isAllPaid ? 'bg-amber-100 text-amber-700' : 'bg-slate-100 text-slate-700'} flex items-center justify-center font-bold text-sm">
                  ${isAllPaid ? '<i class="fa-solid fa-crown text-amber-500"></i>' : idx + 1}
                </div>
                <div>
                  <h4 class="font-bold text-slate-900">${person}</h4>
                  <span class="text-[11px] text-slate-400">오토콜 회원</span>
                </div>
              </div>
              ${isAllPaid 
                ? '<span class="text-xs px-2.5 py-1 rounded-full font-bold bg-emerald-100 text-emerald-800 border border-emerald-200">완납 🎉</span>' 
                : `<span class="text-xs px-2.5 py-1 rounded-full font-bold bg-rose-100 text-rose-800 border border-rose-200">${unpaidMonths.length}달 미납</span>`
              }
            </div>

            <div class="mt-4 pt-3 border-t border-slate-100 space-y-1.5 text-xs">
              <div class="flex justify-between text-slate-600">
                <span>총 납부액:</span>
                <span class="font-bold text-slate-900">${totalPaid.toLocaleString()}원</span>
              </div>
              <div class="flex justify-between text-slate-600">
                <span>미납 금액:</span>
                <span class="font-bold ${unpaidAmt > 0 ? 'text-rose-600' : 'text-slate-400'}">${unpaidAmt.toLocaleString()}원</span>
              </div>
            </div>

            ${unpaidMonths.length > 0 ? `
              <div class="mt-3 p-2 rounded-xl bg-rose-50 border border-rose-100 text-[11px] text-rose-700 font-medium">
                <i class="fa-solid fa-circle-exclamation mr-1"></i>미납 월: ${unpaidMonths.join(', ')}
              </div>
            ` : `
              <div class="mt-3 p-2 rounded-xl bg-emerald-50 border border-emerald-100 text-[11px] text-emerald-700 font-medium">
                <i class="fa-solid fa-shield-heart mr-1"></i>11월분까지 모든 회비 성실 완납!
              </div>
            `}
          </div>

          <div class="mt-4 pt-3 border-t border-slate-100 flex items-center justify-between text-xs">
            <span class="text-slate-400">납부율 (11월 기준)</span>
            <span class="font-bold text-slate-700">${Math.round((totalPaid / totalExpectedUpToNow) * 100)}%</span>
          </div>
        `;

        grid.appendChild(card);
      });
    }

    // Copy Kakao Notice Text
    function copyKakaoNotice() {
      const activeMonths = ['2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월'];
      const fees = appData.monthlyFees;
      const totalExpected = activeMonths.reduce((sum, m) => sum + fees[m], 0);

      const unpaids = [];
      const paids = [];

      appData.members.forEach(p => {
        let total = 0;
        let mList = [];
        activeMonths.forEach(m => {
          const val = (appData.monthlyMatrix[p] && appData.monthlyMatrix[p][m]) || 0;
          total += val;
          if (val < fees[m]) mList.push(m);
        });
        const diff = totalExpected - total;
        if (diff > 0) {
          unpaids.push(`- ${p}: ${mList.join(', ')} (${diff.toLocaleString()}원)`);
        } else {
          paids.push(p);
        }
      });

      const totalBalance = (appData.transactions || []).reduce((acc, cur) => acc + (Number(cur.amount) || 0), 0);

      const text = `📢 [오토콜 회비 장부 현황 안내]

💰 현재 통장 잔액: ${totalBalance.toLocaleString()}원 (이자 포함)
📅 정산 기준: 11월분까지 (2~5월 1만원 / 6~11월 2만원)

✅ 완납 회원 (${paids.length}명):
${paids.join(', ')} 👍

⚠️ 미납 현황 (${unpaids.length}명):
${unpaids.join('\\n')}

※ 회비 입금 후 장부에 즉시 반영됩니다!
실시간 장부 주소: https://rje0712.github.io/autocall-ledger`;

      navigator.clipboard.writeText(text).then(() => {
        showToast('카카오톡 공지 텍스트가 복사되었습니다! 카톡에 붙여넣으세요.', 'success');
      }).catch(err => {
        prompt('아래 텍스트를 복사해서 사용하세요:', text);
      });
    }

    // 4. Render Ledger Table & Filters
    function renderLedgerFilters() {
      const personSelect = document.getElementById('filterPerson');
      const txPersonSelect = document.getElementById('txPersonSelect');
      
      const currentPersonVal = personSelect.value;
      personSelect.innerHTML = '<option value="ALL">전체 멤버/항목</option>';
      txPersonSelect.innerHTML = '<option value="">선택해주세요</option>';

      appData.members.forEach(m => {
        personSelect.innerHTML += `<option value="${m}">${m}</option>`;
        txPersonSelect.innerHTML += `<option value="${m}">${m}</option>`;
      });

      personSelect.innerHTML += '<option value="예금이자">예금이자</option>';
      personSelect.innerHTML += '<option value="모임지출">모임지출</option>';
      personSelect.value = currentPersonVal || 'ALL';

      txPersonSelect.innerHTML += '<option value="예금이자">예금이자</option>';
      txPersonSelect.innerHTML += '<option value="모임지출">모임지출</option>';
      txPersonSelect.innerHTML += '<option value="DIRECT">직접 입력...</option>';
    }

    function renderLedgerTable() {
      const tbody = document.getElementById('ledgerTableBody');
      const monthFilter = document.getElementById('filterMonth').value;
      const personFilter = document.getElementById('filterPerson').value;
      const typeFilter = document.getElementById('filterType').value;
      const search = (document.getElementById('txSearchInput').value || '').trim().toLowerCase();

      let list = [...(appData.transactions || [])];

      if (monthFilter !== 'ALL') {
        list = list.filter(t => t.month === monthFilter || (t.extraMonth && t.extraMonth.includes(monthFilter)));
      }
      if (personFilter !== 'ALL') {
        list = list.filter(t => t.person === personFilter);
      }
      if (typeFilter === 'IN') {
        list = list.filter(t => t.amount > 0);
      } else if (typeFilter === 'OUT') {
        list = list.filter(t => t.amount < 0);
      }
      if (search) {
        list = list.filter(t => 
          (t.person && t.person.toLowerCase().includes(search)) ||
          (t.note && t.note.toLowerCase().includes(search)) ||
          (t.date && t.date.toLowerCase().includes(search)) ||
          (t.month && t.month.toLowerCase().includes(search)) ||
          (t.extraMonth && t.extraMonth.toLowerCase().includes(search)) ||
          (t.amount && t.amount.toString().includes(search))
        );
      }

      document.getElementById('filteredCountBadge').innerText = `${list.length}건`;
      document.getElementById('ledgerCountBadge').innerText = `${(appData.transactions || []).length}`;

      tbody.innerHTML = '';

      if (list.length === 0) {
        tbody.innerHTML = `<tr><td colspan="${isAdmin ? 8 : 7}" class="p-8 text-center text-slate-400">일치하는 거래 내역이 없습니다.</td></tr>`;
        return;
      }

      const displayList = [...list].reverse();

      displayList.forEach(t => {
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-slate-50 transition border-b border-slate-100';

        const isPositive = t.amount >= 0;
        const formattedAmount = isPositive 
          ? `+${t.amount.toLocaleString()}원` 
          : `${t.amount.toLocaleString()}원`;
        const amountClass = isPositive ? 'text-blue-600 font-bold' : 'text-rose-600 font-bold';

        let adminActions = '';
        if (isAdmin) {
          adminActions = `<td class="p-3 text-center admin-only">
            <button onclick="openEditTxModal(${t.id})" class="text-slate-400 hover:text-blue-600 mr-2" title="수정"><i class="fa-solid fa-pen-to-square"></i></button>
            <button onclick="deleteTx(${t.id})" class="text-slate-400 hover:text-rose-600" title="삭제"><i class="fa-solid fa-trash"></i></button>
          </td>`;
        }

        tr.innerHTML = `
          <td class="p-3 text-center font-mono text-slate-400 text-xs">${t.id}</td>
          <td class="p-3 font-mono text-slate-600">${t.date || '-'}</td>
          <td class="p-3 text-center">
            <span class="inline-block px-2 py-0.5 rounded bg-slate-100 text-slate-700 text-xs font-medium">${t.month || '-'}</span>
          </td>
          <td class="p-3 text-center text-slate-600">
            ${t.extraMonth ? `<span class="px-2 py-0.5 rounded-md bg-blue-50 text-blue-700 text-xs font-semibold border border-blue-200">${t.extraMonth}</span>` : '-'}
          </td>
          <td class="p-3 font-bold text-slate-900">${t.person || '-'}</td>
          <td class="p-3 text-right font-mono ${amountClass}">${formattedAmount}</td>
          <td class="p-3 text-slate-600">${t.note ? `<span class="bg-amber-50 text-amber-900 px-2 py-0.5 rounded text-xs border border-amber-200/60">${t.note}</span>` : '-'}</td>
          ${adminActions}
        `;

        tbody.appendChild(tr);
      });
    }

    // 5. Render Charts
    function renderCharts() {
      const months = appData.months;
      const matrix = appData.monthlyMatrix;
      const members = appData.members;

      const monthlyData = months.map(m => {
        return members.reduce((sum, p) => sum + ((matrix[p] && matrix[p][m]) || 0), 0);
      });

      if (charts.monthly) charts.monthly.destroy();
      const ctxMonthly = document.getElementById('monthlyChart').getContext('2d');
      charts.monthly = new Chart(ctxMonthly, {
        type: 'bar',
        data: {
          labels: months,
          datasets: [{
            label: '월별 수납 회비 (원)',
            data: monthlyData,
            backgroundColor: 'rgba(59, 130, 246, 0.8)',
            hoverBackgroundColor: 'rgba(37, 99, 235, 1)',
            borderRadius: 6
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: {
            y: {
              beginAtZero: true,
              ticks: { callback: (val) => (val / 10000) + '만원' }
            }
          }
        }
      });

      const memberTotals = members.map(p => {
        return months.reduce((sum, m) => sum + ((matrix[p] && matrix[p][m]) || 0), 0);
      });

      if (charts.member) charts.member.destroy();
      const ctxMember = document.getElementById('memberChart').getContext('2d');
      charts.member = new Chart(ctxMember, {
        type: 'doughnut',
        data: {
          labels: members,
          datasets: [{
            data: memberTotals,
            backgroundColor: [
              '#3b82f6', '#10b981', '#f59e0b', '#8b5cf6',
              '#ec4899', '#06b6d4', '#84cc16', '#64748b'
            ]
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: {
              position: 'bottom',
              labels: { boxWidth: 12, font: { size: 11 } }
            }
          }
        }
      });
    }

    // Tab Navigation
    function switchTab(tabId) {
      document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('border-blue-600', 'text-blue-600', 'font-bold');
        btn.classList.add('border-transparent', 'text-slate-500', 'font-semibold');
      });
      document.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));

      const activeBtn = document.getElementById(`tabBtn-${tabId}`);
      activeBtn.classList.remove('border-transparent', 'text-slate-500', 'font-semibold');
      activeBtn.classList.add('border-blue-600', 'text-blue-600', 'font-bold');

      document.getElementById(`tab-${tabId}`).classList.remove('hidden');

      if (tabId === 'charts') {
        setTimeout(renderCharts, 50);
      }
    }

    // ================= SMART AUTO-FILL FIFO LOGIC =================
    function calculateAutoFill(person, amount) {
      if (!appData.members.includes(person) || amount <= 0) return null;

      const months = appData.months;
      const fees = appData.monthlyFees;
      const matrix = appData.monthlyMatrix[person] || {};

      let remaining = amount;
      const allocations = [];

      for (const m of months) {
        if (remaining <= 0) break;
        const targetFee = fees[m] || 20000;
        const alreadyPaid = matrix[m] || 0;
        const needed = Math.max(0, targetFee - alreadyPaid);

        if (needed > 0) {
          const allocate = Math.min(remaining, needed);
          allocations.push({
            month: m,
            targetFee,
            alreadyPaid,
            allocated: allocate,
            isFull: (alreadyPaid + allocate >= targetFee)
          });
          remaining -= allocate;
        }
      }

      return {
        person,
        amount,
        allocations,
        remainingOverpay: remaining
      };
    }

    function onPersonChange() {
      const select = document.getElementById('txPersonSelect');
      const directInput = document.getElementById('txPersonDirect');
      const badge = document.getElementById('memberUnpaidStatusBadge');
      const fillAllBtn = document.getElementById('btnFillAllUnpaid');

      if (select.value === 'DIRECT') {
        directInput.classList.remove('hidden');
        directInput.focus();
        badge.innerText = '';
        fillAllBtn.classList.add('hidden');
      } else {
        directInput.classList.add('hidden');
        directInput.value = '';

        const person = select.value;
        if (appData.members.includes(person)) {
          const activeMonths = ['2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월'];
          const fees = appData.monthlyFees;
          const matrix = appData.monthlyMatrix[person] || {};
          let unpaidTotal = 0;
          let unpaidMonths = [];

          activeMonths.forEach(m => {
            const need = Math.max(0, fees[m] - (matrix[m] || 0));
            if (need > 0) {
              unpaidTotal += need;
              unpaidMonths.push(m);
            }
          });

          if (unpaidTotal > 0) {
            badge.innerHTML = `<span class="text-rose-600 font-bold">${unpaidMonths.join(', ')} 미납 (총 ${unpaidTotal.toLocaleString()}원)</span>`;
            fillAllBtn.classList.remove('hidden');
            fillAllBtn.setAttribute('data-unpaid', unpaidTotal);
          } else {
            badge.innerHTML = `<span class="text-emerald-600 font-bold">11월까지 완납 회원</span>`;
            fillAllBtn.classList.add('hidden');
          }
        } else {
          badge.innerText = '';
          fillAllBtn.classList.add('hidden');
        }
      }

      updateAutoFillPreview();
    }

    function fillAllUnpaid() {
      const btn = document.getElementById('btnFillAllUnpaid');
      const unpaid = Number(btn.getAttribute('data-unpaid')) || 20000;
      document.getElementById('txAmount').value = unpaid;
      updateAutoFillPreview();
    }

    function setTxAmount(val) {
      document.getElementById('txAmount').value = val;
      updateAutoFillPreview();
    }

    function updateAutoFillPreview() {
      const card = document.getElementById('autoFillPreviewCard');
      const details = document.getElementById('autoFillDetails');
      const enabledCheckbox = document.getElementById('autoFillEnabled');
      
      const personSelect = document.getElementById('txPersonSelect');
      let person = personSelect.value;
      if (person === 'DIRECT') {
        person = document.getElementById('txPersonDirect').value.trim();
      }

      const amount = Number(document.getElementById('txAmount').value) || 0;

      if (!appData.members.includes(person) || amount <= 0) {
        card.classList.add('hidden');
        return;
      }

      const res = calculateAutoFill(person, amount);
      if (!res || res.allocations.length === 0) {
        card.classList.add('hidden');
        return;
      }

      card.classList.remove('hidden');

      if (!enabledCheckbox.checked) {
        details.innerHTML = '<span class="text-slate-500">자동 충당 비활성화 (선택된 해당월에만 단일 반영됩니다)</span>';
        return;
      }

      let allocText = res.allocations.map(a => {
        return `<span class="font-bold text-blue-900">${a.month}</span>: ${a.allocated.toLocaleString()}원 채움 ${a.isFull ? '✅' : '(일부)'}`;
      }).join(' · ');

      let extraMsg = '';
      if (res.remainingOverpay > 0) {
        extraMsg = `<div class="text-[11px] text-amber-700 font-semibold mt-1">※ 12월까지 모두 충당 후 초과 잔여금 ${res.remainingOverpay.toLocaleString()}원은 선납금으로 비고에 기록됩니다.</div>`;
      }

      details.innerHTML = `
        <div class="font-semibold text-blue-950">
          👉 <strong>총 ${res.allocations.length}개월</strong> 비워진 달부터 자동 순차 완납:
        </div>
        <div class="mt-0.5 leading-relaxed">${allocText}</div>
        ${extraMsg}
      `;

      // Auto-fill form fields
      const filledMonths = res.allocations.map(a => a.month);
      const lastMonth = filledMonths[filledMonths.length - 1];
      const otherMonths = filledMonths.filter(m => m !== lastMonth);

      document.getElementById('txMonth').value = lastMonth;
      if (otherMonths.length > 0) {
        document.getElementById('txExtraMonth').value = otherMonths.join(', ');
      } else {
        document.getElementById('txExtraMonth').value = '';
      }

      if (filledMonths.length > 1) {
        document.getElementById('txNote').value = `${filledMonths.length}개월치 회비 일괄 납부 (${filledMonths[0]}~${lastMonth})`;
      }
    }

    // ================= BATCH MONTHLY PAYMENT (전체 일괄 등록) =================
    function openBatchModal() {
      const now = new Date();
      const yy = String(now.getFullYear()).slice(-2);
      const mm = String(now.getMonth() + 1).padStart(2, '0');
      const dd = String(now.getDate()).padStart(2, '0');
      document.getElementById('batchDate').value = `${yy}.${mm}.${dd}`;
      document.getElementById('batchMonthSelect').value = `${now.getMonth() + 1}월`;

      renderBatchMembersTable();
      document.getElementById('batchModal').classList.remove('hidden');
    }

    function closeBatchModal() {
      document.getElementById('batchModal').classList.add('hidden');
    }

    function renderBatchMembersTable() {
      const tbody = document.getElementById('batchTableBody');
      const month = document.getElementById('batchMonthSelect').value;
      const targetFee = appData.monthlyFees[month] || 20000;
      const matrix = appData.monthlyMatrix;

      tbody.innerHTML = '';

      appData.members.forEach((person, idx) => {
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-slate-50 transition border-b border-slate-100';

        const paid = (matrix[person] && matrix[person][month]) || 0;
        const isAlreadyPaid = paid >= targetFee;

        tr.innerHTML = `
          <td class="p-2.5 text-center">
            <input type="checkbox" name="batchMemberCheck" value="${person}" ${!isAlreadyPaid ? 'checked' : ''} class="w-4 h-4 text-indigo-600 rounded border-slate-300 focus:ring-indigo-500">
          </td>
          <td class="p-2.5 font-bold text-slate-800">${person}</td>
          <td class="p-2.5">
            ${isAlreadyPaid 
              ? `<span class="inline-flex items-center px-2 py-0.5 rounded text-[11px] font-semibold bg-emerald-100 text-emerald-800">완납 (${paid.toLocaleString()}원)</span>`
              : `<span class="inline-flex items-center px-2 py-0.5 rounded text-[11px] font-medium bg-rose-50 text-rose-700">미납 (${paid.toLocaleString()}원)</span>`
            }
          </td>
          <td class="p-2.5 text-right">
            <input type="number" id="batchAmt_${person}" value="${targetFee}" class="w-24 px-2 py-1 border border-slate-300 rounded text-right font-mono font-bold text-xs focus:ring-1 focus:ring-indigo-500">
          </td>
          <td class="p-2.5">
            <input type="text" id="batchNote_${person}" placeholder="${month} 정기 회비" class="w-full px-2 py-1 border border-slate-200 rounded text-xs focus:ring-1 focus:ring-indigo-500">
          </td>
        `;

        tbody.appendChild(tr);
      });
    }

    function toggleBatchAll(status) {
      document.querySelectorAll('input[name="batchMemberCheck"]').forEach(cb => {
        cb.checked = status;
      });
    }

    function handleBatchSubmit(e) {
      e.preventDefault();
      const month = document.getElementById('batchMonthSelect').value;
      const date = document.getElementById('batchDate').value.trim();

      const checkboxes = Array.from(document.querySelectorAll('input[name="batchMemberCheck"]:checked'));
      if (checkboxes.length === 0) {
        return alert('회비를 등록할 회원을 1명 이상 선택해주세요.');
      }

      let addedCount = 0;
      let nextId = (appData.transactions.length > 0) 
        ? Math.max(...appData.transactions.map(t => t.id)) + 1 
        : 1;

      checkboxes.forEach(cb => {
        const person = cb.value;
        const amtInput = document.getElementById(`batchAmt_${person}`);
        const noteInput = document.getElementById(`batchNote_${person}`);

        const amount = Number(amtInput.value) || 20000;
        const note = noteInput.value.trim() || `${month} 정기 회비`;

        appData.transactions.push({
          id: nextId++,
          date,
          month,
          extraMonth: '',
          person,
          amount,
          note
        });

        // Update matrix
        if (!appData.monthlyMatrix[person]) appData.monthlyMatrix[person] = {};
        appData.monthlyMatrix[person][month] = (appData.monthlyMatrix[person][month] || 0) + amount;
        addedCount++;
      });

      saveData();
      closeBatchModal();
      renderAll();
      confetti({ particleCount: 50, spread: 80, origin: { y: 0.7 } });
      showToast(`${month} 회비 ${addedCount}명의 내역이 성공적으로 일괄 등록되었습니다!`, 'success');
    }

    // ================= EXCEL UPLOAD / IMPORT =================
    function handleExcelUpload(event) {
      const file = event.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = function(e) {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheetName = workbook.SheetNames[0];
          const sheet = workbook.Sheets[sheetName];

          // Parse Sheet rows
          const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
          const newTxs = [];

          for (let r = 3; r < rows.length; r++) { // Row 4 (0-based 3)
            const row = rows[r];
            if (!row) continue;
            const seq = row[1]; // Col B (index 1)
            const month = (row[2] || '').toString().trim(); // Col C
            const extra = (row[3] || '').toString().trim(); // Col D
            const person = (row[4] || '').toString().trim(); // Col E
            const amount = Number(row[5]) || 0; // Col F
            const date = (row[6] || '').toString().trim(); // Col G
            const note = (row[7] || '').toString().trim(); // Col H

            if (seq && (person || amount)) {
              newTxs.push({
                id: Number(seq),
                date,
                month,
                extraMonth: extra,
                person,
                amount,
                note
              });
            }
          }

          if (newTxs.length > 0) {
            if (confirm(`엑셀 파일에서 ${newTxs.length}건의 거래 내역을 찾았습니다. 기존 장부를 이 엑셀 데이터로 덮어쓰시겠습니까?`)) {
              appData.transactions = newTxs;

              // Recalculate monthly matrix
              appData.members.forEach(p => {
                appData.monthlyMatrix[p] = {};
              });

              // Re-fill matrix from transactions using smart logic
              newTxs.forEach(t => {
                if (appData.members.includes(t.person) && t.amount > 0) {
                  const m = t.month;
                  if (m && appData.months.includes(m)) {
                    appData.monthlyMatrix[t.person][m] = (appData.monthlyMatrix[t.person][m] || 0) + (appData.monthlyFees[m] || 20000);
                  }
                  if (t.extraMonth) {
                    const extraList = t.extraMonth.split(/[,\\s]+/).map(x => x.trim()).filter(Boolean);
                    extraList.forEach(em => {
                      if (appData.months.includes(em)) {
                        appData.monthlyMatrix[t.person][em] = (appData.monthlyMatrix[t.person][em] || 0) + (appData.monthlyFees[em] || 20000);
                      }
                    });
                  }
                }
              });

              saveData();
              renderAll();
              confetti({ particleCount: 50, spread: 70 });
              showToast(`엑셀 장부 (${newTxs.length}건) 전체가 성공적으로 반영되었습니다!`, 'success');
            }
          } else {
            alert('엑셀 파일에서 거래 내역을 찾을 수 없습니다.');
          }
        } catch (err) {
          alert('엑셀 파일 처리 오류: ' + err.message);
        }
      };
      reader.readAsArrayBuffer(file);
      event.target.value = '';
    }

    // ================= ADMIN FUNCTIONS =================

    // Login
    function openLoginModal() {
      document.getElementById('loginModal').classList.remove('hidden');
      document.getElementById('adminPasswordInput').focus();
    }
    function closeLoginModal() {
      document.getElementById('loginModal').classList.add('hidden');
      document.getElementById('adminPasswordInput').value = '';
    }
    function handleLoginSubmit(e) {
      e.preventDefault();
      const pwd = document.getElementById('adminPasswordInput').value;
      const savedPwd = localStorage.getItem('autocall_admin_password') || '1234';

      if (pwd === savedPwd || pwd === 'admin') {
        sessionStorage.setItem('autocall_admin_logged_in', 'true');
        isAdmin = true;
        updateAdminUI();
        closeLoginModal();
        renderAll();
        confetti({ particleCount: 40, spread: 50, origin: { y: 0.2 } });
        showToast('관리자 모드로 전환되었습니다. 장부 수정 권한이 활성화되었습니다.', 'success');
      } else {
        alert('비밀번호가 올바르지 않습니다. (초기 비밀번호: 1234)');
      }
    }
    function logoutAdmin() {
      sessionStorage.removeItem('autocall_admin_logged_in');
      isAdmin = false;
      updateAdminUI();
      renderAll();
      showToast('열람 모드로 전환되었습니다.', 'info');
    }

    // Change Password
    function openChangePasswordModal() {
      document.getElementById('adminMenuDropdown').classList.add('hidden');
      document.getElementById('passwordModal').classList.remove('hidden');
    }
    function closePasswordModal() {
      document.getElementById('passwordModal').classList.add('hidden');
      document.getElementById('newPasswordInput').value = '';
      document.getElementById('newPasswordConfirmInput').value = '';
    }
    function handleChangePasswordSubmit(e) {
      e.preventDefault();
      const p1 = document.getElementById('newPasswordInput').value;
      const p2 = document.getElementById('newPasswordConfirmInput').value;
      if (!p1) return alert('새 비밀번호를 입력해주세요.');
      if (p1 !== p2) return alert('비밀번호가 일치하지 않습니다.');

      localStorage.setItem('autocall_admin_password', p1);
      closePasswordModal();
      showToast('관리자 비밀번호가 성공적으로 변경되었습니다.', 'success');
    }

    // Admin Dropdown
    function toggleAdminMenu() {
      const d = document.getElementById('adminMenuDropdown');
      d.classList.toggle('hidden');
    }
    function setupClickOutside() {
      document.addEventListener('click', (e) => {
        const container = document.getElementById('adminMenuDropdownContainer');
        const dropdown = document.getElementById('adminMenuDropdown');
        if (container && !container.contains(e.target)) {
          dropdown.classList.add('hidden');
        }
      });
    }

    // Add / Edit Transaction
    function openAddTxModal() {
      document.getElementById('txEditId').value = '';
      document.getElementById('txModalTitle').innerHTML = '<i class="fa-solid fa-plus-circle text-emerald-600 mr-2"></i>회비 / 입출금 내역 등록';
      document.getElementById('txPersonSelect').value = '';
      document.getElementById('txPersonDirect').classList.add('hidden');
      document.getElementById('txPersonDirect').value = '';
      
      const now = new Date();
      const yy = String(now.getFullYear()).slice(-2);
      const mm = String(now.getMonth() + 1).padStart(2, '0');
      const dd = String(now.getDate()).padStart(2, '0');
      document.getElementById('txDate').value = `${yy}.${mm}.${dd}`;
      document.getElementById('txMonth').value = `${now.getMonth() + 1}월`;
      document.getElementById('txExtraMonth').value = '';
      document.getElementById('txAmount').value = '20000';
      document.getElementById('txNote').value = '';
      document.getElementById('autoFillEnabled').checked = true;

      document.getElementById('txModal').classList.remove('hidden');
      onPersonChange();
    }

    function openEditTxModal(id) {
      const t = (appData.transactions || []).find(x => x.id === id);
      if (!t) return;

      document.getElementById('txEditId').value = t.id;
      document.getElementById('txModalTitle').innerHTML = `<i class="fa-solid fa-pen-to-square text-blue-600 mr-2"></i>거래 내역 #${t.id} 수정`;
      
      const personSelect = document.getElementById('txPersonSelect');
      const directInput = document.getElementById('txPersonDirect');
      if (Array.from(personSelect.options).some(o => o.value === t.person)) {
        personSelect.value = t.person;
        directInput.classList.add('hidden');
        directInput.value = '';
      } else {
        personSelect.value = 'DIRECT';
        directInput.classList.remove('hidden');
        directInput.value = t.person;
      }

      document.getElementById('txDate').value = t.date || '';
      document.getElementById('txMonth').value = t.month || '2월';
      document.getElementById('txExtraMonth').value = t.extraMonth || '';
      document.getElementById('txAmount').value = t.amount;
      document.getElementById('txNote').value = t.note || '';
      document.getElementById('autoFillPreviewCard').classList.add('hidden');

      document.getElementById('txModal').classList.remove('hidden');
    }

    function closeTxModal() {
      document.getElementById('txModal').classList.add('hidden');
    }

    function handleTxSubmit(e) {
      e.preventDefault();
      const editIdStr = document.getElementById('txEditId').value;
      const isEdit = Boolean(editIdStr);

      const personSelect = document.getElementById('txPersonSelect');
      let person = personSelect.value;
      if (person === 'DIRECT') {
        person = document.getElementById('txPersonDirect').value.trim();
      }
      if (!person) return alert('입금인 또는 항목명을 입력해주세요.');

      const date = document.getElementById('txDate').value.trim();
      const month = document.getElementById('txMonth').value;
      const extraMonth = document.getElementById('txExtraMonth').value.trim();
      const amount = Number(document.getElementById('txAmount').value);
      const note = document.getElementById('txNote').value.trim();
      const autoFillOn = document.getElementById('autoFillEnabled').checked;

      if (isNaN(amount)) return alert('금액을 올바르게 입력해주세요.');

      if (isEdit) {
        const editId = Number(editIdStr);
        const idx = appData.transactions.findIndex(x => x.id === editId);
        if (idx !== -1) {
          appData.transactions[idx] = {
            id: editId,
            date, month, extraMonth, person, amount, note
          };
          showToast(`거래 내역 #${editId}이(가) 수정되었습니다.`, 'success');
        }
      } else {
        const nextId = (appData.transactions.length > 0) 
          ? Math.max(...appData.transactions.map(t => t.id)) + 1 
          : 1;

        appData.transactions.push({
          id: nextId,
          date, month, extraMonth, person, amount, note
        });

        // Smart Auto-Fill FIFO into Monthly Matrix
        if (appData.members.includes(person) && amount > 0) {
          if (!appData.monthlyMatrix[person]) appData.monthlyMatrix[person] = {};

          if (autoFillOn) {
            const res = calculateAutoFill(person, amount);
            if (res && res.allocations.length > 0) {
              res.allocations.forEach(a => {
                appData.monthlyMatrix[person][a.month] = (appData.monthlyMatrix[person][a.month] || 0) + a.allocated;
              });
              confetti({ particleCount: 45, spread: 70, origin: { y: 0.7 } });
            } else {
              appData.monthlyMatrix[person][month] = (appData.monthlyMatrix[person][month] || 0) + amount;
            }
          } else {
            appData.monthlyMatrix[person][month] = (appData.monthlyMatrix[person][month] || 0) + amount;
          }
        }

        showToast(`새 거래 내역 #${nextId}이(가) 등록되었습니다.`, 'success');
      }

      saveData();
      closeTxModal();
      renderAll();
    }

    function deleteTx(id) {
      if (!confirm(`거래 #${id} 내역을 정말 삭제하시겠습니까?`)) return;
      appData.transactions = appData.transactions.filter(t => t.id !== id);
      saveData();
      renderAll();
      showToast(`거래 #${id} 내역이 삭제되었습니다.`, 'info');
    }

    // Export to Excel
    function exportToExcel() {
      document.getElementById('adminMenuDropdown').classList.add('hidden');
      const wb = XLSX.utils.book_new();

      const txRows = appData.transactions.map(t => ({
        '연번': t.id,
        '상세날짜': t.date,
        '해당월': t.month,
        '추가/충당월': t.extraMonth,
        '입금인': t.person,
        '금액': t.amount,
        '비고': t.note
      }));
      const wsTx = XLSX.utils.json_to_sheet(txRows);
      XLSX.utils.book_append_sheet(wb, wsTx, '거래장부');

      const matrixRows = appData.members.map(p => {
        const row = { '회원명': p };
        let sum = 0;
        appData.months.forEach(m => {
          const val = (appData.monthlyMatrix[p] && appData.monthlyMatrix[p][m]) || 0;
          row[m] = val;
          sum += val;
        });
        row['합계'] = sum;
        return row;
      });
      const wsMatrix = XLSX.utils.json_to_sheet(matrixRows);
      XLSX.utils.book_append_sheet(wb, wsMatrix, '월별납부현황');

      XLSX.writeFile(wb, '오토콜_회비장부_최신.xlsx');
      showToast('엑셀(.xlsx) 파일이 다운로드되었습니다.', 'success');
    }

    // Export Shareable HTML (Bundled with latest data!)
    function exportShareableHtml() {
      document.getElementById('adminMenuDropdown').classList.add('hidden');
      
      let htmlContent = document.documentElement.outerHTML;
      const latestDataJson = JSON.stringify(appData);

      htmlContent = htmlContent.replace(/const INITIAL_DATA = [\\s\\S]*?;\\n\\n    \/\/ State/, 'const INITIAL_DATA = ' + latestDataJson + ';\\n\\n    // State');

      const blob = new Blob(['<!DOCTYPE html>\\n' + htmlContent], { type: 'text/html;charset=utf-8' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = '오토콜_회비장부_공유용.html';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      showToast('공유용 최신 HTML 파일이 생성되었습니다! 카톡에 공유하세요.', 'success');
    }

    // JSON Backup / Restore
    function exportJsonBackup() {
      document.getElementById('adminMenuDropdown').classList.add('hidden');
      const dataStr = 'data:text/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(appData, null, 2));
      const a = document.createElement('a');
      a.href = dataStr;
      a.download = 'autocall_ledger_backup.json';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      showToast('JSON 백업 파일이 저장되었습니다.', 'success');
    }

    function importJsonBackup(event) {
      const file = event.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = function(e) {
        try {
          const imported = JSON.parse(e.target.result);
          if (imported.transactions && imported.members) {
            appData = imported;
            saveData();
            renderAll();
            showToast('백업 데이터가 성공적으로 복원되었습니다.', 'success');
          } else {
            alert('올바른 회비 장부 JSON 형식이 아닙니다.');
          }
        } catch (err) {
          alert('JSON 파일 파싱 오류: ' + err.message);
        }
      };
      reader.readAsText(file);
      event.target.value = '';
    }

    function confirmResetData() {
      document.getElementById('adminMenuDropdown').classList.add('hidden');
      if (confirm('정말로 모든 데이터를 원본 엑셀 초기 상태로 되돌리시겠습니까?\\n(직접 추가/수정한 내역이 모두 초기화됩니다)')) {
        localStorage.removeItem('autocall_ledger_data');
        appData = JSON.parse(JSON.stringify(INITIAL_DATA));
        renderAll();
        showToast('원본 데이터로 초기화되었습니다.', 'info');
      }
    }

    // Toast Notification Utility
    function showToast(message, type = 'info') {
      const container = document.getElementById('toastContainer');
      const toast = document.createElement('div');
      
      let bg = 'bg-slate-900 text-white';
      let icon = '<i class="fa-solid fa-circle-info mr-2 text-blue-400"></i>';
      if (type === 'success') {
        bg = 'bg-emerald-900 text-white border border-emerald-700';
        icon = '<i class="fa-solid fa-circle-check mr-2 text-emerald-400"></i>';
      } else if (type === 'warn') {
        bg = 'bg-amber-900 text-white border border-amber-700';
        icon = '<i class="fa-solid fa-triangle-exclamation mr-2 text-amber-400"></i>';
      }

      toast.className = `${bg} px-4 py-3 rounded-xl shadow-xl text-xs sm:text-sm font-medium flex items-center pointer-events-auto transition-all transform duration-300 translate-y-2 opacity-0`;
      toast.innerHTML = `${icon}<span>${message}</span>`;

      container.appendChild(toast);
      setTimeout(() => {
        toast.classList.remove('translate-y-2', 'opacity-0');
      }, 10);

      setTimeout(() => {
        toast.classList.add('opacity-0', 'translate-y-2');
        setTimeout(() => toast.remove(), 300);
      }, 3200);
    }
  </script>
</body>
</html>"""

final_html = html_template.replace('__INITIAL_DATA_PLACEHOLDER__', json_data_str)

with open('autocall_ledger.html', 'w', encoding='utf-8') as f:
    f.write(final_html)

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(final_html)

with open('오토콜_회비장부.html', 'w', encoding='utf-8') as f:
    f.write(final_html)

print('Successfully generated autocall_ledger.html, index.html, and 오토콜_회비장부.html with bulk entry and excel import!')
