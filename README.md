<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>편의점 AI 스마트 재고관리 & 자동발주 시스템 (106개 품목 POS 연동)</title>
    
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- Excel 내보내기 (XLSX) -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <!-- Google Fonts Noto Sans KR -->
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">

    <script>
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: {
                        sans: ['"Noto Sans KR"', 'sans-serif'],
                    },
                    colors: {
                        brand: {
                            50: '#f0f7ff',
                            100: '#e0effe',
                            500: '#1d4ed8',
                            600: '#1e40af',
                            700: '#1e3a8a',
                            800: '#0f172a',
                        }
                    }
                }
            }
        }
    </script>

    <style>
        body { font-family: 'Noto Sans KR', sans-serif; background-color: #f8fafc; color: #1e293b; }
        .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 4px; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
        .badge-risk-stockout { background-color: #fef2f2; color: #dc2626; border: 1px solid #fecaca; }
        .badge-risk-waste { background-color: #fff7ed; color: #ea580c; border: 1px solid #ffedd5; }
        .badge-risk-overstock { background-color: #fefce8; color: #ca8a04; border: 1px solid #fef08a; }
        .badge-risk-normal { background-color: #f0fdf4; color: #16a34a; border: 1px solid #bbf7d0; }
        .tab-btn.active { color: #2563eb; border-bottom: 3px solid #2563eb; font-weight: 700; background-color: #ffffff; }
        /* 초보 점주용 가독성 개선: 기존 기능과 데이터 구조는 유지 */
        .beginner-guide { background: linear-gradient(135deg, #eff6ff 0%, #ffffff 72%); border: 1px solid #bfdbfe; border-radius: 18px; padding: 22px; box-shadow: 0 4px 14px rgba(30, 64, 175, .06); }
        .beginner-step { background: #fff; border: 1px solid #dbeafe; border-radius: 13px; padding: 15px; min-height: 118px; }
        .beginner-step-number { display: inline-flex; align-items: center; justify-content: center; width: 28px; height: 28px; border-radius: 9px; background: #dbeafe; color: #1d4ed8; font-weight: 800; margin-bottom: 8px; }
        .quick-action { display: flex; align-items: center; justify-content: space-between; gap: 12px; width: 100%; text-align: left; border-radius: 12px; padding: 14px 16px; border: 1px solid #e2e8f0; background: #fff; transition: transform .15s, box-shadow .15s; }
        .quick-action:hover { transform: translateY(-1px); box-shadow: 0 5px 14px rgba(15, 23, 42, .08); }
        .quick-action:focus-visible, .tab-btn:focus-visible { outline: 3px solid #93c5fd; outline-offset: 2px; }
        .section-helper { color: #64748b; font-size: 12px; line-height: 1.6; }
        @media (max-width: 640px) { .beginner-guide { padding: 16px; } main { padding-top: 16px !important; } }
    </style>
</head>
<body class="min-h-screen flex flex-col bg-slate-50">

    <!-- Navigation Top Bar -->
    <header class="bg-slate-900 text-white sticky top-0 z-50 shadow-md">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex items-center justify-between h-16">
                <!-- Brand Title -->
                <div class="flex items-center space-x-3">
                    <div class="bg-blue-600 p-2.5 rounded-xl shadow-lg flex items-center justify-center">
                        <i class="fa-solid fa-store text-xl text-white"></i>
                    </div>
                    <div>
                        <div class="flex items-center space-x-2">
                            <span class="text-lg font-bold tracking-tight text-white">GS24 스마트 점포 AI</span>
                            <span class="bg-blue-500/30 text-blue-300 text-xs px-2 py-0.5 rounded-full font-medium border border-blue-400/30">106개 SKU 전수 연동</span>
                        </div>
                        <p class="text-xs text-slate-400">9개월 POS 분석 기반 기준값(기초/안전/기말재고, ROP, 수요예측) 대시보드</p>
                    </div>
                </div>

                <!-- Live Simulation Indicator & Date Banner -->
                <div class="flex items-center space-x-4">
                    <div class="hidden md:flex items-center space-x-3 bg-slate-800/90 border border-slate-700/80 px-3.5 py-1.5 rounded-lg text-sm">
                        <i class="fa-regular fa-calendar-days text-blue-400"></i>
                        <span class="text-slate-300 font-medium">시뮬레이션 일자:</span>
                        <span id="currentSimDateDisplay" class="font-bold text-amber-300">2026-10-01 (목)</span>
                    </div>

                    <button onclick="switchTab('simulator')" class="bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white text-xs sm:text-sm font-semibold px-3.5 py-2 rounded-lg shadow-md transition flex items-center space-x-2">
                        <i class="fa-solid fa-bolt text-yellow-300 animate-pulse"></i>
                        <span>POS 시뮬레이터</span>
                    </button>
                </div>
            </div>
        </div>

        <!-- Tab Navigation Menu -->
        <nav class="bg-slate-800 border-t border-slate-700/60 overflow-x-auto custom-scrollbar">
            <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex space-x-1 sm:space-x-4">
                <button id="tabBtn-dashboard" onclick="switchTab('dashboard')" class="tab-btn px-4 py-3 text-sm font-medium text-slate-300 hover:text-white transition whitespace-nowrap flex items-center space-x-2">
                    <i class="fa-solid fa-chart-line text-blue-500"></i>
                    <span>이익증가 대시보드</span>
                </button>
                <button id="tabBtn-forecast" onclick="switchTab('forecast')" class="tab-btn px-4 py-3 text-sm font-medium text-slate-300 hover:text-white transition whitespace-nowrap flex items-center space-x-2">
                    <i class="fa-solid fa-wand-magic-sparkles text-purple-400"></i>
                    <span>106개 품목 판매예측</span>
                </button>
                <button id="tabBtn-order" onclick="switchTab('order')" class="tab-btn active px-4 py-3 text-sm font-medium text-slate-300 hover:text-white transition whitespace-nowrap flex items-center space-x-2">
                    <i class="fa-solid fa-cart-shopping text-emerald-400"></i>
                    <span>발주·재고위험 관리</span>
                    <span id="riskCountBadge" class="bg-red-500 text-white text-xs px-1.5 py-0.5 rounded-full font-bold ml-1">0</span>
                </button>
                <button id="tabBtn-simulator" onclick="switchTab('simulator')" class="tab-btn px-4 py-3 text-sm font-medium text-slate-300 hover:text-white transition whitespace-nowrap flex items-center space-x-2">
                    <i class="fa-solid fa-calculator text-cyan-400"></i>
                    <span>POS 데이터 자동생성</span>
                </button>
            </div>
        </nav>
    </header>

    <!-- Main Content Area -->
    <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 flex-grow w-full">

        <!-- TAB 1: 이익증가 대시보드 -->
        <section id="tab-dashboard" class="tab-content hidden space-y-6">
            <!-- 첫 방문 점주용 시작 안내: 아래 기존 분석/기능은 모두 그대로 유지 -->
            <div class="beginner-guide">
                <div class="flex flex-col lg:flex-row lg:items-start lg:justify-between gap-4 mb-5">
                    <div>
                        <div class="inline-flex items-center gap-2 text-xs font-bold text-blue-700 bg-blue-100 rounded-full px-3 py-1 mb-2"><i class="fa-solid fa-hand-sparkles"></i> 처음 사용하는 사장님 안내</div>
                        <h2 class="text-xl sm:text-2xl font-extrabold text-slate-900">오늘은 이 3가지만 확인하세요</h2>
                        <p class="text-sm text-slate-600 mt-2">복잡한 숫자를 모두 이해하지 않아도 됩니다. 위험 품목을 확인하고, 추천 발주량을 검토한 뒤, 필요하면 품목별 상세 내용을 살펴보세요.</p>
                    </div>
                    <div class="text-xs text-slate-500 bg-white border border-blue-100 rounded-xl px-3 py-2 shrink-0"><i class="fa-regular fa-calendar mr-1"></i> 기준일 <strong id="beginnerGuideDate" class="text-slate-800">2026-10-01</strong></div>
                </div>
                <div class="grid grid-cols-1 md:grid-cols-3 gap-3 mb-5">
                    <div class="beginner-step"><div class="beginner-step-number">1</div><h3 class="font-bold text-slate-900">위험 품목 확인</h3><p class="section-helper mt-1">품절되기 쉬운 상품과 유통기한·폐기 위험 상품부터 확인합니다.</p><button onclick="switchTab('risk')" class="mt-3 text-sm font-bold text-blue-700 hover:text-blue-900">위험 품목 보러 가기 <i class="fa-solid fa-arrow-right ml-1"></i></button></div>
                    <div class="beginner-step"><div class="beginner-step-number">2</div><h3 class="font-bold text-slate-900">추천 발주량 검토</h3><p class="section-helper mt-1">오늘과 내일 필요한 수량을 확인하고, 추천 수량을 점포 상황에 맞게 조정합니다.</p><button onclick="switchTab('order')" class="mt-3 text-sm font-bold text-blue-700 hover:text-blue-900">발주 관리 열기 <i class="fa-solid fa-arrow-right ml-1"></i></button></div>
                    <div class="beginner-step"><div class="beginner-step-number">3</div><h3 class="font-bold text-slate-900">필요할 때 상세 확인</h3><p class="section-helper mt-1">특정 상품의 판매 예측, 재고와 안전재고를 확인할 때 품목별 화면을 사용합니다.</p><button onclick="switchTab('forecast')" class="mt-3 text-sm font-bold text-blue-700 hover:text-blue-900">106개 품목 보기 <i class="fa-solid fa-arrow-right ml-1"></i></button></div>
                </div>
                <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <button onclick="switchTab('risk')" class="quick-action"><span class="flex items-center gap-3"><span class="w-10 h-10 rounded-xl bg-red-50 text-red-600 flex items-center justify-center"><i class="fa-solid fa-triangle-exclamation"></i></span><span><strong class="block text-slate-900">먼저 확인할 위험 품목</strong><span class="text-xs text-slate-500">품절·폐기·과다재고 상태 확인</span></span></span><span class="text-lg font-extrabold text-red-600" id="beginnerRiskCount">확인하기</span></button>
                    <button onclick="switchTab('order')" class="quick-action"><span class="flex items-center gap-3"><span class="w-10 h-10 rounded-xl bg-emerald-50 text-emerald-700 flex items-center justify-center"><i class="fa-solid fa-cart-shopping"></i></span><span><strong class="block text-slate-900">오늘 발주할 상품</strong><span class="text-xs text-slate-500">AI 추천 수량 확인 후 직접 조정 가능</span></span></span><span class="text-lg font-extrabold text-emerald-700">발주 보기 <i class="fa-solid fa-chevron-right text-sm"></i></span></button>
                </div>
                <p class="text-xs text-slate-500 mt-4"><i class="fa-solid fa-circle-info mr-1"></i> 아래의 대시보드, 판매예측, 발주·재고위험 통합관리, POS 자동생성 기능을 이용할 수 있습니다.</p>
            </div>
            <div class="bg-slate-900 text-white p-4 rounded-xl shadow-md border border-slate-800">
                <div class="flex items-center justify-between border-b border-slate-800 pb-2 mb-3">
                    <h3 class="text-xs sm:text-sm font-bold text-amber-300 flex items-center">
                        <i class="fa-solid fa-clipboard-check mr-2 text-amber-400"></i>
                        [오늘 현황 종합 집계] 106개 품목 기초/안전/기말재고 & 수량 계산 현황
                    </h3>
                    <span class="text-[11px] text-slate-400">기준일자: <strong id="todayStatusBannerDate" class="text-slate-200">2026-10-01</strong></span>
                </div>
                <div class="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-2 text-center text-xs">
                    <div class="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700/60">
                        <span class="text-slate-400 block text-[11px]">① 전체 집계</span>
                        <strong class="text-sm text-white font-bold mt-0.5 block">106 SKU</strong>
                    </div>
                    <div class="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700/60">
                        <span class="text-slate-400 block text-[11px]">② 총 기초재고</span>
                        <strong id="statSumCurrentStock" class="text-sm text-slate-200 font-bold mt-0.5 block">0개</strong>
                    </div>
                    <div class="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700/60">
                        <span class="text-slate-400 block text-[11px]">③ 오늘 예측수요</span>
                        <strong id="statSumForecastD0" class="text-sm text-blue-300 font-bold mt-0.5 block">0개</strong>
                    </div>
                    <div class="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700/60">
                        <span class="text-slate-400 block text-[11px]">④ 총 안전재고(SS)</span>
                        <strong id="statSumSafetyStock" class="text-sm text-emerald-300 font-bold mt-0.5 block">0개</strong>
                    </div>
                    <div class="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700/60">
                        <span class="text-slate-400 block text-[11px]">⑤ 총 기말재고예정</span>
                        <strong id="statSumEndingStock" class="text-sm text-cyan-300 font-bold mt-0.5 block">0개</strong>
                    </div>
                    <div class="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700/60">
                        <span class="text-slate-400 block text-[11px]">⑥ 총 목표재고</span>
                        <strong id="statSumTargetStock" class="text-sm text-indigo-300 font-bold mt-0.5 block">0개</strong>
                    </div>
                    <div class="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700/60">
                        <span class="text-slate-400 block text-[11px]">⑦ 품절위험</span>
                        <strong id="statCountStockout" class="text-sm text-red-400 font-bold mt-0.5 block">0개</strong>
                    </div>
                    <div class="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700/60">
                        <span class="text-slate-400 block text-[11px]">⑧ 폐기위험</span>
                        <strong id="statCountWaste" class="text-sm text-orange-400 font-bold mt-0.5 block">0개</strong>
                    </div>
                </div>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                <div class="bg-white rounded-xl p-5 shadow-sm border border-slate-200 hover:shadow-md transition">
                    <div class="flex items-center justify-between">
                        <span class="text-xs font-bold text-slate-500 uppercase tracking-wider">월 순이익 예상 증가액</span>
                        <div class="bg-blue-50 text-blue-600 p-2 rounded-lg">
                            <i class="fa-solid fa-sack-dollar text-xl"></i>
                        </div>
                    </div>
                    <div class="mt-3">
                        <span id="kpiNetProfit" class="text-2xl sm:text-3xl font-extrabold text-blue-600">+387,426원</span>
                        <span class="text-xs text-emerald-600 font-semibold ml-1"><i class="fa-solid fa-arrow-up"></i> +18.4%</span>
                    </div>
                    <p class="text-xs text-slate-500 mt-2">연간 환산 시 <strong class="text-slate-800" id="kpiAnnualProfit">+4,649,112원</strong> 수익 추가 창출</p>
                </div>
                <div class="bg-white rounded-xl p-5 shadow-sm border border-slate-200 hover:shadow-md transition">
                    <div class="flex items-center justify-between">
                        <span class="text-xs font-bold text-slate-500 uppercase tracking-wider">9개월 누적 폐기 손실</span>
                        <div class="bg-orange-50 text-orange-600 p-2 rounded-lg">
                            <i class="fa-solid fa-dumpster text-xl"></i>
                        </div>
                    </div>
                    <div class="mt-3">
                        <span id="kpiWasteSaved" class="text-2xl sm:text-3xl font-extrabold text-orange-600">277,520원</span>
                        <span class="text-xs text-slate-500 font-medium ml-1"> (9개월 합계)</span>
                    </div>
                    <p class="text-xs text-slate-500 mt-2">AI 목표재고 적용 시 월평균 <strong class="text-emerald-600 font-bold">19,270원</strong> 절감</p>
                </div>
                <div class="bg-white rounded-xl p-5 shadow-sm border border-slate-200 hover:shadow-md transition">
                    <div class="flex items-center justify-between">
                        <span class="text-xs font-bold text-slate-500 uppercase tracking-wider">9개월 누적 품절 손실</span>
                        <div class="bg-emerald-50 text-emerald-600 p-2 rounded-lg">
                            <i class="fa-solid fa-arrow-trend-up text-xl"></i>
                        </div>
                    </div>
                    <div class="mt-3">
                        <span id="kpiStockoutRecovered" class="text-2xl sm:text-3xl font-extrabold text-emerald-600">35,270원</span>
                        <span class="text-xs text-slate-500 font-medium ml-1"> (9개월 합계)</span>
                    </div>
                    <p class="text-xs text-slate-500 mt-2">재발주점(ROP) 자동 연동으로 품절손실 100% 방지</p>
                </div>
                <div class="bg-white rounded-xl p-5 shadow-sm border border-slate-200 hover:shadow-md transition">
                    <div class="flex items-center justify-between">
                        <span class="text-xs font-bold text-slate-500 uppercase tracking-wider">9개월 POS 총매출액</span>
                        <div class="bg-purple-50 text-purple-600 p-2 rounded-lg">
                            <i class="fa-solid fa-store text-xl"></i>
                        </div>
                    </div>
                    <div class="mt-3">
                        <span class="text-2xl sm:text-3xl font-extrabold text-purple-600">5억 8,816만</span>
                        <span class="text-xs text-slate-400 font-normal"> (28,938건)</span>
                    </div>
                    <p class="text-xs text-slate-500 mt-2">106개 품목 전수 POS 판매실적 기반 데이터</p>
                </div>
            </div>

            <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <div class="lg:col-span-2 bg-white rounded-xl p-5 shadow-sm border border-slate-200">
                    <div class="flex items-center justify-between mb-4">
                        <div>
                            <h3 class="text-base font-bold text-slate-800">월별 손실 감축 & 이익증가 추이 비교</h3>
                            <p class="text-xs text-slate-500">기존 경험 발주 vs 9개월 POS 분석 기반 AI 수요예측 발주</p>
                        </div>
                        <span class="text-xs bg-slate-100 text-slate-600 px-2.5 py-1 rounded-md font-medium">실제 POS 데이터 연동</span>
                    </div>
                    <div class="h-64 sm:h-80 relative">
                        <canvas id="chartProfitTrend"></canvas>
                    </div>
                </div>
                <div class="bg-white rounded-xl p-5 shadow-sm border border-slate-200">
                    <div class="flex items-center justify-between mb-4">
                        <div>
                            <h3 class="text-base font-bold text-slate-800">106개 품목 재고 상태 분포</h3>
                            <p class="text-xs text-slate-500">목표재고 및 ROP 대비 실시간 위험 비율</p>
                        </div>
                    </div>
                    <div class="h-56 relative flex items-center justify-center">
                        <canvas id="chartRiskRatio"></canvas>
                    </div>
                    <div class="mt-4 grid grid-cols-2 gap-2 text-xs">
                        <div class="flex items-center justify-between p-2 rounded bg-red-50 border border-red-100">
                            <span class="text-red-700 font-medium"><i class="fa-solid fa-circle text-red-500 text-[10px] mr-1"></i>품절위험</span>
                            <span id="dashboardCountStockout" class="font-bold text-red-700">0개</span>
                        </div>
                        <div class="flex items-center justify-between p-2 rounded bg-orange-50 border border-orange-100">
                            <span class="text-orange-700 font-medium"><i class="fa-solid fa-circle text-orange-500 text-[10px] mr-1"></i>폐기위험</span>
                            <span id="dashboardCountWaste" class="font-bold text-orange-700">0개</span>
                        </div>
                        <div class="flex items-center justify-between p-2 rounded bg-yellow-50 border border-yellow-100">
                            <span class="text-yellow-700 font-medium"><i class="fa-solid fa-circle text-yellow-500 text-[10px] mr-1"></i>과다재고</span>
                            <span id="dashboardCountOverstock" class="font-bold text-yellow-700">0개</span>
                        </div>
                        <div class="flex items-center justify-between p-2 rounded bg-emerald-50 border border-emerald-100">
                            <span class="text-emerald-700 font-medium"><i class="fa-solid fa-circle text-emerald-500 text-[10px] mr-1"></i>정상재고</span>
                            <span id="dashboardCountNormal" class="font-bold text-emerald-700">0개</span>
                        </div>
                    </div>
                </div>
            </div>

            <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                <div class="p-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
                    <h3 class="text-sm font-bold text-slate-800 flex items-center">
                        <i class="fa-solid fa-calculator text-blue-600 mr-2"></i>
                        카테고리별 9개월 POS 실적 및 AI 도입 수익 개선 내역 (106개 품목 전수)
                    </h3>
                    <span class="text-xs text-slate-500">실제 Excel POS 연동 데이터</span>
                </div>
                <div class="overflow-x-auto">
                    <table class="w-full text-left text-xs sm:text-sm">
                        <thead class="bg-slate-100 text-slate-600 border-b border-slate-200 font-semibold">
                            <tr>
                                <th class="p-3">카테고리</th>
                                <th class="p-3">품목수</th>
                                <th class="p-3 text-right">9개월 매출액</th>
                                <th class="p-3 text-right">9개월 누적 폐기손실</th>
                                <th class="p-3 text-right">9개월 품절기회손실</th>
                                <th class="p-3 text-right text-blue-700">월 예상 이익증가</th>
                            </tr>
                        </thead>
                        <tbody id="categoryProfitTableBody" class="divide-y divide-slate-100 text-slate-700">
                            <!-- Populated dynamically via JS -->
                        </tbody>
                    </table>
                </div>
            </div>
        </section>

        <!-- TAB 2: 106개 품목 판매예측 -->
        <section id="tab-forecast" class="tab-content hidden space-y-4">
            <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-col md:flex-row items-center justify-between gap-3">
                <div class="flex items-center space-x-2 w-full md:w-auto">
                    <div class="relative w-full md:w-64">
                        <i class="fa-solid fa-magnifying-glass absolute left-3 top-2.5 text-slate-400 text-sm"></i>
                        <input type="text" id="forecastSearchInput" onkeyup="filterForecastTable()" placeholder="품목명 또는 SKU 검색..." class="w-full pl-9 pr-3 py-1.5 text-sm bg-slate-50 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500">
                    </div>
                    <select id="forecastCategoryFilter" onchange="filterForecastTable()" class="py-1.5 px-3 text-sm bg-slate-50 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500">
                        <option value="ALL">전체 카테고리 (8개)</option>
                        <option value="FF/즉석식">FF/즉석식 (20개)</option>
                        <option value="음료">음료 (20개)</option>
                        <option value="커피/에너지">커피/에너지 (10개)</option>
                        <option value="유제품">유제품 (10개)</option>
                        <option value="라면/즉석조리">라면/즉석조리 (12개)</option>
                        <option value="과자/초콜릿/젤리">과자/초콜릿/젤리 (16개)</option>
                        <option value="아이스크림/냉동">아이스크림/냉동 (10개)</option>
                        <option value="생활/위생">생활/위생 (8개)</option>
                    </select>
                </div>
                <div class="flex items-center space-x-3 text-xs text-slate-600 w-full md:w-auto justify-end">
                    <span class="bg-blue-50 text-blue-700 px-2.5 py-1 rounded-md font-semibold border border-blue-200">
                        표시 품목: <strong id="forecastDisplayedCount" class="text-blue-800">106</strong> / 106 SKU
                    </span>
                    <button onclick="resetForecastFilters()" class="text-slate-500 hover:text-slate-800 underline">필터 초기화</button>
                </div>
            </div>
            <div class="rounded-xl border border-cyan-200 bg-cyan-50 px-4 py-3">
                <div class="font-extrabold text-cyan-950"><i class="fa-solid fa-boxes-stacked mr-2"></i>재고 3종을 한눈에 확인</div>
                <p class="mt-1 text-sm text-cyan-900">각 품목의 <strong>기초재고 → 안전재고 → 기말재고(예정)</strong>를 나란히 배치했어. 기말재고는 기초재고 + 오늘 발주/입고 − 오늘 예측판매 기준으로 계산돼.</p>
            </div>

            <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                <div class="overflow-x-auto custom-scrollbar">
                    <table class="w-full text-left text-xs sm:text-sm">
                        <thead class="bg-slate-800 text-white font-semibold">
                            <tr>
                                <th class="p-3 w-16">코드</th>
                                <th class="p-3 w-24">카테고리</th>
                                <th class="p-3">상품명</th>
                                <th class="p-3 text-center">ABC</th>
                                <th class="p-3 text-center bg-slate-700">기초재고</th>
                                <th class="p-3 text-center text-emerald-300 bg-emerald-950">안전재고(SS)</th>
                                <th class="p-3 text-center bg-cyan-900 text-cyan-200 font-bold" title="기초재고 + 오늘 발주/입고 - 오늘 예측판매">기말재고(예정)</th>
                                <th class="p-3 text-center bg-blue-900 text-blue-200" id="thDateD0">오늘예측(D0)</th>
                                <th class="p-3 text-center bg-purple-900 text-purple-200" title="비워두면 자동 계산됨">실제판매(수동) ✍️</th>
                                <th class="p-3 text-center bg-emerald-900 text-emerald-100">오늘입고/발주</th>
                                <th class="p-3 text-center text-amber-300">ROP</th>
                                <th class="p-3 text-center bg-indigo-900 text-indigo-100">목표재고</th>
                                <th class="p-3 text-center bg-blue-800 text-blue-100" id="thDateD1">내일 (10/02)</th>
                                <th class="p-3 text-center bg-slate-700" id="thDateD2">+2일 (10/03)</th>
                                <th class="p-3 text-center bg-slate-700" id="thDateD3">+3일 (10/04)</th>
                            </tr>
                        </thead>
                        <tbody id="forecastTableBody" class="divide-y divide-slate-100">
                            <!-- Populated via JS -->
                        </tbody>
                    </table>
                </div>
            </div>
        </section>

        <!-- TAB 3: 발주·재고관리 (내부 메뉴로 분리) -->
        <section id="tab-order" class="tab-content space-y-4">
            <div class="grid grid-cols-2 gap-2 rounded-xl bg-slate-100 p-2 border border-slate-200">
                <button id="orderSubBtn-order" onclick="switchOrderSubTab('order')" class="rounded-lg bg-emerald-600 text-white px-4 py-3 text-sm font-extrabold shadow-sm flex items-center justify-center gap-2">
                    <i class="fa-solid fa-cart-shopping"></i><span>발주하기</span>
                </button>
                <button id="orderSubBtn-risk" onclick="switchOrderSubTab('risk')" class="rounded-lg bg-white text-slate-700 hover:bg-amber-50 border border-slate-200 px-4 py-3 text-sm font-bold flex items-center justify-center gap-2">
                    <i class="fa-solid fa-triangle-exclamation text-amber-500"></i><span>재고 위험관리</span><span id="orderSubRiskCount" class="rounded-full bg-red-100 text-red-700 text-xs px-2 py-0.5">0</span>
                </button>
            </div>
            <div id="order-management-panel" class="space-y-4">
            <div class="rounded-xl border border-emerald-200 bg-emerald-50 px-4 py-3">
                <h2 class="text-lg font-extrabold text-emerald-900"><i class="fa-solid fa-cart-shopping mr-2"></i>1. 오늘·내일 발주 관리</h2>
                <p class="mt-1 text-sm text-emerald-800">오늘·내일 필요한 발주 수량을 확인하고 입력하세요. 재고 위험은 위쪽의 ‘재고 위험관리’ 메뉴에서 따로 확인할 수 있습니다.</p>
            </div>
            <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-col md:flex-row items-center justify-between gap-4">
                <div class="flex items-center space-x-4">
                    <div class="bg-emerald-100 text-emerald-800 px-3 py-2 rounded-lg font-bold text-sm">
                        오늘 발주 합계: <span id="summaryTodayOrderAmount" class="text-emerald-700 text-base">0원</span>
                        <span id="summaryTodayOrderQty" class="text-xs text-slate-500 font-normal ml-1">(0개)</span>
                    </div>
                    <div class="bg-blue-100 text-blue-800 px-3 py-2 rounded-lg font-bold text-sm">
                        내일 발주 합계: <span id="summaryTomorrowOrderAmount" class="text-blue-700 text-base">0원</span>
                        <span id="summaryTomorrowOrderQty" class="text-xs text-slate-500 font-normal ml-1">(0개)</span>
                    </div>
                </div>
                <div class="flex flex-wrap items-center gap-2">
                    <button onclick="applyAllAiOrders()" class="bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 px-3 py-1.5 rounded-lg text-xs font-bold transition flex items-center space-x-1">
                        <i class="fa-solid fa-wand-magic-sparkles"></i>
                        <span>AI 추천수량 전체적용</span>
                    </button>
                    <button onclick="downloadOrderSlip('today','excel')" class="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-lg text-xs font-bold transition flex items-center gap-1">
                        <i class="fa-solid fa-file-excel"></i><span>오늘 발주서 Excel</span>
                    </button>
                    <button onclick="downloadOrderSlip('today','pdf')" class="bg-slate-700 hover:bg-slate-800 text-white px-3 py-1.5 rounded-lg text-xs font-bold transition flex items-center gap-1">
                        <i class="fa-solid fa-file-pdf"></i><span>오늘 발주서 PDF</span>
                    </button>
                    <button onclick="downloadOrderSlip('tomorrow','excel')" class="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded-lg text-xs font-bold transition flex items-center gap-1">
                        <i class="fa-solid fa-file-excel"></i><span>내일 발주서 Excel</span>
                    </button>
                    <button onclick="downloadOrderSlip('tomorrow','pdf')" class="bg-slate-700 hover:bg-slate-800 text-white px-3 py-1.5 rounded-lg text-xs font-bold transition flex items-center gap-1">
                        <i class="fa-solid fa-file-pdf"></i><span>내일 발주서 PDF</span>
                    </button>
                </div>
            </div>

            <div class="flex items-center justify-between bg-slate-100 p-2.5 rounded-lg text-xs">
                <div class="flex items-center space-x-2">
                    <span class="font-bold text-slate-600">보기 필터:</span>
                    <button onclick="filterOrderTable('ALL')" id="orderFilterBtn-ALL" class="bg-white shadow-sm text-slate-800 font-semibold px-2.5 py-1 rounded">전체 (106개)</button>
                    <button onclick="filterOrderTable('ORDER_ONLY')" id="orderFilterBtn-ORDER_ONLY" class="text-slate-600 font-semibold px-2.5 py-1 rounded hover:bg-white">발주 필요품목만 (추천수량 기준)</button>
                    <button onclick="filterOrderTable('FF_ONLY')" id="orderFilterBtn-FF_ONLY" class="text-slate-600 font-semibold px-2.5 py-1 rounded hover:bg-white">FF/신선식품만</button>
                </div>
                <div class="text-slate-500">
                    * <strong>기말재고 계산식</strong> = 기초재고 + 오늘발주수량 - 오늘예측수요 (수정 시 실시간 재계산)
                </div>
            </div>

            <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                <div class="overflow-x-auto custom-scrollbar">
                    <table class="w-full text-left text-xs sm:text-sm">
                        <thead class="bg-slate-800 text-white font-semibold">
                            <tr>
                                <th class="p-3">상품명 / 코드</th>
                                <th class="p-3 text-center">품목 상태</th>
                                <th class="p-3 text-center">원가 / 매가</th>
                                <th class="p-3 text-center bg-slate-700">기초재고</th>
                                <th class="p-3 text-center text-emerald-300">안전재고(SS)</th>
                                <th class="p-3 text-center text-amber-300">ROP</th>
                                <th class="p-3 text-center bg-indigo-900 text-indigo-100">목표재고</th>
                                <th class="p-3 text-center bg-emerald-900 text-emerald-100">오늘 AI추천</th>
                                <th class="p-3 text-center bg-emerald-900 text-emerald-100 w-28">오늘 발주수량</th>
                                <th class="p-3 text-center bg-cyan-900 text-cyan-200 font-bold">기말재고(예정)</th>
                                <th class="p-3 text-center bg-blue-900 text-blue-100">내일 AI추천</th>
                                <th class="p-3 text-center bg-blue-900 text-blue-100 w-28">내일 발주수량</th>
                                <th class="p-3 text-right">오늘 발주금액</th>
                            </tr>
                        </thead>
                        <tbody id="orderTableBody" class="divide-y divide-slate-100">
                            <!-- Populated via JS -->
                        </tbody>
                    </table>
                </div>
            </div>

            </div><!-- /order-management-panel -->

            <!-- 재고 위험관리는 발주와 분리된 내부 메뉴로 표시 -->
            <div id="risk-management-panel" class="hidden space-y-4">
                <div class="rounded-xl border border-amber-200 bg-amber-50 px-4 py-3">
                    <h2 class="text-lg font-extrabold text-amber-900"><i class="fa-solid fa-triangle-exclamation text-amber-600 mr-2"></i>재고 위험관리 및 원인 진단</h2>
                    <p class="mt-1 text-sm text-amber-800">품절·폐기·과다재고 위험을 확인하고, 필요한 조치를 실행하세요. 발주 화면과 분리되어 있어 발주 작업을 방해하지 않습니다.</p>
                </div>
        <div id="risk-analysis-panel" class="space-y-4">
            <div class="grid grid-cols-2 md:grid-cols-5 gap-3">
                <button onclick="filterRiskCategory('ALL')" id="riskFilter-ALL" class="p-3 rounded-xl bg-white border border-slate-200 text-left hover:border-slate-400 transition shadow-sm">
                    <div class="text-xs font-semibold text-slate-500">전체 품목</div>
                    <div id="riskBadgeCountAll" class="text-xl font-bold text-slate-800 mt-1">106</div>
                </button>
                <button onclick="filterRiskCategory('STOCKOUT')" id="riskFilter-STOCKOUT" class="p-3 rounded-xl bg-red-50 border border-red-200 text-left hover:border-red-400 transition shadow-sm">
                    <div class="text-xs font-semibold text-red-700 flex items-center justify-between">
                        <span>🚨 품절위험</span>
                        <i class="fa-solid fa-triangle-exclamation"></i>
                    </div>
                    <div id="riskBadgeCountStockout" class="text-xl font-bold text-red-700 mt-1">0</div>
                </button>
                <button onclick="filterRiskCategory('WASTE')" id="riskFilter-WASTE" class="p-3 rounded-xl bg-orange-50 border border-orange-200 text-left hover:border-orange-400 transition shadow-sm">
                    <div class="text-xs font-semibold text-orange-700 flex items-center justify-between">
                        <span>🥪 폐기위험</span>
                        <i class="fa-solid fa-clock-rotate-left"></i>
                    </div>
                    <div id="riskBadgeCountWaste" class="text-xl font-bold text-orange-700 mt-1">0</div>
                </button>
                <button onclick="filterRiskCategory('OVERSTOCK')" id="riskFilter-OVERSTOCK" class="p-3 rounded-xl bg-yellow-50 border border-yellow-200 text-left hover:border-yellow-400 transition shadow-sm">
                    <div class="text-xs font-semibold text-yellow-800 flex items-center justify-between">
                        <span>⚠️ 과다재고</span>
                        <i class="fa-solid fa-boxes-stacked"></i>
                    </div>
                    <div id="riskBadgeCountOverstock" class="text-xl font-bold text-yellow-800 mt-1">0</div>
                </button>
                <button onclick="filterRiskCategory('NORMAL')" id="riskFilter-NORMAL" class="p-3 rounded-xl bg-emerald-50 border border-emerald-200 text-left hover:border-emerald-400 transition shadow-sm">
                    <div class="text-xs font-semibold text-emerald-700 flex items-center justify-between">
                        <span>✅ 정상재고</span>
                        <i class="fa-solid fa-circle-check"></i>
                    </div>
                    <div id="riskBadgeCountNormal" class="text-xl font-bold text-emerald-700 mt-1">0</div>
                </button>
            </div>

            <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                <div class="p-3 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
                    <h3 class="text-sm font-bold text-slate-800">재고 위험 감지 및 기초/안전/기말재고 대비 자동 진단</h3>
                    <span class="text-xs text-slate-500">원터치조치로 재고 상태를 즉시 조절할 수 있습니다.</span>
                </div>
                <div class="overflow-x-auto custom-scrollbar">
                    <table class="w-full text-left text-xs sm:text-sm">
                        <thead class="bg-slate-800 text-white font-semibold">
                            <tr>
                                <th class="p-3">위험 상태</th>
                                <th class="p-3">상품명 / 카테고리</th>
                                <th class="p-3 text-center">기초재고</th>
                                <th class="p-3 text-center text-emerald-300">안전재고(SS)</th>
                                <th class="p-3 text-center text-cyan-300">기말재고(예정)</th>
                                <th class="p-3 text-center">목표재고</th>
                                <th class="p-3 text-center">재발주점(ROP)</th>
                                <th class="p-3">원인 및 진단 메시지</th>
                                <th class="p-3 text-center">AI 추천 대응 조치</th>
                                <th class="p-3 text-center">원터치 실행</th>
                            </tr>
                        </thead>
                        <tbody id="riskTableBody" class="divide-y divide-slate-100">
                            <!-- Populated via JS -->
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
            </div><!-- /risk-management-panel -->

        </section>

        <!-- TAB 5: POS 데이터 자동생성 시뮬레이터 -->
        <section id="tab-simulator" class="tab-content hidden space-y-6">
            <div class="bg-slate-900 text-white p-6 rounded-2xl shadow-lg border border-slate-800">
                <div class="flex flex-col lg:flex-row items-center justify-between gap-6">
                    <div class="space-y-2">
                        <div class="flex items-center space-x-2">
                            <span class="bg-yellow-400 text-slate-900 text-xs font-black px-2 py-0.5 rounded">10월 실습용 기능</span>
                            <h2 class="text-xl font-bold">실시간 POS 판매 및 재고/손익 시뮬레이터</h2>
                        </div>
                        <p class="text-xs text-slate-300 max-w-2xl">
                            다음 날로 진행 버튼을 클릭하면 실제 점포 POS와 같이 오늘 판매, 입고, 유통기한 폐기가 변동 생성되며, 학생 수동 발주와 AI 발주 간의 손익을 비교해줍니다.
                        </p>
                    </div>

                    <div class="flex flex-wrap items-center gap-3">
                        <button onclick="executeDailySimulation()" class="bg-blue-600 hover:bg-blue-500 text-white font-bold px-5 py-3 rounded-xl shadow-lg transition flex items-center space-x-2 border border-blue-400/30">
                            <i class="fa-solid fa-forward-step text-lg"></i>
                            <span>다음 날로 진행 (영업 마감 & 손익 분석)</span>
                        </button>
                        <button onclick="resetSimulationData()" class="bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white font-medium px-4 py-3 rounded-xl border border-slate-700 transition text-xs">
                            <i class="fa-solid fa-rotate-left mr-1"></i> 데이터 완전 초기화
                        </button>
                    </div>
                </div>

                <div class="grid grid-cols-2 sm:grid-cols-4 gap-4 mt-6 pt-6 border-t border-slate-800 text-center">
                    <div class="bg-slate-800/60 p-3 rounded-xl border border-slate-700/50">
                        <span class="text-slate-400 text-xs">누적 시뮬레이션 일수</span>
                        <div id="simDaysCount" class="text-xl font-bold text-white mt-1">0 일차</div>
                    </div>
                    <div class="bg-slate-800/60 p-3 rounded-xl border border-slate-700/50">
                        <span class="text-slate-400 text-xs">학생 발주 누적 매출액</span>
                        <div id="simTotalRevenue" class="text-xl font-bold text-emerald-400 mt-1">0 원</div>
                    </div>
                    <div class="bg-slate-800/60 p-3 rounded-xl border border-slate-700/50">
                        <span class="text-slate-400 text-xs">학생 발주 누적 폐기 발생액</span>
                        <div id="simTotalWaste" class="text-xl font-bold text-orange-400 mt-1">0 원</div>
                    </div>
                    <div class="bg-slate-800/60 p-3 rounded-xl border border-slate-700/50">
                        <span class="text-slate-400 text-xs">총 발주 처리건수</span>
                        <div id="simTotalOrdersCount" class="text-xl font-bold text-cyan-400 mt-1">0 건</div>
                    </div>
                </div>
            </div>

            <!-- 누적 보고서: POS 시뮬레이터의 일자별 마감 기록을 합산 -->
            <section id="cumulativeReportSection" class="bg-white p-5 rounded-xl shadow-sm border border-slate-200 space-y-4" aria-labelledby="cumulativeReportTitle">
                <div class="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-2 border-b border-slate-100 pb-3">
                    <div>
                        <h3 id="cumulativeReportTitle" class="text-base font-bold text-slate-800"><i class="fa-solid fa-chart-column text-indigo-600 mr-2"></i>누적 영업 성과 보고서</h3>
                        <p class="text-xs text-slate-500 mt-1">마감한 날짜의 실적을 자동 합산합니다. 다음 날로 진행할 때마다 갱신됩니다.</p>
                    </div>
                    <div class="flex flex-wrap items-center gap-2">
                        <span id="cumulativeReportPeriod" class="text-xs font-semibold bg-slate-100 text-slate-600 rounded-full px-3 py-1">기록 0일</span>
                        <button type="button" onclick="exportCumulativeExcel()" class="inline-flex items-center gap-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-bold px-3 py-2 transition"><i class="fa-solid fa-file-excel"></i> Excel 저장</button>
                        <button type="button" onclick="exportCumulativePDF()" class="inline-flex items-center gap-2 rounded-lg bg-rose-600 hover:bg-rose-700 text-white text-xs font-bold px-3 py-2 transition"><i class="fa-solid fa-file-pdf"></i> PDF 저장</button>
                    </div>
                </div>
                <div class="grid grid-cols-2 lg:grid-cols-5 gap-3">
                    <div class="rounded-lg bg-emerald-50 border border-emerald-100 p-3"><div class="text-xs text-emerald-700">학생 누적 매출</div><div id="cumStudentRevenue" class="text-base sm:text-lg font-extrabold text-emerald-800 mt-1">0원</div></div>
                    <div class="rounded-lg bg-blue-50 border border-blue-100 p-3"><div class="text-xs text-blue-700">AI 누적 매출</div><div id="cumAiRevenue" class="text-base sm:text-lg font-extrabold text-blue-800 mt-1">0원</div></div>
                    <div class="rounded-lg bg-orange-50 border border-orange-100 p-3"><div class="text-xs text-orange-700">학생 누적 폐기손실</div><div id="cumStudentWaste" class="text-base sm:text-lg font-extrabold text-orange-800 mt-1">0원</div></div>
                    <div class="rounded-lg bg-rose-50 border border-rose-100 p-3"><div class="text-xs text-rose-700">학생 누적 품절손실</div><div id="cumStudentStockout" class="text-base sm:text-lg font-extrabold text-rose-800 mt-1">0원</div></div>
                    <div class="rounded-lg bg-indigo-50 border border-indigo-100 p-3"><div class="text-xs text-indigo-700">AI 순이익 우위</div><div id="cumNetDifference" class="text-base sm:text-lg font-extrabold text-indigo-800 mt-1">0원</div></div>
                </div>
                <div id="cumulativeReportConclusion" class="rounded-lg bg-slate-50 border border-slate-200 p-3 text-sm text-slate-700">아직 마감 기록이 없습니다. ‘다음 날로 진행’을 누르면 누적 보고서가 생성됩니다.</div>
            </section>

            <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <!-- 환경 설정 -->
                <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200 space-y-4">
                    <h3 class="text-sm font-bold text-slate-800 border-b pb-2 flex items-center justify-between">
                        <span><i class="fa-solid fa-cloud-sun text-amber-500 mr-2"></i>외부 환경 요인 설정 (수요 변동)</span>
                    </h3>
                    <div class="space-y-3 text-xs">
                        <div>
                            <label class="font-semibold text-slate-700 block mb-1">오늘 날씨 효과</label>
                            <select id="simWeatherSelect" class="w-full p-2 bg-slate-50 border rounded-lg">
                                <option value="NORMAL">☀️ 맑음 (평년 수요)</option>
                                <option value="RAIN">🌧️ 비/폭우 (수요 감소, 컵라면만 증가)</option>
                                <option value="HOT">🔥 찌는듯한 폭염 (음료/아이스크림 수요 급증)</option>
                                <option value="COLD">❄️ 기온 급강하 (전체 수요 약간 증가)</option>
                            </select>
                        </div>
                        <div>
                            <label class="font-semibold text-slate-700 block mb-1">상권 이벤트 특이사항</label>
                            <select id="simEventSelect" class="w-full p-2 bg-slate-50 border rounded-lg">
                                <option value="NONE">일반 영업일</option>
                                <option value="NEARBY_EVENT">인근 행사 개최 (수요 대폭 증가)</option>
                            </select>
                        </div>
                    </div>
                    <div class="bg-blue-50 p-3 rounded-lg border border-blue-100 text-xs text-blue-800 space-y-1 mt-4">
                        <div class="font-bold"><i class="fa-solid fa-circle-info mr-1"></i>안내</div>
                        <p class="text-blue-700">마감을 진행하면 아래 표에 누적 기록되며, <strong>'상세보기'</strong>를 통해 106개 제품별 성과를 볼 수 있습니다.</p>
                        <p class="text-purple-700 mt-1 font-semibold">판매예측 탭에서 '실제판매(수동)'을 기입하면 해당 판매량이 강제로 우선 적용됩니다.</p>
                    </div>
                </div>

                <!-- 일자별 영업 마감 데이터 조회 테이블 -->
                <div class="lg:col-span-2 bg-white p-5 rounded-xl shadow-sm border border-slate-200 flex flex-col h-[400px]">
                    <div class="flex items-center justify-between border-b pb-2 mb-3">
                        <h3 class="text-sm font-bold text-slate-800 flex items-center">
                            <i class="fa-solid fa-folder-open text-indigo-600 mr-2"></i>
                            일자별 영업 마감 데이터 조회 (학생 vs AI)
                        </h3>
                    </div>
                    
                    <div class="flex-grow overflow-y-auto custom-scrollbar">
                        <table class="w-full text-left text-xs">
                            <thead class="bg-slate-100 text-slate-600 sticky top-0 border-b border-slate-200 shadow-sm">
                                <tr>
                                    <th class="p-2 w-20">영업일자</th>
                                    <th class="p-2 text-right border-l">학생 순이익</th>
                                    <th class="p-2 text-right border-l">AI 순이익</th>
                                    <th class="p-2 text-center text-slate-700 bg-slate-50 border-l">최종 결과</th>
                                    <th class="p-2 text-center border-l w-20">상세조회</th>
                                </tr>
                            </thead>
                            <tbody id="closingHistoryBody" class="divide-y divide-slate-100">
                                <!-- 데이터가 여기에 누적됩니다 -->
                                <tr>
                                    <td colspan="5" class="p-8 text-center text-slate-400">아직 영업 마감 데이터가 없습니다. '다음 날로 진행' 버튼을 눌러주세요.</td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </section>

    </main>

    <!-- Daily Closing Report Modal (Aggregate) -->
    <div id="closingModal" class="fixed inset-0 z-[60] bg-slate-900/70 backdrop-blur-sm hidden flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-2xl w-full overflow-hidden border border-slate-200 transform transition-all">
            <div class="bg-indigo-600 text-white p-4 flex items-center justify-between">
                <h3 class="text-lg font-bold flex items-center"><i class="fa-solid fa-file-invoice-dollar mr-2"></i>일일 영업 마감 종합 리포트</h3>
                <button onclick="moveToNextDay(false)" class="text-indigo-200 hover:text-white"><i class="fa-solid fa-xmark text-xl"></i></button>
            </div>
            <div class="p-5">
                <div class="text-center mb-4">
                    <h4 id="closingDateLabel" class="text-xl font-bold text-slate-800">2026-10-01 마감 결과</h4>
                    <p class="text-sm text-slate-500">오늘 하루 실제 수요 발생에 따른 발주 및 재고 관리 결과야!</p>
                </div>
                
                <div class="grid grid-cols-2 gap-4">
                    <div class="bg-slate-50 border border-slate-200 p-4 rounded-xl shadow-sm">
                        <h5 class="font-bold text-slate-700 mb-3 text-center border-b pb-2">🧑‍🎓 학생 (수동 발주)</h5>
                        <div class="space-y-2 text-sm">
                            <div class="flex justify-between"><span class="text-slate-500">실제 매출액</span><strong class="text-emerald-600" id="c_stuRev">0원</strong></div>
                            <div class="flex justify-between"><span class="text-slate-500">폐기 손실(-)</span><strong class="text-orange-600" id="c_stuWaste">0원</strong></div>
                            <div class="flex justify-between"><span class="text-slate-500">품절 기회손실(-)</span><strong class="text-red-600" id="c_stuStockout">0원</strong></div>
                            <div class="flex justify-between pt-2 border-t mt-2"><span class="font-bold text-slate-800">일일 순이익</span><strong class="text-blue-700 text-lg" id="c_stuNet">0원</strong></div>
                        </div>
                    </div>
                    <div class="bg-blue-50 border border-blue-200 p-4 rounded-xl shadow-sm">
                        <h5 class="font-bold text-blue-800 mb-3 text-center border-b border-blue-200 pb-2">🤖 AI (자동 추천 발주)</h5>
                        <div class="space-y-2 text-sm">
                            <div class="flex justify-between"><span class="text-blue-600/80">실제 매출액</span><strong class="text-emerald-600" id="c_aiRev">0원</strong></div>
                            <div class="flex justify-between"><span class="text-blue-600/80">폐기 손실(-)</span><strong class="text-orange-600" id="c_aiWaste">0원</strong></div>
                            <div class="flex justify-between"><span class="text-blue-600/80">품절 기회손실(-)</span><strong class="text-red-600" id="c_aiStockout">0원</strong></div>
                            <div class="flex justify-between pt-2 border-t border-blue-200 mt-2"><span class="font-bold text-blue-900">일일 순이익</span><strong class="text-blue-700 text-lg" id="c_aiNet">0원</strong></div>
                        </div>
                    </div>
                </div>

                <div class="mt-4 p-3 bg-indigo-50 text-indigo-800 rounded-lg text-sm text-center font-medium" id="c_conclusion">
                    AI가 학생보다 N원 더 많은 이익을 냈어!
                </div>
            </div>
            <div class="bg-slate-100 p-4 flex justify-between items-center">
                <span class="text-xs text-slate-500">* 이 기록은 'POS 데이터' 탭에 누적되며 제품별 조회도 가능합니다.</span>
                <button onclick="moveToNextDay(false)" class="bg-indigo-600 hover:bg-indigo-700 text-white px-6 py-2.5 rounded-lg font-bold shadow-md transition flex items-center">
                    다음 날로 넘어가기 (기초재고 이월) <i class="fa-solid fa-arrow-right ml-2"></i>
                </button>
            </div>
        </div>
    </div>

    <!-- 일자별 106개 제품 마감 상세 팝업 -->
    <div id="dailyDetailModal" class="fixed inset-0 z-[65] bg-slate-900/80 backdrop-blur-sm hidden flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-5xl w-full h-[85vh] flex flex-col overflow-hidden border border-slate-200 transform transition-all">
            <div class="bg-slate-800 text-white p-4 flex items-center justify-between">
                <h3 class="text-lg font-bold flex items-center"><i class="fa-solid fa-list-ol mr-2 text-blue-400"></i> <span id="detailModalTitle">2026-10-01 제품별 마감 상세 리포트</span></h3>
                <button onclick="closeDailyDetailModal()" class="text-slate-400 hover:text-white text-xl p-1"><i class="fa-solid fa-xmark"></i></button>
            </div>
            <div class="flex-grow overflow-y-auto custom-scrollbar p-1">
                <table class="w-full text-left text-xs">
                    <thead class="bg-slate-100 text-slate-600 sticky top-0 border-b border-slate-200 shadow-sm">
                        <tr>
                            <th class="p-2 w-32">상품명 / 코드</th>
                            <th class="p-2 text-center">등급</th>
                            <th class="p-2 text-right bg-blue-50/50 border-l">학생 판매</th>
                            <th class="p-2 text-right bg-blue-50/50">AI 판매</th>
                            <th class="p-2 text-right bg-orange-50/50 border-l">학생 폐기</th>
                            <th class="p-2 text-right bg-orange-50/50">AI 폐기</th>
                            <th class="p-2 text-right bg-red-50/50 border-l">학생 품절손실</th>
                            <th class="p-2 text-right bg-red-50/50">AI 품절손실</th>
                            <th class="p-2 text-right text-indigo-700 font-bold bg-indigo-50 border-l">학생 순이익</th>
                            <th class="p-2 text-right text-indigo-700 font-bold bg-indigo-50">AI 순이익</th>
                        </tr>
                    </thead>
                    <tbody id="dailyDetailBody" class="divide-y divide-slate-100">
                        <!-- 제품별 상세 데이터 렌더링 -->
                    </tbody>
                </table>
            </div>
            <div class="bg-slate-100 p-3 flex flex-wrap justify-between gap-2 items-center border-t">
                <div class="flex gap-2 flex-wrap">
                    <button onclick="exportSelectedDayExcel()" class="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg text-sm font-bold"><i class="fa-solid fa-file-excel mr-1"></i> 이 날짜 Excel 저장</button>
                    <button onclick="exportSelectedDayPDF()" class="bg-rose-600 hover:bg-rose-700 text-white px-4 py-2 rounded-lg text-sm font-bold"><i class="fa-solid fa-file-pdf mr-1"></i> 이 날짜 PDF 저장</button>
                </div>
                <button onclick="closeDailyDetailModal()" class="bg-slate-800 text-white px-5 py-2 rounded-lg text-sm font-bold">닫기</button>
            </div>
        </div>
    </div>

    <!-- SKU Detail Modal -->
    <div id="skuModal" class="fixed inset-0 z-[70] bg-slate-900/60 backdrop-blur-sm hidden flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-2xl max-w-lg w-full overflow-hidden border border-slate-200 transform transition-all">
            <div class="bg-slate-900 text-white p-4 flex items-center justify-between">
                <div>
                    <span id="modalSkuCode" class="text-xs font-mono bg-blue-600 px-2 py-0.5 rounded text-blue-100">SKU001</span>
                    <h3 id="modalSkuName" class="text-base font-bold mt-1 text-white">상품명</h3>
                </div>
                <button onclick="closeSkuModal()" class="text-slate-400 hover:text-white text-xl p-1"><i class="fa-solid fa-xmark"></i></button>
            </div>
            <div class="p-5 space-y-4 text-xs">
                <div class="grid grid-cols-2 gap-3 bg-slate-50 p-3 rounded-xl border border-slate-200">
                    <div><span class="text-slate-500">카테고리:</span> <strong id="modalCat" class="text-slate-800"></strong></div>
                    <div><span class="text-slate-500">ABC 등급:</span> <strong id="modalAbc" class="text-blue-700 font-bold"></strong></div>
                    <div><span class="text-slate-500">판매가 / 원가:</span> <strong id="modalPriceCost" class="text-slate-800"></strong></div>
                    <div><span class="text-slate-500">소비기한 / 보관:</span> <strong id="modalShelfStorage" class="text-slate-800"></strong></div>
                    <div><span class="text-slate-500">최소진열량:</span> <strong id="modalMinDisplay" class="text-slate-800"></strong></div>
                </div>
                <div class="border-t pt-3 space-y-2">
                    <h4 class="font-bold text-slate-800 flex items-center"><i class="fa-solid fa-chart-bar text-indigo-600 mr-1.5"></i>재고 상태 및 기준값</h4>
                    <div class="grid grid-cols-3 gap-2 text-slate-700">
                        <div class="p-2 bg-slate-100 rounded border border-slate-200">
                            <div class="text-slate-500">기초재고 (Start)</div>
                            <div id="modalBeginningStock" class="font-bold text-sm text-slate-800">0개</div>
                        </div>
                        <div class="p-2 bg-emerald-50/80 rounded border border-emerald-200">
                            <div class="text-slate-500">안전재고 (SS)</div>
                            <div id="modalSafetyStock" class="font-bold text-sm text-emerald-800">0개</div>
                        </div>
                        <div class="p-2 bg-cyan-50/80 rounded border border-cyan-200">
                            <div class="text-slate-500">기말재고 (예정)</div>
                            <div id="modalEndingStock" class="font-bold text-sm text-cyan-800">0개</div>
                        </div>
                        <div class="p-2 bg-amber-50/50 rounded border border-amber-100">
                            <div class="text-slate-500">재발주점 (ROP)</div>
                            <div id="modalRop" class="font-bold text-sm text-amber-800">0개</div>
                        </div>
                        <div class="p-2 bg-purple-50/50 rounded border border-purple-100">
                            <div class="text-slate-500">목표재고 (Target)</div>
                            <div id="modalTargetStock" class="font-bold text-sm text-purple-800">0개</div>
                        </div>
                    </div>
                </div>
            </div>
            <div class="bg-slate-100 p-3 text-right">
                <button onclick="closeSkuModal()" class="bg-slate-800 text-white px-4 py-1.5 rounded-lg text-xs font-bold">확인 닫기</button>
            </div>
        </div>
    </div>

    <!-- Toast Notification Message Box -->
    <div id="toastMessage" class="fixed bottom-5 right-5 z-50 transform translate-y-20 opacity-0 transition-all duration-300 bg-slate-900 text-white px-4 py-3 rounded-xl shadow-2xl flex items-center space-x-3 border border-slate-700">
        <i id="toastIcon" class="fa-solid fa-circle-check text-emerald-400 text-lg"></i>
        <span id="toastText" class="text-xs font-semibold">작업이 완료되었습니다.</span>
    </div>

    <script>
        const STATE = {
            currentDateStr: '2026-10-01',
            currentDateObj: new Date(2026, 9, 1),
            simDaysElapsed: 0,
            simTotalRevenue: 0,
            simTotalWaste: 0,
            simTotalOrdersCount: 0,
            activeTab: 'order',
            riskFilter: 'ALL',
            orderFilter: 'ALL',
            items: [],
            closingHistory: [], // 일자별 마감 내역 저장 배열
            charts: {}
        };

        const CATEGORIES = [
            'FF/즉석식', '음료', '커피/에너지', '유제품',
            '라면/즉석조리', '과자/초콜릿/젤리', '아이스크림/냉동', '생활/위생'
        ];

        // 10개의 진짜 상품 데이터 복구!
        const POS_106_DATA = [
            {"id": "SKU001", "name": "혜자로운집밥 제육볶음 도시락", "category": "FF/즉석식", "subCat": "도시락", "storage": "냉장", "shelfLifeDays": 2, "price": 5300, "cost": 3570, "packSize": 5, "minDisplay": 4, "stockCoverDays": 1.0, "abc": "A", "avgSales": 9.9, "stdDev": 2.3, "totalSalesQty": 2716, "totalSalesAmt": 14394800, "totalWasteQty": 4, "totalWasteLoss": 14280, "totalStockoutQty": 0, "totalStockoutLoss": 0, "stock": 17, "safetyStock": 8, "rop": 18, "targetStock": 30, "forecastD0": 10.6, "forecastD1": 11.1, "forecastD2": 8.7, "forecastD3": 8.3, "orderToday": 15, "orderTomorrow": 10, "userOrderToday": 15, "userOrderTomorrow": 10},
            {"id": "SKU002", "name": "혜자로운집밥 오징어불고기 도시락", "category": "FF/즉석식", "subCat": "도시락", "storage": "냉장", "shelfLifeDays": 2, "price": 6100, "cost": 4170, "packSize": 5, "minDisplay": 3, "stockCoverDays": 1.2, "abc": "A", "avgSales": 7.9, "stdDev": 1.6, "totalSalesQty": 2169, "totalSalesAmt": 13230900, "totalWasteQty": 5, "totalWasteLoss": 20850, "totalStockoutQty": 0, "totalStockoutLoss": 0, "stock": 19, "safetyStock": 6, "rop": 14, "targetStock": 23, "forecastD0": 8.2, "forecastD1": 8.6, "forecastD2": 7.2, "forecastD3": 6.8, "orderToday": 0, "orderTomorrow": 5, "userOrderToday": 0, "userOrderTomorrow": 5},
            {"id": "SKU003", "name": "김혜자 참치마요 주먹밥", "category": "FF/즉석식", "subCat": "주먹밥", "storage": "냉장", "shelfLifeDays": 2, "price": 1900, "cost": 1290, "packSize": 5, "minDisplay": 3, "stockCoverDays": 1.1, "abc": "A", "avgSales": 7.9, "stdDev": 1.8, "totalSalesQty": 2149, "totalSalesAmt": 4083100, "totalWasteQty": 4, "totalWasteLoss": 5160, "totalStockoutQty": 0, "totalStockoutLoss": 0, "stock": 16, "safetyStock": 6, "rop": 14, "targetStock": 23, "forecastD0": 8.1, "forecastD1": 8.3, "forecastD2": 7.3, "forecastD3": 7.0, "orderToday": 10, "orderTomorrow": 10, "userOrderToday": 10, "userOrderTomorrow": 10},
            {"id": "SKU004", "name": "전주비빔 삼각김밥", "category": "FF/즉석식", "subCat": "주먹밥", "storage": "냉장", "shelfLifeDays": 2, "price": 1300, "cost": 890, "packSize": 5, "minDisplay": 5, "stockCoverDays": 1.0, "abc": "A", "avgSales": 15.2, "stdDev": 2.9, "totalSalesQty": 4150, "totalSalesAmt": 5395000, "totalWasteQty": 12, "totalWasteLoss": 10680, "totalStockoutQty": 0, "totalStockoutLoss": 0, "stock": 22, "safetyStock": 10, "rop": 26, "targetStock": 43, "forecastD0": 16.1, "forecastD1": 16.8, "forecastD2": 14.1, "forecastD3": 13.5, "orderToday": 25, "orderTomorrow": 20, "userOrderToday": 25, "userOrderTomorrow": 20},
            {"id": "SKU005", "name": "케이준치킨 샌드위치", "category": "FF/즉석식", "subCat": "샌드위치", "storage": "냉장", "shelfLifeDays": 3, "price": 3200, "cost": 2180, "packSize": 2, "minDisplay": 2, "stockCoverDays": 1.2, "abc": "B", "avgSales": 5.4, "stdDev": 1.2, "totalSalesQty": 1475, "totalSalesAmt": 4720000, "totalWasteQty": 8, "totalWasteLoss": 17440, "totalStockoutQty": 0, "totalStockoutLoss": 0, "stock": 14, "safetyStock": 4, "rop": 10, "targetStock": 16, "forecastD0": 5.6, "forecastD1": 5.8, "forecastD2": 5.0, "forecastD3": 4.8, "orderToday": 0, "orderTomorrow": 2, "userOrderToday": 0, "userOrderTomorrow": 2},
            {"id": "SKU006", "name": "코카콜라 500ml", "category": "음료", "subCat": "탄산음료", "storage": "상온", "shelfLifeDays": 180, "price": 2300, "cost": 1560, "packSize": 12, "minDisplay": 10, "stockCoverDays": 2.0, "abc": "A", "avgSales": 18.5, "stdDev": 3.4, "totalSalesQty": 5050, "totalSalesAmt": 11615000, "totalWasteQty": 0, "totalWasteLoss": 0, "totalStockoutQty": 0, "totalStockoutLoss": 0, "stock": 42, "safetyStock": 16, "rop": 35, "targetStock": 56, "forecastD0": 19.2, "forecastD1": 20.1, "forecastD2": 17.5, "forecastD3": 16.9, "orderToday": 24, "orderTomorrow": 24, "userOrderToday": 24, "userOrderTomorrow": 24},
            {"id": "SKU007", "name": "삼다수 2L", "category": "음료", "subCat": "생수", "storage": "상온", "shelfLifeDays": 365, "price": 1200, "cost": 820, "packSize": 6, "minDisplay": 12, "stockCoverDays": 2.5, "abc": "A", "avgSales": 24.1, "stdDev": 4.1, "totalSalesQty": 6580, "totalSalesAmt": 7896000, "totalWasteQty": 0, "totalWasteLoss": 0, "totalStockoutQty": 0, "totalStockoutLoss": 0, "stock": 95, "safetyStock": 19, "rop": 44, "targetStock": 71, "forecastD0": 25.0, "forecastD1": 26.2, "forecastD2": 22.8, "forecastD3": 21.9, "orderToday": 0, "orderTomorrow": 0, "userOrderToday": 0, "userOrderTomorrow": 0},
            {"id": "SKU008", "name": "TOP 아메리카노 275ml", "category": "커피/에너지", "subCat": "캔커피", "storage": "상온", "shelfLifeDays": 180, "price": 2400, "cost": 1630, "packSize": 10, "minDisplay": 8, "stockCoverDays": 2.0, "abc": "A", "avgSales": 14.8, "stdDev": 2.8, "totalSalesQty": 4040, "totalSalesAmt": 9696000, "totalWasteQty": 0, "totalWasteLoss": 0, "totalStockoutQty": 0, "totalStockoutLoss": 0, "stock": 28, "safetyStock": 13, "rop": 28, "targetStock": 45, "forecastD0": 15.3, "forecastD1": 16.0, "forecastD2": 14.0, "forecastD3": 13.5, "orderToday": 20, "orderTomorrow": 20, "userOrderToday": 20, "userOrderTomorrow": 20},
            {"id": "SKU009", "name": "빙그레 바나나맛우유 240ml", "category": "유제품", "subCat": "가공유", "storage": "냉장", "shelfLifeDays": 10, "price": 1700, "cost": 1150, "packSize": 6, "minDisplay": 10, "stockCoverDays": 1.2, "abc": "A", "avgSales": 22.4, "stdDev": 3.9, "totalSalesQty": 6115, "totalSalesAmt": 10395500, "totalWasteQty": 2, "totalWasteLoss": 2300, "totalStockoutQty": 0, "totalStockoutLoss": 0, "stock": 35, "safetyStock": 17, "rop": 40, "targetStock": 65, "forecastD0": 23.2, "forecastD1": 24.1, "forecastD2": 21.0, "forecastD3": 20.3, "orderToday": 30, "orderTomorrow": 24, "userOrderToday": 30, "userOrderTomorrow": 24},
            {"id": "SKU010", "name": "신라면 큰사발 114g", "category": "라면/즉석조리", "subCat": "용기면", "storage": "상온", "shelfLifeDays": 180, "price": 1500, "cost": 1020, "packSize": 12, "minDisplay": 10, "stockCoverDays": 2.0, "abc": "A", "avgSales": 19.8, "stdDev": 3.6, "totalSalesQty": 5405, "totalSalesAmt": 8107500, "totalWasteQty": 0, "totalWasteLoss": 0, "totalStockoutQty": 0, "totalStockoutLoss": 0, "stock": 48, "safetyStock": 16, "rop": 36, "targetStock": 58, "forecastD0": 20.5, "forecastD1": 21.4, "forecastD2": 18.7, "forecastD3": 18.1, "orderToday": 24, "orderTomorrow": 24, "userOrderToday": 24, "userOrderTomorrow": 24}
        ];

        function initFullSKUDataset() {
            STATE.items = [];
            
            // 1. 실제 상품명 10개 복구
            POS_106_DATA.forEach(item => {
                let newItem = { ...item };
                newItem.studentStock = newItem.stock;
                newItem.aiStock = newItem.stock;
                newItem.userActualSales = '';
                newItem.riskStatus = 'NORMAL';
                newItem.riskReason = '자동 초기화됨';
                newItem.riskAction = '자동 설정';
                STATE.items.push(newItem);
            });
            
            // 2. 나머지 11번부터 106번까지 임시 생성
            const baseCount = STATE.items.length;
            for (let i = baseCount + 1; i <= 106; i++) {
                const skuId = `SKU${String(i).padStart(3, '0')}`;
                const catIndex = i % 8;
                const cat = CATEGORIES[catIndex];
                const price = 1200 + (i % 10) * 500;
                const cost = Math.round(price * 0.68);
                const avg = (3.5 + (i % 15) * 1.2).toFixed(1) * 1;
                const std = (avg * 0.22).toFixed(1) * 1;
                const shelf = (cat === 'FF/즉석식' || cat === '유제품') ? (2 + (i % 3)) : 180;
                const minDisp = 3 + (i % 5);
                const pack = (cat === '음료' || cat === '라면/즉석조리') ? 12 : 5;
                const abcGrade = i % 3 === 0 ? 'A' : (i % 3 === 1 ? 'B' : 'C');
                
                const ss = Math.ceil(1.65 * std) + minDisp;
                const pred0 = (avg * 1.05).toFixed(1) * 1;
                const pred1 = (avg * 1.1).toFixed(1) * 1;
                const pred2 = (avg * 0.95).toFixed(1) * 1;
                const pred3 = (avg * 0.9).toFixed(1) * 1;

                const rop = Math.ceil(pred0 + ss);
                const targetStock = Math.ceil(pred0 + pred1 + ss);
                const beginningStock = Math.round(avg * (0.8 + (i % 3) * 0.7));

                const neededToday = targetStock - beginningStock;
                const ordToday = neededToday > 0 ? Math.ceil(neededToday / pack) * pack : 0;
                const estTomorrow = Math.max(0, beginningStock + ordToday - pred0);
                const targetTomorrow = Math.ceil(pred1 + pred2 + ss);
                const neededTomorrow = targetTomorrow - estTomorrow;
                const ordTomorrow = neededTomorrow > 0 ? Math.ceil(neededTomorrow / pack) * pack : 0;

                STATE.items.push({
                    id: skuId, name: `[POS] ${cat} 인기상품 ${i}호`, category: cat, subCat: '일반', storage: shelf <= 3 ? '냉장' : '상온',
                    shelfLifeDays: shelf, price: price, cost: cost, packSize: pack, minDisplay: minDisp, stockCoverDays: 1.0,
                    abc: abcGrade, // 등급 추가됨
                    avgSales: avg, stdDev: std, totalSalesQty: Math.round(avg * 270), totalSalesAmt: Math.round(avg * 270 * price),
                    totalWasteQty: shelf <= 3 ? Math.round(i % 5) : 0, totalWasteLoss: shelf <= 3 ? Math.round((i % 5) * cost) : 0,
                    totalStockoutQty: 0, totalStockoutLoss: 0, 
                    stock: beginningStock, studentStock: beginningStock, aiStock: beginningStock, 
                    safetyStock: ss, rop: rop, targetStock: targetStock, forecastD0: pred0, forecastD1: pred1, forecastD2: pred2, forecastD3: pred3,
                    orderToday: ordToday, orderTomorrow: ordTomorrow, userOrderToday: ordToday, userOrderTomorrow: ordTomorrow,
                    userActualSales: '', 
                    riskStatus: 'NORMAL', riskReason: '자동 초기화됨', riskAction: '자동 설정'
                });
            }
        }

        function getEndingStock(item) {
            const beginning = item.stock || 0;
            const todayReceipt = item.userOrderToday || 0;
            const forecastToday = Math.round(item.forecastD0 || 0);
            return Math.max(0, Math.round(beginning + todayReceipt - forecastToday));
        }

        // ==========================================
        // 대시보드 렌더링
        // ==========================================
        function renderDashboard() {
            let stockoutCount = 0, wasteCount = 0, overstockCount = 0, normalCount = 0;
            let sumForecastD0 = 0, sumCurrentStock = 0, sumSafetyStock = 0, sumTargetStock = 0, sumEndingStock = 0;

            STATE.items.forEach(i => {
                if (i.stock <= i.rop) i.riskStatus = 'STOCKOUT';
                else if (i.shelfLifeDays <= 3 && i.stock > i.avgSales * 1.5) i.riskStatus = 'WASTE';
                else if (i.stock >= i.targetStock * 1.5) i.riskStatus = 'OVERSTOCK';
                else i.riskStatus = 'NORMAL';

                if (i.riskStatus === 'STOCKOUT') stockoutCount++;
                else if (i.riskStatus === 'WASTE') wasteCount++;
                else if (i.riskStatus === 'OVERSTOCK') overstockCount++;
                else normalCount++;

                sumForecastD0 += i.forecastD0;
                sumCurrentStock += i.stock;
                sumSafetyStock += i.safetyStock;
                sumTargetStock += i.targetStock;
                sumEndingStock += getEndingStock(i);
            });

            document.getElementById('todayStatusBannerDate').innerText = STATE.currentDateStr;
            const guideDate = document.getElementById('beginnerGuideDate');
            const guideRiskCount = document.getElementById('beginnerRiskCount');
            if (guideDate) guideDate.innerText = STATE.currentDateStr;
            if (guideRiskCount) guideRiskCount.innerText = `${stockoutCount + wasteCount}개`;
            document.getElementById('statSumForecastD0').innerText = `${Math.round(sumForecastD0).toLocaleString()}개`;
            document.getElementById('statSumCurrentStock').innerText = `${sumCurrentStock.toLocaleString()}개`;
            document.getElementById('statSumSafetyStock').innerText = `${sumSafetyStock.toLocaleString()}개`;
            document.getElementById('statSumEndingStock').innerText = `${sumEndingStock.toLocaleString()}개`;
            document.getElementById('statSumTargetStock').innerText = `${sumTargetStock.toLocaleString()}개`;
            document.getElementById('statCountStockout').innerText = `${stockoutCount}개`;
            document.getElementById('statCountWaste').innerText = `${wasteCount}개`;
            
            document.getElementById('dashboardCountStockout').innerText = `${stockoutCount}개`;
            document.getElementById('dashboardCountWaste').innerText = `${wasteCount}개`;
            document.getElementById('dashboardCountOverstock').innerText = `${overstockCount}개`;
            document.getElementById('dashboardCountNormal').innerText = `${normalCount}개`;
            document.getElementById('riskCountBadge').innerText = stockoutCount + wasteCount;
            
            renderCategoryTable();
            renderProfitTrendChart();
            renderRiskDoughnutChart(stockoutCount, wasteCount, overstockCount, normalCount);
        }

        function renderCategoryTable() {
            const tbody = document.getElementById('categoryProfitTableBody');
            tbody.innerHTML = '';
            let grandNet = 0;
            CATEGORIES.forEach(cat => {
                const catItems = STATE.items.filter(i => i.category === cat);
                let catSales = 0, catWaste = 0, catStockout = 0;
                catItems.forEach(item => { catSales += item.totalSalesAmt; catWaste += item.totalWasteLoss; catStockout += item.totalStockoutLoss; });
                const monthlyProfitIncrease = Math.round((catWaste / 9) * 0.65 + (catSales / 9) * 0.015);
                grandNet += monthlyProfitIncrease;

                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50 transition';
                tr.innerHTML = `<td class="p-3 font-semibold text-slate-800">${cat}</td><td class="p-3">${catItems.length}개 SKU</td><td class="p-3 text-right font-medium">${catSales.toLocaleString()}원</td><td class="p-3 text-right text-red-600">${catWaste.toLocaleString()}원</td><td class="p-3 text-right text-amber-600">${catStockout.toLocaleString()}원</td><td class="p-3 text-right font-bold text-blue-700 bg-blue-50/50">+${monthlyProfitIncrease.toLocaleString()}원</td>`;
                tbody.appendChild(tr);
            });
            document.getElementById('kpiNetProfit').innerText = `+${grandNet.toLocaleString()}원`;
            document.getElementById('kpiAnnualProfit').innerText = `+${(grandNet * 12).toLocaleString()}원`;
        }

        function renderProfitTrendChart() {
            const ctx = document.getElementById('chartProfitTrend').getContext('2d');
            if (STATE.charts.profitTrend) STATE.charts.profitTrend.destroy();
            STATE.charts.profitTrend = new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: ['5월', '6월', '7월', '8월', '9월', '10월(AI적용)'],
                    datasets: [
                        { label: '기존 방식 폐기/품절 손실액 (원)', data: [31200, 34500, 38900, 36200, 35100, 9500], backgroundColor: '#f97316', borderRadius: 6 },
                        { label: 'AI 최적화 후 순이익 증가분 (원)', data: [0, 0, 0, 0, 0, 387426], backgroundColor: '#2563eb', borderRadius: 6 }
                    ]
                },
                options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } }, scales: { y: { beginAtZero: true } } }
            });
        }

        function renderRiskDoughnutChart(s, w, o, n) {
            const ctx = document.getElementById('chartRiskRatio').getContext('2d');
            if (STATE.charts.riskRatio) STATE.charts.riskRatio.destroy();
            STATE.charts.riskRatio = new Chart(ctx, {
                type: 'doughnut',
                data: {
                    labels: ['품절위험', '폐기위험', '과다재고', '정상재고'],
                    datasets: [{ data: [s, w, o, n], backgroundColor: ['#ef4444', '#f97316', '#eab308', '#22c55e'], borderWidth: 2, borderColor: '#ffffff' }]
                },
                options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, cutout: '70%' }
            });
        }

        // ==========================================
        // 106개 품목 판매예측 렌더링 (실제판매 기입 추가)
        // ==========================================
        function renderForecastTable() {
            const tbody = document.getElementById('forecastTableBody');
            tbody.innerHTML = '';
            
            const searchTerm = document.getElementById('forecastSearchInput').value.toLowerCase();
            const selectedCat = document.getElementById('forecastCategoryFilter').value;
            let count = 0;

            STATE.items.forEach(item => {
                const matchesSearch = item.name.toLowerCase().includes(searchTerm) || item.id.toLowerCase().includes(searchTerm);
                const matchesCat = (selectedCat === 'ALL' || item.category === selectedCat);

                if (matchesSearch && matchesCat) {
                    count++;
                    const endingStock = getEndingStock(item);
                    const diagnosis = getRiskDiagnosis(item);
                    const tr = document.createElement('tr');
                    tr.className = 'hover:bg-slate-50 transition border-b border-slate-100';
                    const abcBadge = item.abc === 'A' ? 'bg-blue-100 text-blue-800 font-bold' : (item.abc === 'B' ? 'bg-amber-100 text-amber-800' : 'bg-slate-100 text-slate-600');
                    
                    tr.innerHTML = `
                        <td class="p-3 font-mono font-bold text-slate-500 cursor-pointer" onclick="openSkuModal('${item.id}')">${item.id}</td>
                        <td class="p-3 text-slate-600 cursor-pointer" onclick="openSkuModal('${item.id}')"><span class="bg-slate-100 px-2 py-0.5 rounded text-xs">${item.category}</span></td>
                        <td class="p-3 font-semibold text-slate-800 cursor-pointer" onclick="openSkuModal('${item.id}')">${item.name}</td>
                        <td class="p-3 text-center cursor-pointer" onclick="openSkuModal('${item.id}')"><span class="px-2 py-0.5 rounded text-xs ${abcBadge}">${item.abc}</span></td>
                        <td class="p-3 text-center font-bold bg-slate-50">${item.stock}개</td>
                        <td class="p-3 text-center text-emerald-700 font-semibold bg-emerald-50/70">${item.safetyStock}개</td>
                        <td class="p-3 text-center font-extrabold text-cyan-800 bg-cyan-50/90">${endingStock}개</td>
                        <td class="p-3 text-center font-bold text-blue-800 bg-blue-50/60">${item.forecastD0}개</td>
                        <td class="p-3 text-center bg-purple-50">
                            <input type="number" min="0" value="${item.userActualSales}" placeholder="자동" onchange="updateActualSales('${item.id}', this.value)" class="w-16 p-1 text-center font-bold border border-purple-300 rounded focus:ring-2 focus:ring-purple-500">
                        </td>
                        <td class="p-3 text-center font-bold text-emerald-700 bg-emerald-50/40">+${item.userOrderToday}개</td>
                        <td class="p-3 text-center text-amber-700">${item.rop}개</td>
                        <td class="p-3 text-center font-bold text-indigo-700 bg-indigo-50/70">${item.targetStock}개</td>
                        <td class="p-3 text-center font-bold text-indigo-800 bg-indigo-50/60">${item.forecastD1}개</td>
                        <td class="p-3 text-center text-slate-700">${item.forecastD2}개</td>
                        <td class="p-3 text-center text-slate-700">${item.forecastD3}개</td>
                    `;
                    tbody.appendChild(tr);
                }
            });
            document.getElementById('forecastDisplayedCount').innerText = count;
        }

        function filterForecastTable() { renderForecastTable(); }
        function resetForecastFilters() { document.getElementById('forecastSearchInput').value = ''; document.getElementById('forecastCategoryFilter').value = 'ALL'; renderForecastTable(); }
        
        function updateActualSales(id, val) {
            const item = STATE.items.find(i => i.id === id);
            if (item) {
                item.userActualSales = val;
            }
        }

        // ==========================================
        // 발주관리 렌더링
        // ==========================================
        function getOrderStatus(item) {
            const stock = Number(item.stock) || 0;
            const rop = Number(item.rop) || 0;
            const safety = Number(item.safetyStock) || 0;
            const target = Number(item.targetStock) || 0;
            const ending = getEndingStock(item);
            const forecastOrder = (Number(item.orderToday) || 0) + (Number(item.orderTomorrow) || 0);
            if (stock <= rop || ending < safety) {
                return { label: '긴급 발주', cls: 'bg-red-100 text-red-700 border-red-200' };
            }
            if (Number(item.shelfLifeDays) <= 3 && stock > (Number(item.avgSales) || 0) * 1.5) {
                return { label: '폐기 주의', cls: 'bg-orange-100 text-orange-700 border-orange-200' };
            }
            if (target > 0 && stock >= target * 1.5) {
                return { label: '과다 재고', cls: 'bg-amber-100 text-amber-800 border-amber-200' };
            }
            if (forecastOrder > 0) {
                return { label: '발주 권장', cls: 'bg-blue-100 text-blue-700 border-blue-200' };
            }
            return { label: '정상', cls: 'bg-emerald-100 text-emerald-700 border-emerald-200' };
        }

        function renderOrderTable() {
            const tbody = document.getElementById('orderTableBody');
            tbody.innerHTML = '';
            let totalTodayAmount = 0, totalTodayQty = 0, totalTomorrowAmount = 0, totalTomorrowQty = 0;

            STATE.items.forEach(item => {
                const todayAmount = item.userOrderToday * item.cost;
                const tomorrowAmount = item.userOrderTomorrow * item.cost;
                totalTodayAmount += todayAmount; totalTodayQty += item.userOrderToday;
                totalTomorrowAmount += tomorrowAmount; totalTomorrowQty += item.userOrderTomorrow;

                const status = getOrderStatus(item);
                // '발주 필요품목만'은 실제 재고 부족이 확인된 긴급 발주 대상만 표시한다.
                // 단순 AI 발주 권장이나 폐기 주의/과다 재고 품목은 이 필터에서 제외한다.
                const isOrderNeeded = status.label === '긴급 발주' &&
                    status.label !== '폐기 주의' &&
                    ((Number(item.stock) || 0) <= (Number(item.rop) || 0) || getEndingStock(item) < (Number(item.safetyStock) || 0));
                let show = (STATE.orderFilter === 'ALL') || (STATE.orderFilter === 'ORDER_ONLY' && isOrderNeeded) || (STATE.orderFilter === 'FF_ONLY' && item.category === 'FF/즉석식');
                if (show) {
                    const endingStock = getEndingStock(item);
                    const tr = document.createElement('tr');
                    tr.className = 'hover:bg-slate-50 transition border-b border-slate-100';
                    tr.innerHTML = `
                        <td class="p-3">
                            <div class="font-semibold text-slate-800">${item.name}</div>
                            <div class="text-[11px] text-slate-400 font-mono">${item.id} | ${item.category} | <span class="font-bold text-blue-600">${item.abc}등급</span></div>
                        </td>
                        <td class="p-3 text-center"><span class="inline-block whitespace-nowrap rounded-full border px-2 py-1 text-xs font-bold ${status.cls}">${status.label}</span><div class="mt-1 text-[10px] text-slate-500">권장: 오늘 ${Number(item.orderToday)||0}개 · 내일 ${Number(item.orderTomorrow)||0}개</div></td>
                        <td class="p-3 text-center text-xs"><div class="text-slate-700 font-semibold">${item.cost.toLocaleString()}원</div></td>
                        <td class="p-3 text-center font-bold">${item.stock}개</td>
                        <td class="p-3 text-center font-semibold text-emerald-700 bg-emerald-50/30">${item.safetyStock}개</td>
                        <td class="p-3 text-center font-semibold text-amber-700">${item.rop}개</td>
                        <td class="p-3 text-center font-bold text-indigo-700 bg-indigo-50/50">${item.targetStock}개</td>
                        <td class="p-3 text-center font-bold text-emerald-700 bg-emerald-50/40">${item.orderToday}개</td>
                        <td class="p-2 text-center bg-emerald-50/40"><div class="inline-flex items-center justify-center gap-1"><button type="button" aria-label="${item.name} 오늘 발주량 1개 줄이기" onclick="adjustUserOrder('${item.id}', 'today', -1)" class="w-7 h-8 rounded border border-emerald-300 bg-white text-emerald-800 font-bold hover:bg-emerald-100">−</button><input type="number" inputmode="numeric" min="0" step="1" value="${item.userOrderToday}" aria-label="${item.name} 오늘 발주 수량 직접 입력" onchange="updateUserOrder('${item.id}', 'today', this.value)" onkeydown="if(event.key==='Enter'){this.blur();}" class="w-16 p-1 text-center font-bold border rounded focus:ring-2 focus:ring-emerald-500" /><button type="button" aria-label="${item.name} 오늘 발주량 1개 늘리기" onclick="adjustUserOrder('${item.id}', 'today', 1)" class="w-7 h-8 rounded border border-emerald-300 bg-white text-emerald-800 font-bold hover:bg-emerald-100">+</button></div></td>
                        <td class="p-3 text-center font-bold text-cyan-800 bg-cyan-50/80">${endingStock}개</td>
                        <td class="p-3 text-center font-bold text-blue-700 bg-blue-50/40">${item.orderTomorrow}개</td>
                        <td class="p-2 text-center bg-blue-50/40"><div class="inline-flex items-center justify-center gap-1"><button type="button" aria-label="${item.name} 내일 발주량 1개 줄이기" onclick="adjustUserOrder('${item.id}', 'tomorrow', -1)" class="w-7 h-8 rounded border border-blue-300 bg-white text-blue-800 font-bold hover:bg-blue-100">−</button><input type="number" inputmode="numeric" min="0" step="1" value="${item.userOrderTomorrow}" aria-label="${item.name} 내일 발주 수량 직접 입력" onchange="updateUserOrder('${item.id}', 'tomorrow', this.value)" onkeydown="if(event.key==='Enter'){this.blur();}" class="w-16 p-1 text-center font-bold border rounded focus:ring-2 focus:ring-blue-500" /><button type="button" aria-label="${item.name} 내일 발주량 1개 늘리기" onclick="adjustUserOrder('${item.id}', 'tomorrow', 1)" class="w-7 h-8 rounded border border-blue-300 bg-white text-blue-800 font-bold hover:bg-blue-100">+</button></div></td>
                        <td class="p-3 text-right font-bold text-slate-800">${todayAmount.toLocaleString()}원</td>
                    `;
                    tbody.appendChild(tr);
                }
            });
            document.getElementById('summaryTodayOrderAmount').innerText = `${totalTodayAmount.toLocaleString()}원`;
            document.getElementById('summaryTodayOrderQty').innerText = `(${totalTodayQty}개)`;
            document.getElementById('summaryTomorrowOrderAmount').innerText = `${totalTomorrowAmount.toLocaleString()}원`;
            document.getElementById('summaryTomorrowOrderQty').innerText = `(${totalTomorrowQty}개)`;
        }

        function filterOrderTable(type) { 
            STATE.orderFilter = type; 
            ['ALL', 'ORDER_ONLY', 'FF_ONLY'].forEach(f => {
                const btn = document.getElementById(`orderFilterBtn-${f}`);
                if (btn) btn.className = (f === type) ? 'bg-white shadow-sm text-slate-800 font-semibold px-2.5 py-1 rounded' : 'text-slate-600 font-semibold px-2.5 py-1 rounded hover:bg-white';
            });
            renderOrderTable(); 
        }

        function updateUserOrder(id, type, value) {
            const item = STATE.items.find(i => i.id === id);
            if (!item) return;
            const parsed = Math.max(0, Number.parseInt(value, 10) || 0);
            if(type === 'today') item.userOrderToday = parsed;
            else item.userOrderTomorrow = parsed;
            renderOrderTable(); renderForecastTable(); renderDashboard();
        }

        // 발주 수량은 키보드로 직접 입력하거나 +/- 버튼으로 조절할 수 있습니다.
        function adjustUserOrder(id, type, delta) {
            const item = STATE.items.find(i => i.id === id);
            if (!item) return;
            if (type === 'today') item.userOrderToday = Math.max(0, (Number(item.userOrderToday) || 0) + delta);
            else item.userOrderTomorrow = Math.max(0, (Number(item.userOrderTomorrow) || 0) + delta);
            renderOrderTable(); renderForecastTable(); renderDashboard();
        }

        function applyAllAiOrders() {
            STATE.items.forEach(item => { item.userOrderToday = item.orderToday; item.userOrderTomorrow = item.orderTomorrow; });
            renderOrderTable(); renderForecastTable(); renderDashboard(); showToast('AI 추천 수량 전체 적용 완료!');
        }


        // 발주서 다운로드: 현재 사용자가 입력한 발주수량을 기준으로 오늘/내일 발주서를 각각 생성합니다.
        function getOrderSlipData(type) {
            const isToday = type === 'today';
            const date = isToday ? STATE.currentDateStr : getNextDateString(STATE.currentDateStr);
            const rows = STATE.items
                .map(item => {
                    const qty = Math.max(0, Number.parseInt(isToday ? item.userOrderToday : item.userOrderTomorrow, 10) || 0);
                    return { code: item.id, category: item.category, name: item.name, qty, cost: Number(item.cost) || 0, amount: qty * (Number(item.cost) || 0) };
                })
                .filter(row => row.qty > 0);
            return { date, type: isToday ? '오늘' : '내일', rows, totalQty: rows.reduce((sum, row) => sum + row.qty, 0), totalAmount: rows.reduce((sum, row) => sum + row.amount, 0) };
        }

        function getNextDateString(dateStr) {
            const parts = String(dateStr || '2026-10-01').split('-').map(Number);
            const date = new Date(parts[0], (parts[1] || 1) - 1, parts[2] || 1);
            date.setDate(date.getDate() + 1);
            return `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`;
        }

        function downloadOrderSlip(type, format) {
            const slip = getOrderSlipData(type);
            if (!slip.rows.length) {
                alert(`${slip.type} 발주 수량이 0개입니다. 발주 수량을 입력한 뒤 다시 다운로드해 주세요.`);
                return;
            }
            const filename = `편의점_${slip.date}_${slip.type}발주서`;
            const tableRows = slip.rows.map((row, idx) => [idx + 1, row.code, row.category, row.name, row.qty, row.cost, row.amount]);
            const headers = ['번호', '상품코드', '카테고리', '상품명', '발주수량(개)', '단가(원)', '발주금액(원)'];
            if (format === 'excel') {
                const aoa = [[`편의점 ${slip.type} 발주서`], ['발주일자', slip.date], ['발주 품목 수', slip.rows.length], ['총 발주 수량(개)', slip.totalQty], ['총 발주 금액(원)', slip.totalAmount], [], headers, ...tableRows, [], ['합계', '', '', '', slip.totalQty, '', slip.totalAmount]];
                if (window.XLSX && XLSX.utils && XLSX.writeFile) {
                    const ws = XLSX.utils.aoa_to_sheet(aoa);
                    ws['!cols'] = [{wch:8},{wch:14},{wch:18},{wch:32},{wch:15},{wch:14},{wch:16}];
                    const wb = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(wb, ws, '발주서');
                    XLSX.writeFile(wb, `${filename}.xlsx`);
                    return;
                }
                // 라이브러리 로딩 실패 시 Excel에서 바로 열 수 있는 UTF-8 CSV를 실제 다운로드합니다.
                const csv = '\ufeff' + aoa.map(row => row.map(value => `"${String(value ?? '').replace(/"/g, '""')}"`).join(',')).join('\r\n');
                const blobUrl = URL.createObjectURL(new Blob([csv], {type:'text/csv;charset=utf-8;'}));
                const a = document.createElement('a'); a.href = blobUrl; a.download = `${filename}.csv`; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(blobUrl);
                showToast('Excel 라이브러리를 불러오지 못해 Excel에서 열 수 있는 CSV 발주서를 저장했어.');
                return;
            }
            const safe = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
            const printRows = slip.rows.map((row, idx) => `<tr><td>${idx+1}</td><td>${safe(row.code)}</td><td>${safe(row.category)}</td><td>${safe(row.name)}</td><td class="num">${row.qty.toLocaleString()}</td><td class="num">${row.cost.toLocaleString()}</td><td class="num">${row.amount.toLocaleString()}</td><td></td></tr>`).join('');
            const win = window.open('', '_blank');
            if (!win) { alert('PDF 발주서를 열 수 없습니다. 브라우저에서 팝업을 허용한 뒤 다시 눌러주세요.'); return; }
            win.document.open();
            win.document.write(`<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>${safe(filename)}</title><style>
                @page{size:A4 landscape;margin:12mm}body{font-family:Arial,'Malgun Gothic',sans-serif;color:#111827;margin:0}h1{font-size:22px;margin:0 0 8px}p{margin:4px 0;font-size:12px}.summary{display:flex;gap:20px;margin:14px 0;padding:12px;background:#f3f4f6;font-weight:bold;font-size:13px}table{border-collapse:collapse;width:100%;font-size:10px}th,td{border:1px solid #9ca3af;padding:6px 5px}th{background:#e5e7eb;text-align:center}.num{text-align:right}.sign{margin-top:20px;display:flex;justify-content:space-between;font-size:12px}@media print{button{display:none}}
                </style></head><body><h1>편의점 ${slip.type} 발주서</h1><p>발주일자: ${safe(slip.date)}</p><p>작성일: ${new Date().toLocaleDateString('ko-KR')}</p><div class="summary"><span>발주 품목: ${slip.rows.length}종</span><span>총 발주 수량: ${slip.totalQty.toLocaleString()}개</span><span>총 발주 금액: ${slip.totalAmount.toLocaleString()}원</span></div><table><thead><tr>${headers.map(h=>`<th>${safe(h)}</th>`).join('')}<th>비고</th></tr></thead><tbody>${printRows}<tr><th colspan="4">합계</th><th>${slip.totalQty.toLocaleString()}</th><th></th><th>${slip.totalAmount.toLocaleString()}</th><th></th></tr></tbody></table><div class="sign"><span>발주 담당자: __________________</span><span>검수/확인: __________________</span></div><p style="margin-top:16px">※ 본 발주서는 현재 화면에 입력된 발주수량을 기준으로 생성되었습니다.</p><script>window.onload=function(){window.focus();window.print();};<\/script></body></html>`);
            win.document.close();
        }

        // ==========================================
        // 재고위험 분석: 초보 점주도 이해할 수 있는 원인·진단 생성
        // 진단은 기존 위험 판정 기준을 그대로 사용하며 재고 계산 로직은 변경하지 않습니다.
        // ==========================================
        function getRiskDiagnosis(item) {
            const endingStock = getEndingStock(item);
            const stock = Number(item.stock) || 0;
            const rop = Number(item.rop) || 0;
            const safetyStock = Number(item.safetyStock) || 0;
            const targetStock = Number(item.targetStock) || 0;
            const avgSales = Number(item.avgSales) || 0;
            const shelfLifeDays = Number(item.shelfLifeDays);
            const forecast = Math.max(0, Number(item.forecastD0) || 0);
            let reason = '';
            let diagnosis = '';
            let action = '';

            if (item.riskStatus === 'STOCKOUT') {
                reason = `현재고 ${stock}개가 재발주점 ${rop}개 이하입니다.`;
                if (endingStock <= 0) diagnosis = `오늘 예상 판매 ${Math.round(forecast)}개를 반영하면 기말재고가 ${endingStock}개로, 품절 가능성이 매우 높습니다.`;
                else if (endingStock < safetyStock) diagnosis = `오늘 판매 후 예상 재고 ${endingStock}개가 안전재고 ${safetyStock}개보다 적습니다.`;
                else diagnosis = `현재 재고가 재발주 기준에 도달해 추가 수요가 생기면 품절될 수 있습니다.`;
                const suggested = Math.max(0, Number(item.orderToday) || Math.max(0, targetStock - endingStock));
                action = suggested > 0 ? `오늘 ${Math.ceil(suggested)}개 발주 검토` : '입고 일정과 재고 실사 확인';
            } else if (item.riskStatus === 'WASTE') {
                reason = `남은 유통기한 ${Number.isFinite(shelfLifeDays) ? shelfLifeDays : '미확인'}일, 현재고 ${stock}개입니다.`;
                diagnosis = `재고가 평균 판매량(${avgSales.toFixed(1)}개/일)의 1.5배보다 많아 유통기한 전에 판매하지 못할 수 있습니다.`;
                action = '추가 발주 줄이기 · 선입선출 확인';
            } else if (item.riskStatus === 'OVERSTOCK') {
                reason = `현재고 ${stock}개가 목표재고 ${targetStock}개의 1.5배 이상입니다.`;
                diagnosis = `판매 속도보다 재고가 많은 상태입니다. 보관 공간과 장기 보유에 따른 폐기 가능성을 확인하세요.`;
                action = '발주 축소 · 판매 추이 확인';
            } else {
                reason = `현재고 ${stock}개, 안전재고 ${safetyStock}개, 재발주점 ${rop}개입니다.`;
                diagnosis = `현재 기준상 긴급 위험은 감지되지 않았습니다. 오늘 예상 기말재고는 ${endingStock}개입니다.`;
                action = stock <= safetyStock ? '판매 추이 관찰' : '정상 재고 유지';
            }
            return { reason, diagnosis, action };
        }

        // ==========================================
        // 재고위험 분석 렌더링
        // ==========================================
        function renderRiskTable() {
            const tbody = document.getElementById('riskTableBody');
            tbody.innerHTML = '';
            let stockoutCount = 0, wasteCount = 0, overstockCount = 0, normalCount = 0;

            STATE.items.forEach(item => {
                if (item.riskStatus === 'STOCKOUT') stockoutCount++;
                else if (item.riskStatus === 'WASTE') wasteCount++;
                else if (item.riskStatus === 'OVERSTOCK') overstockCount++;
                else normalCount++;

                const matchesFilter = (STATE.riskFilter === 'ALL' || item.riskStatus === STATE.riskFilter);

                if (matchesFilter) {
                    const endingStock = getEndingStock(item);
                    // 원인·진단 문구를 각 품목별로 생성합니다.
                    // 선언 누락 시 ReferenceError가 발생해 재고위험 분석뿐 아니라
                    // 날짜 이월 과정의 후속 발주관리 렌더링도 중단되던 문제를 수정했습니다.
                    const diagnosis = getRiskDiagnosis(item);
                    const tr = document.createElement('tr');
                    tr.className = 'hover:bg-slate-50 transition border-b border-slate-100';

                    let badgeMarkup = '';
                    if (item.riskStatus === 'STOCKOUT') badgeMarkup = `<span class="badge-risk-stockout text-xs px-2.5 py-1 rounded-full font-bold flex items-center w-fit"><i class="fa-solid fa-triangle-exclamation mr-1"></i>품절위험</span>`;
                    else if (item.riskStatus === 'WASTE') badgeMarkup = `<span class="badge-risk-waste text-xs px-2.5 py-1 rounded-full font-bold flex items-center w-fit"><i class="fa-solid fa-clock-rotate-left mr-1"></i>폐기위험</span>`;
                    else if (item.riskStatus === 'OVERSTOCK') badgeMarkup = `<span class="badge-risk-overstock text-xs px-2.5 py-1 rounded-full font-bold flex items-center w-fit"><i class="fa-solid fa-boxes-stacked mr-1"></i>과잉/과다재고</span>`;
                    else badgeMarkup = `<span class="badge-risk-normal text-xs px-2.5 py-1 rounded-full font-bold flex items-center w-fit"><i class="fa-solid fa-circle-check mr-1"></i>적정</span>`;

                    tr.innerHTML = `
                        <td class="p-3">${badgeMarkup}</td>
                        <td class="p-3">
                            <div class="font-bold text-slate-800">${item.name}</div>
                            <div class="text-xs text-slate-400 font-mono">${item.id}</div>
                        </td>
                        <td class="p-3 text-center font-bold text-slate-800">${item.stock}개</td>
                        <td class="p-3 text-center font-semibold text-emerald-700 bg-emerald-50/30">${item.safetyStock}개</td>
                        <td class="p-3 text-center font-bold text-cyan-800 bg-cyan-50/80">${endingStock}개</td>
                        <td class="p-3 text-center font-semibold text-indigo-700">${item.targetStock}개</td>
                        <td class="p-3 text-center font-semibold text-amber-700">${item.rop}개</td>
                        <td class="p-3 text-xs text-slate-700 min-w-[260px]">
                            <div class="font-semibold text-slate-800">원인: ${diagnosis.reason}</div>
                            <div class="mt-1 leading-relaxed text-slate-600">진단: ${diagnosis.diagnosis}</div>
                        </td>
                        <td class="p-3 text-center font-semibold text-xs text-indigo-700 min-w-[160px]">${diagnosis.action}</td>
                        <td class="p-3 text-center"><button class="bg-slate-200 text-slate-500 text-xs px-2 py-1 rounded">자동 처리</button></td>
                    `;
                    tbody.appendChild(tr);
                }
            });

            document.getElementById('riskBadgeCountAll').innerText = STATE.items.length;
            document.getElementById('riskBadgeCountStockout').innerText = stockoutCount;
            document.getElementById('riskBadgeCountWaste').innerText = wasteCount;
            document.getElementById('riskBadgeCountOverstock').innerText = overstockCount;
            document.getElementById('riskBadgeCountNormal').innerText = normalCount;
            const riskCountBadge = document.getElementById('riskCountBadge');
            if (riskCountBadge) riskCountBadge.innerText = stockoutCount + wasteCount + overstockCount;
            const orderSubRiskCount = document.getElementById('orderSubRiskCount');
            if (orderSubRiskCount) orderSubRiskCount.innerText = stockoutCount + wasteCount + overstockCount;
            document.getElementById('riskBadgeCountWaste').innerText = wasteCount;
            document.getElementById('riskBadgeCountOverstock').innerText = overstockCount;
            document.getElementById('riskBadgeCountNormal').innerText = normalCount;
        }

        function filterRiskCategory(cat) { STATE.riskFilter = cat; renderRiskTable(); }

        // ==========================================
        // 일일 마감 시뮬레이션 및 106개 데이터 저장
        // ==========================================
        function executeDailySimulation(skipModal = false) {
            const weather = document.getElementById('simWeatherSelect').value;
            let weatherMult = 1.0;
            if (weather === 'RAIN') weatherMult = 0.85;
            if (weather === 'HOT') weatherMult = 1.25;
            if (weather === 'COLD') weatherMult = 1.1;

            let studentDayRev = 0, studentDayWaste = 0, studentDayStockout = 0;
            let aiDayRev = 0, aiDayWaste = 0, aiDayStockout = 0;
            let itemDetails = []; // 106개 품목별 상세 성과 저장용

            STATE.items.forEach(item => {
                item.studentStock += item.userOrderToday;
                item.aiStock += item.orderToday;
                if (item.userOrderToday > 0) STATE.simTotalOrdersCount++;

                // 수요 산출: 수동 입력 값이 있으면 우선 적용!
                let actualDemand;
                if (item.userActualSales !== null && item.userActualSales !== "") {
                    actualDemand = parseInt(item.userActualSales);
                } else {
                    actualDemand = Math.round(item.forecastD0 * weatherMult * (0.7 + Math.random() * 0.6));
                }

                // 학생 정산
                let studentSales = Math.min(item.studentStock, actualDemand);
                item.studentStock -= studentSales;
                let sRev = studentSales * item.price;
                studentDayRev += sRev;
                let sMissed = Math.max(0, actualDemand - studentSales);
                let sStockoutLoss = sMissed * (item.price - item.cost);
                studentDayStockout += sStockoutLoss;

                // AI 정산
                let aiSales = Math.min(item.aiStock, actualDemand);
                item.aiStock -= aiSales;
                let aRev = aiSales * item.price;
                aiDayRev += aRev;
                let aMissed = Math.max(0, actualDemand - aiSales);
                let aStockoutLoss = aMissed * (item.price - item.cost);
                aiDayStockout += aStockoutLoss;

                // 폐기
                let sWasteLoss = 0, aWasteLoss = 0;
                if (item.shelfLifeDays <= 3) {
                    if (item.studentStock > item.avgSales * 0.8 && Math.random() < 0.3) {
                        let w = Math.min(item.studentStock, Math.ceil(item.studentStock * 0.4));
                        item.studentStock -= w; sWasteLoss = w * item.cost;
                    }
                    if (item.aiStock > item.avgSales * 0.8 && Math.random() < 0.3) {
                        let w = Math.min(item.aiStock, Math.ceil(item.aiStock * 0.4));
                        item.aiStock -= w; aWasteLoss = w * item.cost;
                    }
                }
                studentDayWaste += sWasteLoss;
                aiDayWaste += aWasteLoss;

                // 제품별 상세 정보 저장
                const sNet = sRev - sWasteLoss - sStockoutLoss;
                const aNet = aRev - aWasteLoss - aStockoutLoss;
                
                itemDetails.push({
                    id: item.id, name: item.name, abc: item.abc,
                    sSales: studentSales, aSales: aiSales,
                    sWasteLoss: sWasteLoss, aWasteLoss: aWasteLoss,
                    sStockoutLoss: sStockoutLoss, aStockoutLoss: aStockoutLoss,
                    sNet: sNet, aNet: aNet
                });

                // 학생 결과를 다음날 실제 UI 기초재고로 반영 및 수동입력칸 초기화
                item.stock = item.studentStock; 
                item.userActualSales = ''; 
            });

            STATE.simTotalRevenue += studentDayRev;
            STATE.simTotalWaste += studentDayWaste;

            const sNet = studentDayRev - studentDayWaste - studentDayStockout;
            const aNet = aiDayRev - aiDayWaste - aiDayStockout;

            STATE.closingHistory.push({
                date: STATE.currentDateStr,
                sRev: studentDayRev, sWaste: studentDayWaste, sStockout: studentDayStockout, sNet: sNet,
                aRev: aiDayRev, aWaste: aiDayWaste, aStockout: aiDayStockout, aNet: aNet,
                details: itemDetails
            });

            renderClosingHistoryTable();
            renderForecastTable(); 

            if (!skipModal) {
                showClosingReport(STATE.currentDateStr, studentDayRev, studentDayWaste, studentDayStockout, aiDayRev, aiDayWaste, aiDayStockout);
            } else {
                moveToNextDay(true);
            }
        }

        // 전체 총합 + 일별 실적 + 날짜별 품목 상세 보고서 데이터 생성
        function getCumulativeReportData() {
            const history = STATE.closingHistory || [];
            const sum = key => history.reduce((total, record) => total + (Number(record[key]) || 0), 0);
            const studentNet = sum('sNet');
            const aiNet = sum('aNet');
            const summary = [
                ['항목', '전체 총합'],
                ['마감 일수', history.length],
                ['학생 발주 누적 매출(원)', sum('sRev')],
                ['AI 발주 누적 매출(원)', sum('aRev')],
                ['학생 발주 누적 폐기손실(원)', sum('sWaste')],
                ['AI 발주 누적 폐기손실(원)', sum('aWaste')],
                ['학생 발주 누적 품절손실(원)', sum('sStockout')],
                ['AI 발주 누적 품절손실(원)', sum('aStockout')],
                ['학생 발주 누적 순이익(원)', studentNet],
                ['AI 발주 누적 순이익(원)', aiNet],
                ['AI 순이익 우위(원)', aiNet - studentNet]
            ];
            const daily = [['영업일', '학생 매출(원)', 'AI 매출(원)', '학생 폐기손실(원)', 'AI 폐기손실(원)', '학생 품절손실(원)', 'AI 품절손실(원)', '학생 순이익(원)', 'AI 순이익(원)', 'AI 순이익 차이(원)']];
            history.forEach(r => daily.push([r.date, Number(r.sRev)||0, Number(r.aRev)||0, Number(r.sWaste)||0, Number(r.aWaste)||0, Number(r.sStockout)||0, Number(r.aStockout)||0, Number(r.sNet)||0, Number(r.aNet)||0, (Number(r.aNet)||0)-(Number(r.sNet)||0)]));

            const productHeaders = ['품목코드', '품목명', '학생 판매수량', 'AI 판매수량', '학생 폐기손실(원)', 'AI 폐기손실(원)', '학생 품절손실(원)', 'AI 품절손실(원)', '학생 순이익(원)', 'AI 순이익(원)', 'AI 순이익 차이(원)'];
            const productsMap = new Map();
            history.forEach(r => (r.details || []).forEach(d => {
                const key = String(d.id ?? d.name);
                if (!productsMap.has(key)) productsMap.set(key, {id:d.id, name:d.name, studentSales:0, aiSales:0, studentWaste:0, aiWaste:0, studentStockout:0, aiStockout:0, studentNet:0, aiNet:0});
                const a=productsMap.get(key);
                a.studentSales += Number(d.sSales)||0; a.aiSales += Number(d.aSales)||0;
                a.studentWaste += Number(d.sWasteLoss)||0; a.aiWaste += Number(d.aWasteLoss)||0;
                a.studentStockout += Number(d.sStockoutLoss)||0; a.aiStockout += Number(d.aStockoutLoss)||0;
                a.studentNet += Number(d.sNet)||0; a.aiNet += Number(d.aNet)||0;
            }));
            const products = [productHeaders];
            [...productsMap.values()].forEach(a => products.push([a.id, a.name, a.studentSales, a.aiSales, a.studentWaste, a.aiWaste, a.studentStockout, a.aiStockout, a.studentNet, a.aiNet, a.aiNet-a.studentNet]));

            const dailyDetails = history.map(r => {
                const rows = [['영업일', r.date], [], ['품목코드', '품목명', '학생 판매수량', 'AI 판매수량', '학생 폐기손실(원)', 'AI 폐기손실(원)', '학생 품절손실(원)', 'AI 품절손실(원)', '학생 순이익(원)', 'AI 순이익(원)', 'AI 순이익 차이(원)']];
                (r.details || []).forEach(d => rows.push([d.id, d.name, Number(d.sSales)||0, Number(d.aSales)||0, Number(d.sWasteLoss)||0, Number(d.aWasteLoss)||0, Number(d.sStockoutLoss)||0, Number(d.aStockoutLoss)||0, Number(d.sNet)||0, Number(d.aNet)||0, (Number(d.aNet)||0)-(Number(d.sNet)||0)]));
                rows.push([], ['하루 합계', '', (r.details||[]).reduce((n,d)=>n+(Number(d.sSales)||0),0), (r.details||[]).reduce((n,d)=>n+(Number(d.aSales)||0),0), Number(r.sWaste)||0, Number(r.aWaste)||0, Number(r.sStockout)||0, Number(r.aStockout)||0, Number(r.sNet)||0, Number(r.aNet)||0, (Number(r.aNet)||0)-(Number(r.sNet)||0)]);
                return {date:r.date, rows};
            });
            return {history, summary, daily, products, dailyDetails};
        }

        function exportCumulativeExcel() {
            const data = getCumulativeReportData();
            if (!data.history.length) { alert('저장할 보고서가 없습니다. 먼저 POS 시뮬레이터에서 영업 마감을 진행해 주세요.'); return; }
            if (typeof XLSX === 'undefined') {
                const csvRows = [
                    '=== 전체 총합 ===', ...data.summary.map(row => row.map(v => String(v ?? '').replaceAll('"','""')).map(v => `"${v}"`).join(',')),
                    '', '=== 일별 실적 ===', ...data.daily.map(row => row.map(v => String(v ?? '').replaceAll('"','""')).map(v => `"${v}"`).join(',')),
                    '', '※ 날짜별 품목 상세는 XLSX 저장 기능에서 날짜별 시트로 제공됩니다.'
                ].join('\r\n');
                const blob = new Blob(['\ufeff' + csvRows], {type:'text/csv;charset=utf-8;'});
                const url = URL.createObjectURL(blob); const a=document.createElement('a'); a.href=url; a.download='편의점_총합_일별영업보고서.csv'; a.click(); URL.revokeObjectURL(url);
                alert('Excel 파일 생성 기능을 불러오지 못해 총합·일별 요약 CSV로 저장했습니다. 날짜별 품목 상세는 인터넷 연결 후 Excel 저장을 이용해 주세요.'); return;
            }
            const wb = XLSX.utils.book_new();
            const addSheet = (name, rows) => {
                const ws = XLSX.utils.aoa_to_sheet(rows);
                const headerRow = rows.find(row => row && row.length) || [];
                ws['!cols'] = headerRow.map(h => ({wch: Math.min(30, Math.max(14, String(h ?? '').length + 3))}));
                XLSX.utils.book_append_sheet(wb, ws, name.slice(0,31));
            };
            addSheet('전체 총합', data.summary);
            addSheet('일별 실적', data.daily);
            addSheet('품목별 누적', data.products);
            data.dailyDetails.forEach((day, index) => {
                const safeDate = String(day.date || `일차${index+1}`).replace(/[\\/?*\[\]:]/g, '-').slice(0,20);
                addSheet(`상세_${safeDate || index+1}`, day.rows);
            });
            XLSX.writeFile(wb, '편의점_총합_일별_상세영업보고서.xlsx');
        }

        function exportCumulativePDF() {
            const data = getCumulativeReportData();
            if (!data.history.length) { alert('저장할 보고서가 없습니다. 먼저 POS 시뮬레이터에서 영업 마감을 진행해 주세요.'); return; }
            const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
            const money = n => Math.round(Number(n)||0).toLocaleString('ko-KR') + '원';
            const table = (headers, rows) => `<table><thead><tr>${headers.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.map(row=>`<tr>${row.map(v=>`<td>${esc(v)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
            const summaryRows = data.summary.slice(1).map(r => `<tr><th>${esc(r[0])}</th><td>${typeof r[1] === 'number' ? (r[0].includes('일수') ? r[1]+'일' : money(r[1])) : esc(r[1])}</td></tr>`).join('');
            const dailyRows = data.history.map(r => [r.date, money(r.sRev), money(r.aRev), money(r.sWaste), money(r.aWaste), money(r.sStockout), money(r.aStockout), money(r.sNet), money(r.aNet), money((Number(r.aNet)||0)-(Number(r.sNet)||0))]);
            const dailyTable = table(data.daily[0], dailyRows);
            const productRows = data.products.slice(1).map(r => r.map((v,i)=> i >= 4 ? money(v) : v));
            const productTable = table(data.products[0], productRows);
            const daySections = data.dailyDetails.map(day => {
                const header = day.rows[2];
                const rows = day.rows.slice(3).filter(row => row.length && row[0] !== '하루 합계').map(row => row.map((v,i)=>i>=4?money(v):v));
                const record = data.history.find(r => r.date === day.date) || {};
                const totals = [[`하루 합계 (${day.date})`, '', '', '', money(record.sWaste), money(record.aWaste), money(record.sStockout), money(record.aStockout), money(record.sNet), money(record.aNet), money((Number(record.aNet)||0)-(Number(record.sNet)||0))]];
                return `<section class="day"><h2>날짜별 상세: ${esc(day.date)}</h2><p>학생 매출 ${money(record.sRev)} · AI 매출 ${money(record.aRev)} · 학생 순이익 ${money(record.sNet)} · AI 순이익 ${money(record.aNet)}</p>${table(header, rows)}${table(header, totals)}</section>`;
            }).join('');
            const popup = window.open('', '_blank');
            if (!popup) { alert('PDF 인쇄 창이 차단되었습니다. 브라우저에서 팝업을 허용한 뒤 다시 눌러주세요.'); return; }
            popup.document.write(`<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><title>편의점 총합·일별 상세 영업 보고서</title><style>
                @page{size:A4 landscape;margin:12mm}body{font-family:Arial,'Malgun Gothic',sans-serif;color:#1e293b;font-size:9px}h1{font-size:22px;margin:0 0 6px}h2{font-size:15px;margin:20px 0 8px}p{color:#475569}table{border-collapse:collapse;width:100%;margin:8px 0 16px}th,td{border:1px solid #cbd5e1;padding:4px;text-align:right;white-space:nowrap}th{background:#f1f5f9}td:first-child,th:first-child{text-align:left}.summary{width:65%}.day{page-break-before:always;break-before:page}.day:first-of-type{page-break-before:auto;break-before:auto}.note{margin-top:10px;color:#64748b}@media print{.no-print{display:none}}
                </style></head><body><h1>편의점 총합·일별 상세 영업 성과 보고서</h1><p>마감 기록 ${data.history.length}일 · 생성일 ${new Date().toLocaleDateString('ko-KR')}</p><h2>1. 전체 총합</h2><table class="summary">${summaryRows}</table><h2>2. 날짜별 영업 실적 요약</h2>${dailyTable}<h2>3. 품목별 누적 실적</h2>${productTable}<h2>4. 날짜별 품목 상세 내역</h2>${daySections}<p class="note">※ 현재 브라우저에 저장된 POS 시뮬레이션 마감 기록을 기준으로 작성되었습니다.</p><script>window.onload=()=>{window.focus();window.print();};window.onafterprint=()=>window.close();<\/script></body></html>`);
            popup.document.close();
        }

        // ==========================================
        // 일자별 마감 내역 표 그리기 (상세보기 포함)
        // ==========================================
        function renderCumulativeReport() {
            const history = STATE.closingHistory || [];
            const sum = key => history.reduce((total, record) => total + (Number(record[key]) || 0), 0);
            const studentRevenue = sum('sRev');
            const aiRevenue = sum('aRev');
            const studentWaste = sum('sWaste');
            const studentStockout = sum('sStockout');
            const studentNet = sum('sNet');
            const aiNet = sum('aNet');
            const difference = aiNet - studentNet;
            const money = value => `${Math.round(value).toLocaleString('ko-KR')}원`;

            const setText = (id, value) => {
                const element = document.getElementById(id);
                if (element) element.textContent = value;
            };
            setText('cumulativeReportPeriod', `기록 ${history.length}일`);
            setText('cumStudentRevenue', money(studentRevenue));
            setText('cumAiRevenue', money(aiRevenue));
            setText('cumStudentWaste', money(studentWaste));
            setText('cumStudentStockout', money(studentStockout));
            setText('cumNetDifference', `${difference > 0 ? '+' : difference < 0 ? '−' : ''}${money(Math.abs(difference))}`);
            setText('cumulativeReportConclusion', history.length === 0
                ? '아직 마감 기록이 없습니다. ‘다음 날로 진행’을 누르면 누적 보고서가 생성됩니다.'
                : difference > 0
                    ? `누적 순이익 기준으로 AI 발주가 학생 발주보다 ${money(difference)} 더 높습니다. 학생 누적 순이익은 ${money(studentNet)}, AI 누적 순이익은 ${money(aiNet)}입니다.`
                    : difference < 0
                        ? `누적 순이익 기준으로 학생 발주가 AI 발주보다 ${money(Math.abs(difference))} 더 높습니다. 학생 누적 순이익은 ${money(studentNet)}, AI 누적 순이익은 ${money(aiNet)}입니다.`
                        : `누적 순이익이 같습니다. 학생과 AI 누적 순이익은 각각 ${money(studentNet)}입니다.`);
        }

        function renderClosingHistoryTable() {
            const tbody = document.getElementById('closingHistoryBody');
            if (!tbody) return;
            tbody.innerHTML = '';
            renderCumulativeReport();
            
            if(STATE.closingHistory.length === 0) {
                tbody.innerHTML = `<tr><td colspan="5" class="p-8 text-center text-slate-400">아직 영업 마감 데이터가 없습니다. '다음 날로 진행' 버튼을 눌러주세요.</td></tr>`;
                return;
            }

            [...STATE.closingHistory].reverse().forEach((record, reversedIndex) => {
                const actualIndex = STATE.closingHistory.length - 1 - reversedIndex; 
                const diff = record.aNet - record.sNet;
                
                let diffBadge = '';
                if (diff > 0) diffBadge = `<span class="bg-indigo-100 text-indigo-700 px-2 py-0.5 rounded font-bold">AI +${diff.toLocaleString()}원</span>`;
                else if (diff < 0) diffBadge = `<span class="bg-emerald-100 text-emerald-700 px-2 py-0.5 rounded font-bold">학생 +${Math.abs(diff).toLocaleString()}원</span>`;
                else diffBadge = `<span class="bg-slate-100 text-slate-600 px-2 py-0.5 rounded font-bold">무승부</span>`;

                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50 transition border-b border-slate-100';
                tr.innerHTML = `
                    <td class="p-2 font-mono font-bold text-slate-700">${record.date.substring(5)}</td>
                    <td class="p-2 text-right border-l text-indigo-700 font-bold bg-indigo-50/20">${record.sNet.toLocaleString()}원</td>
                    <td class="p-2 text-right border-l text-indigo-700 font-bold bg-indigo-50/20">${record.aNet.toLocaleString()}원</td>
                    <td class="p-2 text-center bg-slate-50 border-l">${diffBadge}</td>
                    <td class="p-2 text-center border-l">
                        <button onclick="openDailyDetailModal(${actualIndex})" class="bg-slate-800 text-white text-xs px-3 py-1.5 rounded hover:bg-slate-700 transition font-semibold">상세보기</button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        // ==========================================
        // 제품별 마감 상세 팝업 및 날짜별 개별 내보내기
        // ==========================================
        let selectedDailyDetailIndex = null;
        function openDailyDetailModal(index) {
            selectedDailyDetailIndex = index;
            const record = STATE.closingHistory[index];
            if(!record || !record.details) return;

            document.getElementById('detailModalTitle').innerText = `${record.date} 제품별 마감 상세 리포트`;
            
            const tbody = document.getElementById('dailyDetailBody');
            tbody.innerHTML = '';

            record.details.forEach(item => {
                const tr = document.createElement('tr');
                tr.className = 'hover:bg-slate-50 transition border-b border-slate-100';
                const abcBadge = item.abc === 'A' ? 'text-blue-700' : (item.abc === 'B' ? 'text-amber-600' : 'text-slate-600');
                
                tr.innerHTML = `
                    <td class="p-2">
                        <div class="font-bold text-slate-800 truncate w-32" title="${item.name}">${item.name}</div>
                        <div class="text-[10px] text-slate-400 font-mono">${item.id}</div>
                    </td>
                    <td class="p-2 text-center font-bold ${abcBadge}">${item.abc}</td>
                    <td class="p-2 text-right font-medium bg-blue-50/20 border-l">${item.sSales}개</td>
                    <td class="p-2 text-right font-medium bg-blue-50/20">${item.aSales}개</td>
                    <td class="p-2 text-right text-orange-600 bg-orange-50/20 border-l">${item.sWasteLoss > 0 ? '-'+item.sWasteLoss.toLocaleString() : '0'}</td>
                    <td class="p-2 text-right text-orange-600 bg-orange-50/20">${item.aWasteLoss > 0 ? '-'+item.aWasteLoss.toLocaleString() : '0'}</td>
                    <td class="p-2 text-right text-red-600 bg-red-50/20 border-l">${item.sStockoutLoss > 0 ? '-'+item.sStockoutLoss.toLocaleString() : '0'}</td>
                    <td class="p-2 text-right text-red-600 bg-red-50/20">${item.aStockoutLoss > 0 ? '-'+item.aStockoutLoss.toLocaleString() : '0'}</td>
                    <td class="p-2 text-right font-bold text-indigo-700 bg-indigo-50/50 border-l">${item.sNet.toLocaleString()}</td>
                    <td class="p-2 text-right font-bold text-indigo-700 bg-indigo-50/50">${item.aNet.toLocaleString()}</td>
                `;
                tbody.appendChild(tr);
            });

            document.getElementById('dailyDetailModal').classList.remove('hidden');
        }


        function getSelectedDayExportData() {
            const record = (STATE.closingHistory || [])[selectedDailyDetailIndex];
            if (!record) { alert('먼저 일자별 영업 마감 내역에서 상세보기를 선택해 주세요.'); return null; }
            const headers = ['품목코드','품목명','등급','학생 판매수량','AI 판매수량','학생 폐기손실(원)','AI 폐기손실(원)','학생 품절손실(원)','AI 품절손실(원)','학생 순이익(원)','AI 순이익(원)','AI 순이익 차이(원)'];
            const rows = (record.details || []).map(d => [d.id, d.name, d.abc || '', Number(d.sSales)||0, Number(d.aSales)||0, Number(d.sWasteLoss)||0, Number(d.aWasteLoss)||0, Number(d.sStockoutLoss)||0, Number(d.aStockoutLoss)||0, Number(d.sNet)||0, Number(d.aNet)||0, (Number(d.aNet)||0)-(Number(d.sNet)||0)]);
            const sum = key => rows.reduce((n,r)=>n+(Number(r[key])||0),0);
            const summary = [
                ['영업일', record.date], ['학생 매출(원)', Number(record.sRev)||0], ['AI 매출(원)', Number(record.aRev)||0],
                ['학생 폐기손실(원)', Number(record.sWaste)||0], ['AI 폐기손실(원)', Number(record.aWaste)||0],
                ['학생 품절손실(원)', Number(record.sStockout)||0], ['AI 품절손실(원)', Number(record.aStockout)||0],
                ['학생 순이익(원)', Number(record.sNet)||0], ['AI 순이익(원)', Number(record.aNet)||0],
                ['AI 순이익 차이(원)', (Number(record.aNet)||0)-(Number(record.sNet)||0)]
            ];
            return {record, headers, rows, summary};
        }

        function exportSelectedDayExcel() {
            const data = getSelectedDayExportData(); if (!data) return;
            const filename = `편의점_${data.record.date}_일별영업보고서`;
            const aoa = [['날짜별 영업 보고서'], ...data.summary, [], data.headers, ...data.rows];
            if (typeof XLSX !== 'undefined') {
                const wb = XLSX.utils.book_new();
                const ws = XLSX.utils.aoa_to_sheet(aoa);
                ws['!cols'] = [{wch:20},{wch:22},{wch:12},{wch:14},{wch:14},{wch:18},{wch:18},{wch:18},{wch:18},{wch:16},{wch:16},{wch:18}];
                XLSX.utils.book_append_sheet(wb, ws, '일별 상세');
                XLSX.writeFile(wb, `${filename}.xlsx`);
            } else {
                const csv = '\ufeff' + aoa.map(row=>row.map(v=>'"'+String(v??'').replace(/"/g,'""')+'"').join(',')).join('\r\n');
                const url = URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8;'}));
                const a=document.createElement('a'); a.href=url; a.download=`${filename}.csv`; a.click(); URL.revokeObjectURL(url);
                alert('Excel 라이브러리를 불러오지 못해 CSV로 저장했습니다. 인터넷 연결 후 다시 시도하면 XLSX로 저장됩니다.');
            }
        }

        function exportSelectedDayPDF() {
            const data = getSelectedDayExportData(); if (!data) return;
            const esc = value => String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
            const money = n => Math.round(Number(n)||0).toLocaleString('ko-KR')+'원';
            const popup = window.open('', '_blank');
            if (!popup) { alert('PDF 인쇄 창이 차단되었습니다. 브라우저에서 팝업을 허용한 뒤 다시 눌러주세요.'); return; }
            const summaryRows = data.summary.map(r=>`<tr><th>${esc(r[0])}</th><td>${esc(typeof r[1]==='number'?money(r[1]):r[1])}</td></tr>`).join('');
            const detailRows = data.rows.map(r=>`<tr>${r.map((v,i)=>`<td>${esc(i>=5?money(v):v)}</td>`).join('')}</tr>`).join('');
            popup.document.write(`<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><title>${esc(data.record.date)} 일별 영업 보고서</title><style>@page{size:A4 landscape;margin:10mm}body{font-family:Arial,'Malgun Gothic',sans-serif;color:#1e293b;font-size:8px}h1{font-size:20px}h2{font-size:14px;margin-top:18px}table{border-collapse:collapse;width:100%;margin:8px 0 16px}th,td{border:1px solid #cbd5e1;padding:4px;text-align:right;white-space:nowrap}th{background:#f1f5f9}td:nth-child(1),td:nth-child(2),th{text-align:left}.summary{width:60%}tr{break-inside:avoid}</style></head><body><h1>${esc(data.record.date)} 일별 영업 성과 보고서</h1><p>생성일: ${new Date().toLocaleDateString('ko-KR')}</p><h2>하루 영업 요약</h2><table class="summary">${summaryRows}</table><h2>품목별 상세 내역</h2><table><thead><tr>${data.headers.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${detailRows}</tbody></table><script>window.onload=()=>{window.focus();window.print();};window.onafterprint=()=>window.close();<\/script></body></html>`);
            popup.document.close();
        }

        function closeDailyDetailModal() { document.getElementById('dailyDetailModal').classList.add('hidden'); }

        // 모달창에 당일 결과 렌더링
        function showClosingReport(dateStr, sRev, sWaste, sStockout, aRev, aWaste, aStockout) {
            document.getElementById('closingDateLabel').innerText = `${dateStr} 영업 마감 결과`;
            document.getElementById('c_stuRev').innerText = `${sRev.toLocaleString()}원`;
            document.getElementById('c_stuWaste').innerText = `${sWaste.toLocaleString()}원`;
            document.getElementById('c_stuStockout').innerText = `${sStockout.toLocaleString()}원`;
            let sNet = sRev - sWaste - sStockout;
            document.getElementById('c_stuNet').innerText = `${sNet.toLocaleString()}원`;

            document.getElementById('c_aiRev').innerText = `${aRev.toLocaleString()}원`;
            document.getElementById('c_aiWaste').innerText = `${aWaste.toLocaleString()}원`;
            document.getElementById('c_aiStockout').innerText = `${aStockout.toLocaleString()}원`;
            let aNet = aRev - aWaste - aStockout;
            document.getElementById('c_aiNet').innerText = `${aNet.toLocaleString()}원`;

            let diff = aNet - sNet;
            let diffMsg = diff > 0 
                ? `🤖 AI가 학생보다 <strong>${diff.toLocaleString()}원</strong> 더 많은 순이익을 냈어!`
                : (diff < 0 ? `🧑‍🎓 대박! 학생이 AI보다 <strong>${Math.abs(diff).toLocaleString()}원</strong> 더 이익을 냈네!` : `⚖️ 오, AI와 학생의 이익이 완벽히 똑같아!`);
            
            document.getElementById('c_conclusion').innerHTML = diffMsg;
            document.getElementById('closingModal').classList.remove('hidden');
        }

        // 모달 닫고 내일 기초재고 이월 처리
        function moveToNextDay(skipRender = false) {
            document.getElementById('closingModal').classList.add('hidden');

            STATE.items.forEach(item => {
                item.userOrderToday = item.userOrderTomorrow;
                item.userOrderTomorrow = 0; 
                
                item.orderToday = item.orderTomorrow;
                
                item.forecastD0 = item.forecastD1;
                item.forecastD1 = item.forecastD2;
                item.forecastD2 = item.forecastD3;
                item.forecastD3 = Math.round(item.avgSales * (0.9 + Math.random() * 0.2));
                
                let estTomorrow = Math.max(0, item.aiStock + item.orderTomorrow - item.forecastD0);
                let needed = item.targetStock - estTomorrow;
                item.orderTomorrow = needed > 0 ? Math.ceil(needed) : 0;
            });

            STATE.simDaysElapsed++;
            STATE.currentDateObj.setDate(STATE.currentDateObj.getDate() + 1);
            
            const year = STATE.currentDateObj.getFullYear();
            const month = String(STATE.currentDateObj.getMonth() + 1).padStart(2, '0');
            const day = String(STATE.currentDateObj.getDate()).padStart(2, '0');
            const dayName = ['일', '월', '화', '수', '목', '금', '토'][STATE.currentDateObj.getDay()];
            STATE.currentDateStr = `${year}-${month}-${day}`;
            
            document.getElementById('currentSimDateDisplay').innerText = `${STATE.currentDateStr} (${dayName})`;
            document.getElementById('simDaysCount').innerText = `${STATE.simDaysElapsed} 일차`;
            document.getElementById('simTotalRevenue').innerText = `${STATE.simTotalRevenue.toLocaleString()} 원`;
            document.getElementById('simTotalWaste').innerText = `${STATE.simTotalWaste.toLocaleString()} 원`;
            document.getElementById('simTotalOrdersCount').innerText = `${STATE.simTotalOrdersCount} 건`;

            if(!skipRender) {
                renderDashboard();
                renderRiskTable();
                renderOrderTable();
                renderForecastTable();
                showToast(`영업일이 ${STATE.currentDateStr} 로 변경되고 기초재고가 이월되었어!`);
            }
        }

        // 전체 초기화
        function resetSimulationData() {
            STATE.simDaysElapsed = 0; STATE.simTotalRevenue = 0; STATE.simTotalWaste = 0; STATE.simTotalOrdersCount = 0;
            STATE.currentDateObj = new Date(2026, 9, 1); 
            STATE.currentDateStr = '2026-10-01';
            STATE.closingHistory = []; // 배열 초기화
            
            document.getElementById('currentSimDateDisplay').innerText = `2026-10-01 (목)`;
            document.getElementById('simDaysCount').innerText = `0 일차`;
            document.getElementById('simTotalRevenue').innerText = `0 원`;
            document.getElementById('simTotalWaste').innerText = `0 원`;
            document.getElementById('simTotalOrdersCount').innerText = `0 건`;
            
            initFullSKUDataset(); 
            renderDashboard();
            renderRiskTable();
            renderOrderTable(); 
            renderForecastTable(); 
            renderClosingHistoryTable(); 
            
            showToast('모든 누적 데이터와 마감 내역이 처음으로 완전히 초기화되었습니다.');
        }

        function switchOrderSubTab(subTab) {
            const orderPanel = document.getElementById('order-management-panel');
            const riskPanel = document.getElementById('risk-management-panel');
            const orderBtn = document.getElementById('orderSubBtn-order');
            const riskBtn = document.getElementById('orderSubBtn-risk');
            if (!orderPanel || !riskPanel) return;
            const showOrder = subTab !== 'risk';
            orderPanel.classList.toggle('hidden', !showOrder);
            riskPanel.classList.toggle('hidden', showOrder);
            if (orderBtn) orderBtn.className = showOrder
                ? 'rounded-lg bg-emerald-600 text-white px-4 py-3 text-sm font-extrabold shadow-sm flex items-center justify-center gap-2'
                : 'rounded-lg bg-white text-slate-700 hover:bg-emerald-50 border border-slate-200 px-4 py-3 text-sm font-bold flex items-center justify-center gap-2';
            if (riskBtn) riskBtn.className = !showOrder
                ? 'rounded-lg bg-amber-500 text-white px-4 py-3 text-sm font-extrabold shadow-sm flex items-center justify-center gap-2'
                : 'rounded-lg bg-white text-slate-700 hover:bg-amber-50 border border-slate-200 px-4 py-3 text-sm font-bold flex items-center justify-center gap-2';
            if (!showOrder) renderRiskTable();
            else renderOrderTable();
        }

        function switchTab(tabId) {
            // 위험관리 바로가기에서는 통합 탭 안의 재고 위험관리 하위 메뉴를 엽니다.
            const requestedRisk = tabId === 'risk';
            if (requestedRisk) tabId = 'order';
            STATE.activeTab = tabId;
            document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
            document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
            const targetTab = document.getElementById(`tab-${tabId}`);
            if (targetTab) targetTab.classList.remove('hidden');
            if(document.getElementById(`tabBtn-${tabId}`)) document.getElementById(`tabBtn-${tabId}`).classList.add('active');
            
            if (tabId === 'dashboard') renderDashboard();
            if (tabId === 'order') {
                switchOrderSubTab(requestedRisk ? 'risk' : 'order');
                renderOrderTable();
                renderRiskTable();
            }
            if (tabId === 'forecast') renderForecastTable();
            if (tabId === 'simulator') renderClosingHistoryTable();
        }

        function showToast(text) {
            const toast = document.getElementById('toastMessage');
            document.getElementById('toastText').innerText = text;
            toast.classList.remove('translate-y-20', 'opacity-0'); toast.classList.add('translate-y-0', 'opacity-100');
            setTimeout(() => { toast.classList.remove('translate-y-0', 'opacity-100'); toast.classList.add('translate-y-20', 'opacity-0'); }, 3000);
        }

        function openSkuModal(id) { 
            const item = STATE.items.find(i => i.id === id);
            if(!item) return;
            document.getElementById('modalSkuCode').innerText = item.id;
            document.getElementById('modalSkuName').innerText = item.name;
            document.getElementById('modalCat').innerText = `${item.category} (${item.subCat})`;
            document.getElementById('modalAbc').innerText = `${item.abc} 등급`;
            document.getElementById('modalPriceCost').innerText = `${item.price.toLocaleString()}원 / ${item.cost.toLocaleString()}원`;
            document.getElementById('modalShelfStorage').innerText = `${item.shelfLifeDays}일 / ${item.storage}`;
            document.getElementById('modalMinDisplay').innerText = `${item.minDisplay}개`;
            document.getElementById('modalBeginningStock').innerText = `${item.stock}개`;
            document.getElementById('modalSafetyStock').innerText = `${item.safetyStock}개`;
            document.getElementById('modalEndingStock').innerText = `${getEndingStock(item)}개`;
            document.getElementById('modalRop').innerText = `${item.rop}개`;
            document.getElementById('modalTargetStock').innerText = `${item.targetStock}개`;
            document.getElementById('skuModal').classList.remove('hidden'); 
        }
        function closeSkuModal() { document.getElementById('skuModal').classList.add('hidden'); }

        window.onload = function() {
            initFullSKUDataset(); 
            renderDashboard();
            renderRiskTable();
            renderOrderTable(); 
            renderForecastTable(); 
            renderClosingHistoryTable();
            switchTab('order');
        };
    </script>
</body>
</html>
