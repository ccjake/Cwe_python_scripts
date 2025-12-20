document.addEventListener('DOMContentLoaded', () => {
    // --- Data Pools --- 56人
    const CW = [
        "王俊文", "蒋紫融", "孙晓泉", "陈眀畅", "唐澜瑄", "张崧", "潘正阳", "舒心",
        "孟甜", "赵玉花", "赵旭", "黄晓菁", "梁永媛", "张炜凡", "卓霖", "王梓襄",
        "叶百超", "邬龙贞", "涂天昱", "田昕", "李岱明", "姜宣兆", "肖倩", "伍漫瑜",
        "陈妤哲", "刘铉达", "王子豪", "冯泽", "陈笑天", "王清岚", "富家丹", "王斯雯",
        "刘岩", "朱诗倩", "戴灵瑶", "何佳", "宁霞", "严露冰", "陈彤", "韩冰雪",
        "吕艳潇", "顾佳乐", "王敏烨", "彭涌", "张泽宇", "郑睿", "李雨昕", "任林",
        "陈叶秋子", "朱振瑶", "陈微笑", "李思琦", "姚舒雯", "史淑彤", "王诗博", "候利利"
    ]

    ///////////////////////////////
    // 12人
    const cw_not = ['王思博', '郑希', '王欣', '李佳佳', '朱丽霞', '左霄莹', '吴帅',
        '李鹏刚', '王志鹏', '邓浩', '许栩', '傅峥'];

    const special_cw = ['王博', '王伟业', '胡星', '贾超']
    const PRIZE_CONFIG = [
        { name: "参与奖", duration: 2 },
        { name: "三等奖", duration: 2 },
        { name: "二等奖", duration: 3 },
        { name: "一等奖", duration: 3 },
        { name: "特等奖", duration: 5 },
        // { name: "参与奖", duration: 0 },
        // { name: "三等奖", duration: 0 },
        // { name: "二等奖", duration: 0 },
        // { name: "一等奖", duration: 0 },
        // { name: "特等奖", duration: 0 }
    ];

    // --- State ---
    let drawnWinners = new Set(); // Track all winners to prevent duplicates
    let prizeCounters = {};       // Track count of winners per prize type
    let historyLog = [];          // Ordered list of {name, prize} for display

    let isRolling = false;
    let rollInterval;
    let currentPrizeIndex = 0; // Controlled by Next Prize button

    // --- DOM Elements ---
    const els = {
        currentPrizeLabel: document.getElementById('currentPrizeLabel'),
        displayArea: document.getElementById('displayArea'),
        startBtn: document.getElementById('startBtn'),
        nextPrizeBtn: document.getElementById('nextPrizeBtn'),
        winnersList: document.getElementById('winnersList'),
        historyPanel: document.getElementById('historyPanel'),
        toggleSidebarBtn: document.getElementById('toggleSidebarBtn'),
        closeSidebarBtn: document.getElementById('closeSidebarBtn'),
        confettiCanvas: document.getElementById('confettiCanvas'),
        secretReset: document.getElementById('secretReset')
    };

    // --- Initialization ---
    function init() {
        // 1. Load stored data first
        loadData();

        // 2. Initialize/Fill missing counters based on config
        PRIZE_CONFIG.forEach(p => {
            if (typeof prizeCounters[p.name] !== 'number') {
                prizeCounters[p.name] = 0;
            }
        });

        renderHistory();
        updateUIState();
        resizeCanvas();
        window.addEventListener('resize', resizeCanvas);
    }

    // --- Core Logic ---
    function getDrawPool() {
        const currentPrize = PRIZE_CONFIG[currentPrizeIndex].name;
        let basePool = [];

        // Pool Selection Logic
        if (currentPrize === "参与奖" || currentPrize === "三等奖") {
            // Combined Pool for lower tier prizes
            basePool = [...CW, ...cw_not];
        } else {
            // CW Only for higher tier prizes (Second, First, Grand)
            basePool = [...CW];
        }

        // Filter out rigged winners (special_cw) and already drawn winners
        // ensuring no duplicates and preserving rigged slots
        return basePool.filter(name =>
            !special_cw.includes(name) &&
            !drawnWinners.has(name)
        );
    }

    function determineNextWinner() {
        const currentPrize = PRIZE_CONFIG[currentPrizeIndex].name;
        const currentCount = (prizeCounters[currentPrize] || 0) + 1; // 1-based index for next draw

        console.log(`Drawing for: ${currentPrize}, Count: ${currentCount}`);

        // --- Special Rules ---
        // 1. Participation Prize: 7th -> Wang Bo, 9th -> Wang Weiye
        if (currentPrize === "参与奖") {
            if (currentCount === 7) return '王博';
            if (currentCount === 9) return '王伟业';
        }

        // 2. Third Prize: 4th -> Hu Xing
        if (currentPrize === "三等奖") {
            if (currentCount === 4) return '胡星';
        }

        // 3. Second Prize: 3rd -> Jia Chao
        if (currentPrize === "二等奖") {
            if (currentCount === 3) return '贾超';
        }

        // --- Normal Draw ---
        const pool = getDrawPool();
        if (pool.length === 0) return null;

        const randomIndex = Math.floor(Math.random() * pool.length);
        return pool[randomIndex];
    }

    function nextPrize() {
        if (isRolling) return;

        currentPrizeIndex++;
        if (currentPrizeIndex >= PRIZE_CONFIG.length) {
            currentPrizeIndex = 0;
        }

        // Trigger Upgrade Animation
        if (els.currentPrizeLabel) {
            els.currentPrizeLabel.classList.remove('animate-upgrade');
            void els.currentPrizeLabel.offsetWidth; // Force reflow
            els.currentPrizeLabel.classList.add('animate-upgrade');
        }

        updateUIState();
    }

    function startLottery() {
        if (isRolling) return;

        // Check if we can draw
        const winner = determineNextWinner();
        if (!winner) {
            alert("没有更多参与者可供抽奖！");
            return;
        }

        isRolling = true;
        updateUIState();

        // 1. Rolling Animation
        // Show random names from the remaining pool + special list (just for visual effect)
        const visualPool = CW.filter(n => !drawnWinners.has(n));

        rollInterval = setInterval(() => {
            const randomName = visualPool[Math.floor(Math.random() * visualPool.length)];
            els.displayArea.textContent = randomName;
        }, 50);

        // 2. Auto-Stop
        const currentPrizeConfig = PRIZE_CONFIG[currentPrizeIndex];
        setTimeout(() => {
            stopLottery(winner);
        }, currentPrizeConfig.duration * 1000);
    }

    function stopLottery(winnerName) {
        clearInterval(rollInterval);
        isRolling = false;

        // Update State
        drawnWinners.add(winnerName);
        const currentPrize = PRIZE_CONFIG[currentPrizeIndex].name;

        console.log("Before Increment:", JSON.stringify(prizeCounters));

        // Safe Increment
        if (typeof prizeCounters[currentPrize] !== 'number') {
            prizeCounters[currentPrize] = 0;
        }
        prizeCounters[currentPrize]++;

        console.log("After Increment:", JSON.stringify(prizeCounters));

        historyLog.push({ name: winnerName, prize: currentPrize });

        // Visuals
        els.displayArea.textContent = winnerName;
        els.displayArea.classList.remove('bounce-in');
        void els.displayArea.offsetWidth;
        els.displayArea.classList.add('bounce-in');

        fireConfetti();

        // Save & Render
        saveData();
        renderHistory();
        updateUIState();
    }

    function updateUIState() {
        const config = PRIZE_CONFIG[currentPrizeIndex];

        // Buttons
        els.startBtn.disabled = isRolling;
        els.nextPrizeBtn.disabled = isRolling;

        // Labels
        if (els.currentPrizeLabel) {
            const count = prizeCounters[config.name] || 0;
            els.currentPrizeLabel.textContent = `当前奖项：${config.name} (已抽取: ${count})`;
        }

        if (isRolling) {
            els.startBtn.textContent = "抽奖中...";
        } else {
            els.startBtn.textContent = "开始抽奖";
        }
    }

    // --- Persistence ---
    function saveData() {
        const data = {
            drawnWinners: Array.from(drawnWinners),
            prizeCounters,
            currentPrizeIndex,
            historyLog
        };
        localStorage.setItem('letto_v3_data', JSON.stringify(data));
    }

    function loadData() {
        const stored = localStorage.getItem('letto_v3_data');
        if (stored) {
            try {
                const data = JSON.parse(stored);
                if (data.drawnWinners) {
                    drawnWinners = new Set(data.drawnWinners);
                    // Validate / Sanitize Prize Counters
                    prizeCounters = {};
                    const loadedCounters = data.prizeCounters || {};

                    PRIZE_CONFIG.forEach(p => {
                        const val = loadedCounters[p.name];
                        // If value is a valid number, keep it. Otherwise 0.
                        if (typeof val === 'number' && !isNaN(val)) {
                            prizeCounters[p.name] = val;
                        } else {
                            prizeCounters[p.name] = 0;
                        }
                    });

                    currentPrizeIndex = data.currentPrizeIndex || 0;
                    historyLog = data.historyLog || [];
                }
            } catch (e) {
                console.error("Data Parse Error, resetting", e);
                // On error, let init() handle defaults
            }
        }
    }

    function renderHistory() {
        els.winnersList.innerHTML = '';
        els.toggleSidebarBtn.textContent = `幸运名单(${drawnWinners.size})`;

        // Use historyLog for order
        [...historyLog].reverse().forEach(w => {
            const item = document.createElement('div');
            item.className = 'history-item';
            item.innerHTML = `
                <span class="history-prize">${w.prize}</span>
                <span class="history-name">${w.name}</span>
            `;
            els.winnersList.appendChild(item);
        });
    }


    // --- Confetti (Using canvas-confetti library) ---
    function fireConfetti() {
        const duration = 3000;
        const animationEnd = Date.now() + duration;
        const defaults = { startVelocity: 30, spread: 360, ticks: 60, zIndex: 999 };

        const randomInRange = (min, max) => Math.random() * (max - min) + min;

        const interval = setInterval(function () {
            const timeLeft = animationEnd - Date.now();

            if (timeLeft <= 0) {
                return clearInterval(interval);
            }

            const particleCount = 50 * (timeLeft / duration);

            // Random bursts from different locations
            confetti({ ...defaults, particleCount, origin: { x: randomInRange(0.1, 0.3), y: Math.random() - 0.2 } });
            confetti({ ...defaults, particleCount, origin: { x: randomInRange(0.7, 0.9), y: Math.random() - 0.2 } });
        }, 250);

        // Also a big burst from center
        confetti({
            particleCount: 150,
            spread: 100,
            origin: { y: 0.6 },
            colors: ['#F1C40F', '#D91E18', '#FFFFFF']
        });
    }

    function resizeCanvas() {
        // No-op or remove listener usage
    }

    // --- Listeners ---
    els.startBtn.addEventListener('click', startLottery);
    els.nextPrizeBtn.addEventListener('click', nextPrize);

    // Sidebar Toggles
    els.toggleSidebarBtn.addEventListener('click', () => {
        els.historyPanel.classList.add('show');
    });
    els.closeSidebarBtn.addEventListener('click', () => {
        els.historyPanel.classList.remove('show');
    });

    // Reset Logic Shared
    const handleReset = () => {
        if (confirm("⚠️ 危险操作 ⚠️\n\n确定要重置所有抽奖数据吗？\n这将清空历史记录并重新开始！")) {
            localStorage.removeItem('letto_v3_data');
            location.reload();
        }
    };

    document.getElementById('sidebarTitleReset').addEventListener('click', handleReset);
    els.secretReset.addEventListener('click', handleReset);

    init();
});

