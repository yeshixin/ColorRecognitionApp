let questions = [];
let currentQuestionIndex = 0;
let startTime, endTime;
let responseTimes = [];
let correctAnswers = 0;
let totalQuestions = 0;
let wrongAnswers = 0;

const startScreen = document.getElementById('start-screen');
const bufferScreen = document.getElementById('buffer-screen');
const questionScreen = document.getElementById('question-screen');
const optionsScreen = document.getElementById('options-screen');
const questionText = document.getElementById('question-text');
const option1 = document.getElementById('option1');
const option2 = document.getElementById('option2');

// 載入 Excel 檔案
document.getElementById('file-input').addEventListener('change', (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();

    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet);

        // 將 Excel 資料轉為題目格式
        questions = json.map(row => ({
            word: row.Word,
            options: [row.Option1, row.Option2],
            answer: row.Answer,
            backgroundColor: row.BackgroundColor || "#FFFFFF", // 預設白色
            textColor: row.TextColor || "#000000" // 預設黑色
        }));

        console.log('題目已載入:', questions);
    };

    reader.readAsArrayBuffer(file);
});

// 顯示開始畫面並顯示 START
startScreen.addEventListener('click', () => {
    document.getElementById('start-text').style.display = 'block';
    startScreen.style.display = 'none';
    bufferScreen.style.display = 'flex';
    setTimeout(startTest, 2000);
});

function startTest() {
    bufferScreen.style.display = 'none';
    loadNextQuestion();
}

function loadNextQuestion() {
    if (currentQuestionIndex >= questions.length) {
        return showResults();
    }

    const currentQuestion = questions[currentQuestionIndex];
    questionText.textContent = currentQuestion.word;
    questionScreen.style.display = 'flex';
    startTime = Date.now();

    setTimeout(() => {
        questionScreen.style.display = 'none';
        showOptions(currentQuestion);
    }, 2000); // 2秒後顯示選項
}

function showOptions(currentQuestion) {
    option1.style.backgroundColor = currentQuestion.backgroundColor;
    option2.style.backgroundColor = currentQuestion.backgroundColor;

    option1.style.color = currentQuestion.textColor;
    option2.style.color = currentQuestion.textColor;

    option1.textContent = currentQuestion.options[0];
    option2.textContent = currentQuestion.options[1];

    optionsScreen.style.display = 'flex';

    // 點擊事件處理
    option1.onclick = () => handleAnswerWrapper(currentQuestion, option1.textContent);
    option2.onclick = () => handleAnswerWrapper(currentQuestion, option2.textContent);

    // 鍵盤事件監聽
    document.addEventListener('keydown', handleKeyPress);

    function handleKeyPress(event) {
        if (event.key === 'ArrowLeft') {
            option1.click();
        } else if (event.key === 'ArrowRight') {
            option2.click();
        }
    }

    function handleAnswerWrapper(currentQuestion, selectedOption) {
        handleAnswer(currentQuestion, selectedOption);
        document.removeEventListener('keydown', handleKeyPress);
    }
}

function handleAnswer(currentQuestion, selectedOption) {
    endTime = Date.now();
    const reactionTime = (endTime - startTime) / 1000;
    responseTimes.push(reactionTime);

    totalQuestions++;
    if (selectedOption === currentQuestion.answer) {
        correctAnswers++;
    } else {
        wrongAnswers++;
    }

    optionsScreen.style.display = 'none';
    currentQuestionIndex++;
    loadNextQuestion();
}

function showResults() {
    const accuracy = (correctAnswers / totalQuestions) * 100;
    const errorRate = (wrongAnswers / totalQuestions) * 100;

    console.log('反應時間:', responseTimes);
    console.log('錯誤率:', errorRate);

    exportToExcel();
}

function exportToExcel() {
    const wb = XLSX.utils.book_new();
    const wsData = [
        ['Question', 'Reaction Time (s)', 'Correct Answer', 'Selected Answer', 'Is Correct?'],
        ...questions.map((q, index) => [
            q.word,
            responseTimes[index] || 0, // 確保反應時間有值
            q.answer,
            q.options.includes(q.answer) ? q.answer : "未選擇",
            q.options.includes(q.answer) ? 'Yes' : 'No'
        ])
    ];

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, "Results");

    XLSX.writeFile(wb, "TestResults.xlsx");
}
