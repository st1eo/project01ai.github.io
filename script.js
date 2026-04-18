const dayLabels = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"];
const analysisHistoryKey = "digital-health-history";
const historyLimit = 6;

const defaultNotification =
  "«Сон и активность ухудшаются уже несколько дней. Возможны признаки переутомления. Постарайтесь сократить вечернюю нагрузку и увеличить время сна»";

const defaultState = {
  age: 21,
  sleep: 5,
  activity: 4800,
};

const ageInput = document.getElementById("ageInput");
const sleepInput = document.getElementById("sleepInput");
const activityInput = document.getElementById("activityInput");

const analyzeButton = document.getElementById("analyzeButton");
const resetButton = document.getElementById("resetButton");
const clearHistoryButton = document.getElementById("clearHistoryButton");

const statusBadge = document.getElementById("statusBadge");
const resultTitle = document.getElementById("resultTitle");
const resultText = document.getElementById("resultText");
const notificationText = document.getElementById("notificationText");
const trendGrid = document.getElementById("trendGrid");
const trendDescription = document.getElementById("trendDescription");
const findingsList = document.getElementById("findingsList");
const recommendationsList = document.getElementById("recommendationsList");

const ageProfileText = document.getElementById("ageProfileText");
const sleepNormText = document.getElementById("sleepNormText");
const activityNormText = document.getElementById("activityNormText");
const personalConclusion = document.getElementById("personalConclusion");

const currentAgeDisplay = document.getElementById("currentAgeDisplay");
const currentSleepDisplay = document.getElementById("currentSleepDisplay");
const currentActivityDisplay = document.getElementById("currentActivityDisplay");
const currentFatigueDisplay = document.getElementById("currentFatigueDisplay");

const currentSleepBar = document.getElementById("currentSleepBar");
const currentActivityBar = document.getElementById("currentActivityBar");
const currentFatigueBar = document.getElementById("currentFatigueBar");
const historyList = document.getElementById("historyList");
const exportPanel = document.getElementById("exportPanel");
const downloadReportButton = document.getElementById("downloadReportButton");

let latestReport = null;

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function formatSteps(value) {
  return `${Math.round(value).toLocaleString("ru-RU")} шагов`;
}

function formatAge(age) {
  const lastTwo = age % 100;
  const lastOne = age % 10;

  if (lastTwo >= 11 && lastTwo <= 14) {
    return `${age} лет`;
  }

  if (lastOne === 1) {
    return `${age} год`;
  }

  if (lastOne >= 2 && lastOne <= 4) {
    return `${age} года`;
  }

  return `${age} лет`;
}

function formatDateTime(timestamp) {
  return new Intl.DateTimeFormat("ru-RU", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  }).format(new Date(timestamp));
}

function escapeXml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function formatDecimal(value) {
  return value.toFixed(1).replace(".", ",");
}

function formatWhole(value) {
  return Math.round(value).toLocaleString("ru-RU");
}

function readHistory() {
  try {
    const raw = localStorage.getItem(analysisHistoryKey);
    const parsed = raw ? JSON.parse(raw) : [];

    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    return [];
  }
}

function saveHistory(history) {
  try {
    localStorage.setItem(analysisHistoryKey, JSON.stringify(history));
  } catch (error) {
    // Если localStorage недоступен, интерфейс продолжит работать без истории.
  }
}

function renderHistory() {
  const history = readHistory();

  historyList.innerHTML = "";

  if (!history.length) {
    historyList.innerHTML = `
      <p class="history-empty">
        История пока пуста. Запустите анализ, и здесь появятся последние результаты.
      </p>
    `;
    return;
  }

  history.forEach((entry) => {
    const item = document.createElement("article");
    item.className = "history-item";

    item.innerHTML = `
      <div class="history-item-top">
        <strong>${entry.badgeText}</strong>
        <span class="history-date">${formatDateTime(entry.createdAt)}</span>
      </div>
      <div class="history-metrics">
        <div class="history-metric">
          <span>Возраст</span>
          <strong>${formatAge(entry.age)}</strong>
        </div>
        <div class="history-metric">
          <span>Сон</span>
          <strong>${entry.sleep.toFixed(1)} ч</strong>
        </div>
        <div class="history-metric">
          <span>Активность</span>
          <strong>${formatSteps(entry.activity)}</strong>
        </div>
        <div class="history-metric">
          <span>Усталость</span>
          <strong>${entry.fatigue}/10</strong>
        </div>
      </div>
      <p class="history-conclusion">${entry.conclusion}</p>
    `;

    historyList.appendChild(item);
  });
}

function storeAnalysis(entry) {
  const history = readHistory();
  const nextHistory = [entry, ...history].slice(0, historyLimit);

  saveHistory(nextHistory);
  renderHistory();
}

function getNormsByAge(age) {
  if (age < 18) {
    return {
      profile: "Подростковый профиль",
      minSleep: 8,
      maxSleep: 10,
      activity: 10000,
      fatigue: "2-3/10",
    };
  }

  if (age <= 25) {
    return {
      profile: "Молодой взрослый профиль",
      minSleep: 7,
      maxSleep: 9,
      activity: 8000,
      fatigue: "2-3/10",
    };
  }

  if (age <= 45) {
    return {
      profile: "Взрослый профиль",
      minSleep: 7,
      maxSleep: 9,
      activity: 7500,
      fatigue: "2-4/10",
    };
  }

  if (age <= 64) {
    return {
      profile: "Профиль 46-64 лет",
      minSleep: 7,
      maxSleep: 8,
      activity: 7000,
      fatigue: "2-4/10",
    };
  }

  return {
    profile: "Профиль 65+",
    minSleep: 7,
    maxSleep: 8,
    activity: 6000,
    fatigue: "2-4/10",
  };
}

function estimateFatigue(sleep, activity, norms) {
  let score = 2;

  if (sleep < norms.minSleep) {
    score += Math.ceil((norms.minSleep - sleep) * 1.6);
  }

  if (activity < norms.activity) {
    score += Math.ceil(((norms.activity - activity) / norms.activity) * 4);
  }

  return clamp(score, 2, 10);
}

function setListItems(container, items) {
  container.innerHTML = "";

  items.forEach((item) => {
    const li = document.createElement("li");
    li.textContent = item;
    container.appendChild(li);
  });
}

function createTrendCard(dayData) {
  const item = document.createElement("article");
  item.className = "trend-item";

  const sleepPercent = clamp((dayData.sleep / 10) * 100, 8, 100);
  const activityPercent = clamp((dayData.activity / 12000) * 100, 8, 100);
  const fatiguePercent = clamp((dayData.fatigue / 10) * 100, 8, 100);

  item.innerHTML = `
    <div class="trend-day">${dayData.day}</div>
    <div class="trend-stat">
      <span><strong>Сон</strong><b>${dayData.sleep.toFixed(1)} ч</b></span>
      <div class="bar"><div class="bar-fill sleep-fill" style="width: ${sleepPercent}%"></div></div>
    </div>
    <div class="trend-stat">
      <span><strong>Активность</strong><b>${Math.round(dayData.activity)}</b></span>
      <div class="bar"><div class="bar-fill activity-fill" style="width: ${activityPercent}%"></div></div>
    </div>
    <div class="trend-stat">
      <span><strong>Усталость</strong><b>${dayData.fatigue}/10</b></span>
      <div class="bar"><div class="bar-fill fatigue-fill" style="width: ${fatiguePercent}%"></div></div>
    </div>
  `;

  return item;
}

function buildWeeklyForecast(age, sleep, activity) {
  const norms = getNormsByAge(age);
  const sleepGap = Math.max(0, norms.minSleep - sleep);
  const activityGap = Math.max(0, norms.activity - activity);

  const sleepShift = sleepGap > 0 ? clamp(sleepGap * 0.7, 0.6, 1.4) : -0.3;
  const activityShift = activityGap > 0 ? clamp(activityGap * 0.45, 400, 2200) : -500;

  return dayLabels.map((day, index) => {
    const factor = dayLabels.length === 1 ? 0 : index / (dayLabels.length - 1);
    const projectedSleep = clamp(sleep + sleepShift - sleepShift * factor, 3, 12);
    const projectedActivity = clamp(activity + activityShift - activityShift * factor, 1000, 30000);
    const projectedFatigue = estimateFatigue(projectedSleep, projectedActivity, norms);

    return {
      day,
      sleep: projectedSleep,
      activity: projectedActivity,
      fatigue: projectedFatigue,
    };
  });
}

function renderTrend(weekData) {
  trendGrid.innerHTML = "";
  weekData.forEach((dayData) => trendGrid.appendChild(createTrendCard(dayData)));
}

function updateReferenceCards(age, norms) {
  ageProfileText.textContent = `${norms.profile}, ${formatAge(age)}`;
  sleepNormText.textContent = `${norms.minSleep}-${norms.maxSleep} ч`;
  activityNormText.textContent = formatSteps(norms.activity);
}

function updateCurrentCard(age, sleep, activity, fatigue) {
  currentAgeDisplay.textContent = formatAge(age);
  currentSleepDisplay.textContent = `${sleep.toFixed(1)} ч`;
  currentActivityDisplay.textContent = formatSteps(activity);
  currentFatigueDisplay.textContent = `${fatigue}/10`;

  currentSleepBar.style.width = `${clamp((sleep / 10) * 100, 8, 100)}%`;
  currentActivityBar.style.width = `${clamp((activity / 12000) * 100, 8, 100)}%`;
  currentFatigueBar.style.width = `${clamp((fatigue / 10) * 100, 8, 100)}%`;
}

function validateInput(age, sleep, activity) {
  return Number.isFinite(age) && Number.isFinite(sleep) && Number.isFinite(activity) && age >= 14 && sleep > 0 && activity >= 0;
}

function getStressLabel(fatigue) {
  if (fatigue >= 8) {
    return "Очень высокий";
  }

  if (fatigue >= 6) {
    return "Высокий";
  }

  if (fatigue >= 4) {
    return "Средний";
  }

  return "Низкий";
}

function getRecoveryLabel(fatigue) {
  if (fatigue >= 8) {
    return "Критическое";
  }

  if (fatigue >= 7) {
    return "Снижается";
  }

  if (fatigue >= 5) {
    return "Удовлетворит.";
  }

  return "Норма";
}

function getRiskInfo(dayData, norms) {
  const sleepDeficit = Math.max(0, norms.minSleep - dayData.sleep);
  const activityPercent = Math.round((Math.max(0, norms.activity - dayData.activity) / norms.activity) * 100);

  if (dayData.fatigue >= 8 || sleepDeficit >= 2 || activityPercent >= 45) {
    return { label: "Критич.", rowStyle: "dangerRow", textStyle: "dangerText" };
  }

  if (dayData.fatigue >= 7 || sleepDeficit >= 1.5 || activityPercent >= 35) {
    return { label: "Высокий", rowStyle: "dangerRow", textStyle: "dangerText" };
  }

  if (dayData.fatigue >= 5 || sleepDeficit >= 1 || activityPercent >= 15) {
    return { label: "Средний", rowStyle: "warnRow", textStyle: "warnText" };
  }

  return { label: "Низкий", rowStyle: "goodRow", textStyle: "goodText" };
}

function getDailyRecommendation(dayData, norms) {
  const sleepDeficit = Math.max(0, norms.minSleep - dayData.sleep);
  const activityDeficit = Math.max(0, norms.activity - dayData.activity);

  if (sleepDeficit >= 2) {
    return `Срочно поднять сон до ${norms.minSleep}-${norms.maxSleep} ч.`;
  }

  if (dayData.fatigue >= 7) {
    return "Снизить вечернюю нагрузку и сделать разгрузочный вечер.";
  }

  if (activityDeficit >= 1800) {
    return `Добавить прогулку и приблизиться к ${formatSteps(norms.activity)}.`;
  }

  if (sleepDeficit > 0 || activityDeficit > 0) {
    return "Слегка скорректировать режим сна и движения.";
  }

  return "Режим выглядит стабильным.";
}

function xmlCell(value, styleId, options = {}) {
  const type = options.type || "String";
  const mergeAcross = Number.isInteger(options.mergeAcross) ? ` ss:MergeAcross="${options.mergeAcross}"` : "";

  if (value === null || value === undefined || value === "") {
    return `<Cell ss:StyleID="${styleId}"${mergeAcross}/>`;
  }

  return `<Cell ss:StyleID="${styleId}"${mergeAcross}><Data ss:Type="${type}">${escapeXml(value)}</Data></Cell>`;
}

function xmlRow(cells, options = {}) {
  const height = Number.isFinite(options.height) ? ` ss:Height="${options.height}" ss:AutoFitHeight="0"` : "";

  return `<Row${height}>${cells.join("")}</Row>`;
}

function buildRecommendationRows(reportData) {
  const {
    sleep,
    activity,
    fatigue,
    norms,
    badgeText,
  } = reportData;

  const sleepDeficit = Math.max(0, norms.minSleep - sleep);
  const activityDeficit = Math.max(0, norms.activity - activity);
  const activityPercent = Math.round((activityDeficit / norms.activity) * 100);
  const rows = [];

  if (sleepDeficit > 0) {
    rows.push({
      priority: sleepDeficit >= 2 ? "Срочно" : "Важно",
      category: "Сон",
      recommendation: `Лечь спать раньше и выйти на ${norms.minSleep}-${norms.maxSleep} часов сна.`,
      effect: "Восстановление концентрации и снижение признаков недосыпа.",
    });
  }

  if (fatigue >= 7) {
    rows.push({
      priority: "Срочно",
      category: "Нагрузка",
      recommendation: "Уменьшить вечернюю учебную нагрузку и убрать экран за 1 час до сна.",
      effect: "Снижение утомления и более глубокое восстановление ночью.",
    });
  }

  if (activityDeficit > 0) {
    rows.push({
      priority: activityPercent >= 35 ? "Важно" : "Полезно",
      category: "Активность",
      recommendation: `Добавить дневную прогулку и приблизиться к ${formatSteps(norms.activity)}.`,
      effect: "Улучшение кровообращения, бодрости и устойчивости к нагрузке.",
    });
  }

  rows.push({
    priority: badgeText === "Высокий риск" ? "Важно" : "Полезно",
    category: "Питание",
    recommendation: "Не пропускать приёмы пищи и поддерживать воду в течение дня.",
    effect: "Более стабильная энергия и меньшее чувство истощения вечером.",
  });

  rows.push({
    priority: "Полезно",
    category: "Планирование",
    recommendation: "Разбивать подготовку на интервалы 25/5 или 50/10 минут с короткими паузами.",
    effect: "Меньше перегрузки и лучшее удержание внимания во время учёбы.",
  });

  rows.push({
    priority: "Полезно",
    category: "Восстановление",
    recommendation: "Добавить 5-10 минут спокойного дыхания или отдыха перед сном.",
    effect: "Снижение внутреннего напряжения и более спокойное засыпание.",
  });

  return rows.slice(0, 6);
}

function getPriorityStyles(priority) {
  if (priority === "Срочно") {
    return {
      rowStyle: "dangerRow",
      textStyle: "dangerText",
    };
  }

  if (priority === "Важно") {
    return {
      rowStyle: "warnRow",
      textStyle: "warnText",
    };
  }

  return {
    rowStyle: "goodRow",
    textStyle: "goodText",
  };
}

function buildSpreadsheetReport(reportData) {
  const {
    createdAt,
    age,
    sleep,
    activity,
    fatigue,
    norms,
    weeklyForecast,
    badgeText,
    description,
    notification,
    conclusion,
  } = reportData;

  const avgSleep = weeklyForecast.reduce((sum, item) => sum + item.sleep, 0) / weeklyForecast.length;
  const avgActivity = weeklyForecast.reduce((sum, item) => sum + item.activity, 0) / weeklyForecast.length;
  const avgFatigue = weeklyForecast.reduce((sum, item) => sum + item.fatigue, 0) / weeklyForecast.length;
  const firstDay = weeklyForecast[0];
  const lastDay = weeklyForecast[weeklyForecast.length - 1];
  const trendFact = lastDay.fatigue > firstDay.fatigue ? "Нарастающая усталость 7 дней" : "Стабильный";
  const reportDate = formatDateTime(createdAt);
  const recommendationRows = buildRecommendationRows(reportData);

  const sheetOneRows = [
    xmlRow([xmlCell("Мониторинг здоровья студента во время сессии", "titleBlue", { mergeAcross: 9 })], { height: 34 }),
    xmlRow([xmlCell(`Период: прогноз на 7 дней | Отчёт сформирован: ${reportDate} | Студент: ${formatAge(age)}`, "subtitleBlue", { mergeAcross: 9 })], { height: 24 }),
    xmlRow([xmlCell("", "Default", { mergeAcross: 9 })], { height: 10 }),
    xmlRow([
      xmlCell("День", "headerBlue"),
      xmlCell("Дата", "headerBlue"),
      xmlCell("Сон (часов)", "headerBlue"),
      xmlCell("Норма сна", "headerBlue"),
      xmlCell("Активность (шаги)", "headerBlue"),
      xmlCell("Уровень стресса", "headerBlue"),
      xmlCell("Усталость (1-10)", "headerBlue"),
      xmlCell("Статус восстановления", "headerBlue"),
      xmlCell("AI-оценка риска", "headerBlue"),
      xmlCell("Рекомендация", "headerBlue"),
    ], { height: 30 }),
  ];

  weeklyForecast.forEach((dayData, index) => {
    const risk = getRiskInfo(dayData, norms);

    sheetOneRows.push(
      xmlRow([
        xmlCell(`День ${index + 1}`, risk.rowStyle),
        xmlCell(dayData.day, risk.rowStyle),
        xmlCell(formatDecimal(dayData.sleep), risk.rowStyle),
        xmlCell(`${norms.minSleep}-${norms.maxSleep} ч`, risk.rowStyle),
        xmlCell(formatWhole(dayData.activity), risk.rowStyle),
        xmlCell(getStressLabel(dayData.fatigue), risk.rowStyle),
        xmlCell(String(dayData.fatigue), risk.rowStyle),
        xmlCell(getRecoveryLabel(dayData.fatigue), risk.rowStyle),
        xmlCell(risk.label, risk.textStyle),
        xmlCell(getDailyRecommendation(dayData, norms), risk.rowStyle),
      ], { height: 26 }),
    );
  });

  sheetOneRows.push(
    xmlRow([
      xmlCell("СРЕДНЕЕ / ИТОГ", "summaryLabel", { mergeAcross: 1 }),
      xmlCell(formatDecimal(avgSleep), "summaryValue"),
      xmlCell(`${norms.minSleep}-${norms.maxSleep} ч`, "summaryValue"),
      xmlCell(formatWhole(avgActivity), "summaryValue"),
      xmlCell("—", "summaryValue"),
      xmlCell(formatDecimal(avgFatigue), "summaryValue"),
      xmlCell(getRecoveryLabel(lastDay.fatigue), "summaryValue"),
      xmlCell(badgeText, "summaryValue"),
      xmlCell(conclusion, "summaryValue"),
    ], { height: 28 }),
  );

  sheetOneRows.push(
    xmlRow([xmlCell("СИСТЕМНОЕ УВЕДОМЛЕНИЕ ОТ AI-МОДУЛЯ", "alertTitle", { mergeAcross: 9 })], { height: 30 }),
    xmlRow([xmlCell(notification, "alertBody", { mergeAcross: 9 })], { height: 44 }),
  );

  const sheetTwoRows = [
    xmlRow([xmlCell("AI-Анализ показателей и логика выводов", "titleBlue", { mergeAcross: 4 })], { height: 34 }),
    xmlRow([xmlCell("", "Default", { mergeAcross: 4 })], { height: 10 }),
    xmlRow([
      xmlCell("Источник данных", "headerBlue"),
      xmlCell("Показатель", "headerBlue"),
      xmlCell("Норма", "headerBlue"),
      xmlCell("Факт", "headerBlue"),
      xmlCell("Вывод AI", "headerBlue"),
    ], { height: 30 }),
    xmlRow([
      xmlCell("Фитнес-браслет", "dangerRow"),
      xmlCell("Продолжительность сна", "dangerRow"),
      xmlCell(`>= ${norms.minSleep} часов`, "goodRow"),
      xmlCell(`${formatDecimal(lastDay.sleep)} часов`, "dangerRow"),
      xmlCell(`Дефицит сна: ${formatDecimal(Math.max(0, norms.minSleep - lastDay.sleep))} ч. ${lastDay.sleep < norms.minSleep ? "Нужно восстановление." : "Сон в допустимом диапазоне."}`, "dangerRow"),
    ], { height: 38 }),
    xmlRow([
      xmlCell("Мобильное приложение", "dangerRow"),
      xmlCell("Дневная активность", "dangerRow"),
      xmlCell(`>= ${formatWhole(norms.activity)} шагов`, "goodRow"),
      xmlCell(`${formatWhole(lastDay.activity)} шагов`, "dangerRow"),
      xmlCell(`Отклонение по шагам: ${formatWhole(Math.max(0, norms.activity - lastDay.activity))}. ${lastDay.activity < norms.activity ? "Подвижность снижена." : "Активность достаточная."}`, "dangerRow"),
    ], { height: 38 }),
    xmlRow([
      xmlCell("Ручной ввод пользователя", "dangerRow"),
      xmlCell("Уровень усталости", "dangerRow"),
      xmlCell(norms.fatigue, "goodRow"),
      xmlCell(`${fatigue} из 10`, "dangerRow"),
      xmlCell(`Расчётная усталость: ${fatigue}/10. ${fatigue >= 7 ? "Самочувствие уже напряжённое." : "Есть пространство для профилактики."}`, "dangerRow"),
    ], { height: 38 }),
    xmlRow([
      xmlCell("AI (сравнение с профилем)", "dangerRow"),
      xmlCell("Тренд восстановления", "dangerRow"),
      xmlCell("Стабильный", "goodRow"),
      xmlCell(trendFact, "dangerRow"),
      xmlCell(lastDay.fatigue > firstDay.fatigue ? "Прогноз показывает нарастание утомления к концу недели." : "Резкого ухудшения не ожидается.", "dangerRow"),
    ], { height: 42 }),
    xmlRow([
      xmlCell("AI (агрегированный вывод)", "summaryAccent"),
      xmlCell("Общий риск переутомления", "summaryAccent"),
      xmlCell("Низкий / средний / высокий", "goodRow"),
      xmlCell(badgeText, "summaryValue"),
      xmlCell(description, "summaryAccent"),
    ], { height: 42 }),
  ];

  const sheetThreeRows = [
    xmlRow([xmlCell("Рекомендации AI-модуля для студента", "titleGreen", { mergeAcross: 3 })], { height: 34 }),
    xmlRow([xmlCell("", "Default", { mergeAcross: 3 })], { height: 10 }),
    xmlRow([
      xmlCell("Приоритет", "headerGreen"),
      xmlCell("Категория", "headerGreen"),
      xmlCell("Рекомендация", "headerGreen"),
      xmlCell("Ожидаемый эффект", "headerGreen"),
    ], { height: 30 }),
  ];

  recommendationRows.forEach((item) => {
    const styles = getPriorityStyles(item.priority);

    sheetThreeRows.push(
      xmlRow([
        xmlCell(item.priority, styles.textStyle),
        xmlCell(item.category, styles.rowStyle),
        xmlCell(item.recommendation, styles.rowStyle),
        xmlCell(item.effect, styles.rowStyle),
      ], { height: 42 }),
    );
  });

  return `<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>Digital Health Monitor</Author>
  <Created>${escapeXml(new Date(createdAt).toISOString())}</Created>
 </DocumentProperties>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#B9C3D0"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#B9C3D0"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#B9C3D0"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#B9C3D0"/>
   </Borders>
   <Font ss:FontName="Arial" ss:Size="12" ss:Color="#111111"/>
   <Interior ss:Color="#FFFFFF" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="titleBlue">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Bold="1" ss:Size="18" ss:Color="#FFFFFF"/>
   <Interior ss:Color="#3973B4" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="subtitleBlue">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Italic="1" ss:Size="12" ss:Color="#334155"/>
   <Interior ss:Color="#DCE8F7" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="titleGreen">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Bold="1" ss:Size="18" ss:Color="#FFFFFF"/>
   <Interior ss:Color="#1F7A3F" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="headerBlue">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="Arial" ss:Bold="1" ss:Size="12" ss:Color="#FFFFFF"/>
   <Interior ss:Color="#223F73" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="headerGreen">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="Arial" ss:Bold="1" ss:Size="12" ss:Color="#FFFFFF"/>
   <Interior ss:Color="#1F7A3F" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="goodRow">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="Arial" ss:Size="12" ss:Color="#111111"/>
   <Interior ss:Color="#D9EFD9" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="warnRow">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="Arial" ss:Size="12" ss:Color="#111111"/>
   <Interior ss:Color="#FFF0C4" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="dangerRow">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="Arial" ss:Size="12" ss:Color="#111111"/>
   <Interior ss:Color="#F7D1D1" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="goodText">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Bold="1" ss:Size="12" ss:Color="#047857"/>
   <Interior ss:Color="#D9EFD9" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="warnText">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Bold="1" ss:Size="12" ss:Color="#B45309"/>
   <Interior ss:Color="#FFF0C4" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="dangerText">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Bold="1" ss:Size="12" ss:Color="#B91C1C"/>
   <Interior ss:Color="#F7D1D1" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="summaryLabel">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Bold="1" ss:Size="12" ss:Color="#FFFFFF"/>
   <Interior ss:Color="#223F73" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="summaryValue">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="Arial" ss:Bold="1" ss:Size="12" ss:Color="#111111"/>
   <Interior ss:Color="#D9E6F2" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="summaryAccent">
   <Alignment ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="Arial" ss:Bold="1" ss:Size="12" ss:Color="#111111"/>
   <Interior ss:Color="#FFD84D" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="alertTitle">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Bold="1" ss:Size="14" ss:Color="#8B0000"/>
   <Interior ss:Color="#F7C9C9" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="alertBody">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="Arial" ss:Italic="1" ss:Size="12" ss:Color="#111111"/>
   <Interior ss:Color="#FFF4BF" ss:Pattern="Solid"/>
  </Style>
 </Styles>
 <Worksheet ss:Name="Дневной мониторинг">
  <Table>
   <Column ss:Width="96"/>
   <Column ss:Width="84"/>
   <Column ss:Width="126"/>
   <Column ss:Width="132"/>
   <Column ss:Width="160"/>
   <Column ss:Width="150"/>
   <Column ss:Width="130"/>
   <Column ss:Width="178"/>
   <Column ss:Width="154"/>
   <Column ss:Width="320"/>
   ${sheetOneRows.join("")}
  </Table>
 </Worksheet>
 <Worksheet ss:Name="AI-анализ">
  <Table>
   <Column ss:Width="220"/>
   <Column ss:Width="235"/>
   <Column ss:Width="155"/>
   <Column ss:Width="190"/>
   <Column ss:Width="430"/>
   ${sheetTwoRows.join("")}
  </Table>
 </Worksheet>
 <Worksheet ss:Name="Рекомендации">
  <Table>
   <Column ss:Width="150"/>
   <Column ss:Width="185"/>
   <Column ss:Width="520"/>
   <Column ss:Width="380"/>
   ${sheetThreeRows.join("")}
  </Table>
 </Worksheet>
</Workbook>`;
}

function downloadSpreadsheetReport() {
  if (!latestReport) {
    return;
  }

  const xml = buildSpreadsheetReport(latestReport);
  const blob = new Blob([xml], { type: "application/vnd.ms-excel;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = `digital-health-report-${latestReport.createdAt}.xls`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function setExportVisibility(isVisible) {
  exportPanel.classList.toggle("export-hidden", !isVisible);
}

function analyzeProfile() {
  const age = Number(ageInput.value);
  const sleep = Number(sleepInput.value);
  const activity = Number(activityInput.value);

  if (!validateInput(age, sleep, activity)) {
    latestReport = null;
    setExportVisibility(false);
    setListItems(findingsList, ["Введите корректные значения возраста, сна и дневной активности."]);
    setListItems(recommendationsList, ["После заполнения формы снова нажмите «Запустить AI-анализ»."]);
    return;
  }

  const norms = getNormsByAge(age);
  const fatigue = estimateFatigue(sleep, activity, norms);
  const weeklyForecast = buildWeeklyForecast(age, sleep, activity);

  const sleepDeficit = Math.max(0, norms.minSleep - sleep);
  const activityDeficit = Math.max(0, norms.activity - activity);
  const sleepPercent = Math.round((sleepDeficit / norms.minSleep) * 100);
  const activityPercent = Math.round((activityDeficit / norms.activity) * 100);

  updateReferenceCards(age, norms);
  updateCurrentCard(age, sleep, activity, fatigue);
  renderTrend(weeklyForecast);

  trendDescription.textContent =
    `Смоделированная неделя для пользователя ${age} лет. Если такой режим сохранится, к концу недели восстановление может ухудшаться.`;

  const findings = [];
  const recommendations = [];

  if (sleepDeficit > 0) {
    findings.push(
      `Для возраста ${age} лет рекомендовано не меньше ${norms.minSleep} часов сна. Сейчас дефицит составляет ${sleepDeficit.toFixed(1)} ч (${sleepPercent}%).`,
    );
    recommendations.push(
      `Поднимите сон хотя бы до ${norms.minSleep}-${norms.maxSleep} часов: сегодня лучше лечь раньше и сократить вечернюю нагрузку.`,
    );
  } else {
    recommendations.push("По сну отклонение не видно. Сохраняйте стабильное время отхода ко сну.");
  }

  if (activityDeficit > 0) {
    findings.push(
      `Активность ниже возрастного ориентира на ${Math.round(activityDeficit).toLocaleString("ru-RU")} шагов (${activityPercent}% ниже нормы).`,
    );
    recommendations.push(
      `Добавьте спокойную прогулку или ходьбу днём, чтобы приблизиться к ${formatSteps(norms.activity)}.`,
    );
  } else {
    recommendations.push("Дневная активность достаточная. Поддерживайте этот уровень движения.");
  }

  if (fatigue >= 7) {
    findings.push("Расчётная усталость высокая: сочетание недосыпа и низкой подвижности ухудшает восстановление.");
    recommendations.push("Сделайте разгрузочный вечер без длительной учёбы и без экрана за 1 час до сна.");
  } else if (fatigue >= 5) {
    findings.push("Есть умеренная усталость. Режим уже не оптимальный и может ухудшиться при продолжении такой нагрузки.");
    recommendations.push("Сделайте короткие перерывы днём и не откладывайте сон на позднюю ночь.");
  }

  const lastDayFatigue = weeklyForecast[weeklyForecast.length - 1].fatigue;
  const firstDayFatigue = weeklyForecast[0].fatigue;

  if (lastDayFatigue > firstDayFatigue) {
    findings.push("Прогноз недели показывает нарастающую усталость, если сохранить такой же режим каждый день.");
  } else {
    findings.push("Прогноз недели не показывает резкого ухудшения, если режим останется таким же.");
  }

  let badgeText = "Низкий риск";
  let badgeClass = "status-badge good";
  let title = "Показатели близки к возрастной норме";
  let description = `Для возраста ${age} лет серьёзных отклонений не видно. AI-модуль советует просто поддерживать текущий режим сна и активности.`;
  let notification =
    "«Серьёзных признаков перегрузки не выявлено. Поддерживайте регулярный сон, движение и умеренную вечернюю нагрузку»";
  let conclusion = "Состояние выглядит стабильным";

  const highRisk = sleepDeficit >= 2 || activityPercent >= 35 || fatigue >= 7;
  const mediumRisk = sleepDeficit >= 1 || activityPercent >= 15 || fatigue >= 5;

  if (highRisk) {
    badgeText = "Высокий риск";
    badgeClass = "status-badge warning";
    title = "Есть признаки ухудшения самочувствия";
    description = `Для возраста ${age} лет AI-модуль видит выраженный недосып, снижение активности и плохое восстановление.`;
    notification =
      "«Ваши текущие показатели выглядят неблагоприятно. Есть признаки перегрузки: сна мало, активность снижена. Постарайтесь уменьшить нагрузку вечером и восстановить режим сна»";
    conclusion = "Нужна разгрузка и восстановление";
  } else if (mediumRisk) {
    badgeText = "Средний риск";
    badgeClass = "status-badge neutral";
    title = "Намечается неблагоприятная тенденция";
    description = `Для возраста ${age} лет часть показателей уже вышла из комфортного диапазона. Лучше скорректировать режим сейчас, пока ухудшение не стало выраженным.`;
    notification =
      "«Показатели начали отклоняться от возрастной нормы. Постарайтесь добавить сон и движение, чтобы не допустить переутомления»";
    conclusion = "Есть слабые места в режиме";
  }

  if (!findings.length) {
    findings.push("По введённым данным выраженных минусов не обнаружено.");
  }

  while (recommendations.length < 3) {
    recommendations.push("Пейте воду в течение дня и делайте короткие паузы во время нагрузки.");
  }

  statusBadge.className = badgeClass;
  statusBadge.textContent = badgeText;
  resultTitle.textContent = title;
  resultText.textContent = description;
  notificationText.textContent = notification;
  personalConclusion.textContent = conclusion;

  setListItems(findingsList, findings);
  setListItems(recommendationsList, recommendations);

  latestReport = {
    createdAt: Date.now(),
    badgeText,
    age,
    sleep,
    activity,
    fatigue,
    norms,
    weeklyForecast,
    findings,
    recommendations,
    description,
    notification,
    conclusion,
  };

  setExportVisibility(true);

  storeAnalysis({
    createdAt: latestReport.createdAt,
    badgeText,
    age,
    sleep,
    activity,
    fatigue,
    conclusion,
  });
}

function resetScenario() {
  ageInput.value = String(defaultState.age);
  sleepInput.value = String(defaultState.sleep);
  activityInput.value = String(defaultState.activity);

  const norms = getNormsByAge(defaultState.age);
  const fatigue = estimateFatigue(defaultState.sleep, defaultState.activity, norms);

  updateReferenceCards(defaultState.age, norms);
  updateCurrentCard(defaultState.age, defaultState.sleep, defaultState.activity, fatigue);
  renderTrend(buildWeeklyForecast(defaultState.age, defaultState.sleep, defaultState.activity));

  trendDescription.textContent =
    "Прогноз строится по введённым пользователем значениям и показывает, как может меняться восстановление в течение недели.";

  statusBadge.className = "status-badge neutral";
  statusBadge.textContent = "Ожидание анализа";
  resultTitle.textContent = "Готов к проверке недельных данных";
  resultText.textContent =
    "Введите данные сверху и нажмите кнопку в верхнем блоке, чтобы получить персональный AI-анализ по возрасту, сну и активности.";
  notificationText.textContent = defaultNotification;
  personalConclusion.textContent = "Ожидает ваши данные";
  latestReport = null;
  setExportVisibility(false);

  setListItems(findingsList, ["После запуска AI здесь появятся минусы именно по вашим введённым данным."]);
  setListItems(recommendationsList, ["После запуска AI здесь появятся персональные рекомендации по режиму."]);

  renderHistory();
}

analyzeButton.addEventListener("click", analyzeProfile);
resetButton.addEventListener("click", resetScenario);
downloadReportButton.addEventListener("click", downloadSpreadsheetReport);
clearHistoryButton.addEventListener("click", () => {
  saveHistory([]);
  renderHistory();
});

resetScenario();
