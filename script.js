const dayLabels = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"];

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

function analyzeProfile() {
  const age = Number(ageInput.value);
  const sleep = Number(sleepInput.value);
  const activity = Number(activityInput.value);

  if (!validateInput(age, sleep, activity)) {
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

  setListItems(findingsList, ["После запуска AI здесь появятся минусы именно по вашим введённым данным."]);
  setListItems(recommendationsList, ["После запуска AI здесь появятся персональные рекомендации по режиму."]);
}

analyzeButton.addEventListener("click", analyzeProfile);
resetButton.addEventListener("click", resetScenario);

resetScenario();
