#!/usr/bin/env node
import fs from "node:fs";
import path from "node:path";
import { createHash } from "node:crypto";

const FORM_URL =
  process.env.FORM_URL ?? "https://forms.yandex.ru/u/69a805c3902902c6015a5aac";
const TARGET_SUCCESSFUL_SUBMITS = Number(
  process.env.TARGET_SUCCESSFUL_SUBMITS ?? "3",
);
const BRANCH_POLICY = process.env.BRANCH_POLICY ?? "random_yes_no";
const CAPTCHA_POLICY = process.env.CAPTCHA_POLICY ?? "restart_browser";
const MAX_TOTAL_ATTEMPTS = Number(process.env.MAX_TOTAL_ATTEMPTS ?? "20");
const HEADLESS = (process.env.HEADLESS ?? "1") !== "0";
const OUT_PATH =
  process.env.OUT_PATH ??
  path.resolve(process.cwd(), "yandex-form-runs-report.json");
const BROWSER_NAME = (process.env.BROWSER ?? "chromium").toLowerCase();
const BROWSER_CHANNEL = process.env.BROWSER_CHANNEL ?? "chrome";
const BROWSER_EXECUTABLE_PATH = process.env.BROWSER_EXECUTABLE_PATH ?? "";
const INCOGNITO = (process.env.INCOGNITO ?? "1") !== "0";
const RUN_MODE = (process.env.RUN_MODE ?? "target_success").toLowerCase();
const STRICT_RULES_ONLY =
  (process.env.STRICT_RULES_ONLY ?? (RUN_MODE === "quota" ? "1" : "0")) !== "0";
const QUOTA_STATE_PATH =
  process.env.QUOTA_STATE_PATH ??
  path.resolve(process.cwd(), "quota-state.json");
const QUOTA_TARGETS_PATH = process.env.QUOTA_TARGETS_PATH ?? "";
const RESET_QUOTA_STATE = (process.env.RESET_QUOTA_STATE ?? "0") === "1";
const LIKERT_TARGETS_PATH = process.env.LIKERT_TARGETS_PATH ?? "";
const PAGE_VALIDATION_RETRIES = Number(process.env.PAGE_VALIDATION_RETRIES ?? "3");
const START_PAGE_RECOVERY_ATTEMPTS = Number(
  process.env.START_PAGE_RECOVERY_ATTEMPTS ?? "3",
);
const START_PAGE_RECOVERY_MODE = (
  process.env.START_PAGE_RECOVERY_MODE ?? "browser_restart"
).toLowerCase();
const FORM_DOM_MODE = (process.env.FORM_DOM_MODE ?? "auto").toLowerCase();
const VALIDATION_VISIBLE_ONLY = (process.env.VALIDATION_VISIBLE_ONLY ?? "1") !== "0";
const RADIO_FALLBACK_STRATEGY = (
  process.env.RADIO_FALLBACK_STRATEGY ?? "aggressive"
).toLowerCase();
const LOG_PAGE_DEBUG = (process.env.LOG_PAGE_DEBUG ?? "1") !== "0";
const SCHEMA_DISCOVERY_MODE = (
  process.env.SCHEMA_DISCOVERY_MODE ?? "startup_localstorage"
).toLowerCase();
const SCHEMA_STORAGE_KEY = process.env.SCHEMA_STORAGE_KEY ?? "auto";
const SELECTOR_STRATEGY = (process.env.SELECTOR_STRATEGY ?? "text").toLowerCase();
const TABLE_SCALE_NO_OPINION_POLICY = (
  process.env.TABLE_SCALE_NO_OPINION_POLICY ?? "never"
).toLowerCase();
const LIKERT_NO_TARGET_SCORE = Math.min(
  5,
  Math.max(1, Math.round(Number(process.env.LIKERT_NO_TARGET_SCORE ?? "4"))),
);
const normalizeLikertRandomness = (value, fallback = 0.35) => {
  const n = Number(value);
  if (!Number.isFinite(n)) return fallback;
  return Math.min(1, Math.max(0, n));
};
const LIKERT_TARGET_RANDOMNESS = normalizeLikertRandomness(
  process.env.LIKERT_TARGET_RANDOMNESS,
  0.35,
);
const LIKERT_FALLBACK_RANDOMNESS = normalizeLikertRandomness(
  process.env.LIKERT_FALLBACK_RANDOMNESS,
  0.55,
);
const SUCCESS_MODE = (process.env.SUCCESS_MODE ?? "click_and_confirm").toLowerCase();
const PROGRESS_MODE = (process.env.PROGRESS_MODE ?? "per_attempt").toLowerCase();
const PROGRESS_EVERY_N_RAW = Number(process.env.PROGRESS_EVERY_N ?? "1");
const PROGRESS_EVERY_N = Number.isFinite(PROGRESS_EVERY_N_RAW)
  ? Math.max(1, Math.floor(PROGRESS_EVERY_N_RAW))
  : 1;
const RUNNER_VERSION = "2026-03-05-quota-precision-v1";
// `rule_based` is counted only when an explicit FIXED_RULES match drives selection.
const FIXED_RULES = [
  {
    id: "organizer-yes",
    question_includes: [
      "были ли вы организатором своего последнего путешествия",
      "бронирование отелей и билетов",
    ],
    option_equals: "Да",
  },
  {
    id: "tickets-online-aggregator",
    question_includes: [
      "каким основным способом",
      "покупали билеты",
      "для этого путешествия",
    ],
    option_includes: "через онлайн-агрегатор/сервис",
  },
  {
    id: "stay-online-aggregator",
    question_includes: [
      "каким основным способом",
      "бронировали",
    ],
    option_includes: "через онлайн-агрегатор/сервис",
  },
];
const AGE_RULES = {
  vacation: {
    with_children: [25, 45],
    without_children: [25, 40],
  },
  weekend: {
    with_children: [28, 40],
    without_children: [18, 35],
  },
};
const QUOTA_CITY_DEFS = [
  { city_key: "nsk", city_label: "Новосибирск" },
  { city_key: "ufa", city_label: "Уфа" },
  { city_key: "prm", city_label: "Пермь" },
];
const QUOTA_MATRIX = [
  { trip_type: "vacation", children_mode: "with_children", yp: 39, other: 30 },
  { trip_type: "vacation", children_mode: "without_children", yp: 38, other: 17 },
  { trip_type: "weekend", children_mode: "with_children", yp: 38, other: 32 },
  { trip_type: "weekend", children_mode: "without_children", yp: 38, other: 32 },
];
const LIKERT_DEFAULT_RULES_9_13 = [
  { id: "9:1", question_includes: ["цены на билеты и отели были ниже"] },
  { id: "9:2", question_includes: ["известен в народе"] },
  { id: "9:3", question_includes: ["друзья и знакомые рекомендовали"] },
  { id: "9:4", question_includes: ["широкий выбор отелей"] },
  { id: "9:5", question_includes: ["широкий выбор билетов"] },
  { id: "9:6", question_includes: ["вся необходимая информация"] },
  { id: "10:1", question_includes: ["визуальную информацию с картами"] },
  { id: "10:2", question_includes: ["контакты службы поддержки или чата"] },
  { id: "10:3", question_includes: ["понятный и удобный интерфейс"] },
  { id: "10:4", question_includes: ["поддержки сервиса быстро и эффективно"] },
  { id: "10:5", question_includes: ["положительный опыт использования"] },
  { id: "10:6", question_includes: ["описание, фотографии и отзывы"] },
  { id: "11:1", question_includes: ["мошенничеством или фейковым предложением"] },
  { id: "11:2", question_includes: ["потерять деньги из-за ошибки"] },
  { id: "11:3", question_includes: ["бронь на сервисе может"] },
  { id: "11:4", question_includes: ["недоверие к работе службы поддержки"] },
  { id: "11:5", question_includes: ["чёткие гарантии и ответственность"] },
  { id: "11:6", question_includes: ["кто именно при таком бронировании"] },
  { id: "11:7", question_includes: ["итоговая цена окажется выше заявленной"] },
  { id: "12:1", question_includes: ["возврат через агрегатор будет дольше"] },
  { id: "12:2", question_includes: ["самостоятельно всё контролировать"] },
  { id: "12:3", question_includes: ["процесс выбора кажется слишком долгим"] },
  { id: "12:4", question_includes: ["выбрать «не тот» вариант"] },
  { id: "12:5", question_includes: ["недостаточно информации в сервисе"] },
  { id: "12:6", question_includes: ["слишком навязчивая реклама"] },
  { id: "12:7", question_includes: ["навязанные дополнительные услуги"] },
  { id: "13:1", question_includes: ["сложно или непонятно, как вернуть деньги"] },
  { id: "13:2", question_includes: ["предпочтение личного общения"] },
  { id: "13:3", question_includes: ["напрямую у отеля или авиакомпании безопаснее"] },
  { id: "13:4", question_includes: ["туроператор лучше разберётся"] },
  { id: "13:5", question_includes: ["туроператору проще доверить организацию"] },
  { id: "13:6", question_includes: ["онлайн-сервис решает проблемы за границей"] },
  { id: "13:7", question_includes: ["в другой стране будет сложно получить помощь"] },
];
const LIKERT_ORDER_ALL_IDS = LIKERT_DEFAULT_RULES_9_13.map((rule) => rule.id);
const LIKERT_ORDER_SKIP_10_IDS = LIKERT_ORDER_ALL_IDS.filter(
  (id) => !id.startsWith("10:"),
);
const QUOTA_CITY_LABEL_BY_KEY = Object.fromEntries(
  QUOTA_CITY_DEFS.map((city) => [city.city_key, city.city_label]),
);

const randInt = (min, max) =>
  Math.floor(Math.random() * (max - min + 1)) + min;
const clampLikertScore = (value) => Math.min(5, Math.max(1, Math.round(Number(value))));

const randomSleep = (minMs, maxMs) =>
  new Promise((resolve) => setTimeout(resolve, randInt(minMs, maxMs)));

const pageNoFromUrl = (url) => {
  const match = /[?&]page=(\d+)/.exec(url);
  return match ? match[1] : "1";
};

const normalizeText = (value) =>
  String(value ?? "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();

const normalizeRuleText = (value) =>
  normalizeText(
    String(value ?? "")
      .replace(/\{[^}]*\}/g, " ")
      .replace(/[*_`~]/g, " ")
      .replace(/[()[\]{}<>]/g, " ")
      .replace(/[,:;!?]/g, " "),
  );

const extractPollId = (url) => {
  const match = /\/(?:poll|u)\/([^/?#]+)/i.exec(String(url ?? ""));
  return match?.[1] ?? null;
};

const resolveSchemaStorageKey = (pollId, explicitKey) => {
  if (explicitKey && explicitKey !== "auto") return explicitKey;
  if (!pollId) return null;
  return `pythia10-${pollId}`;
};
const QUOTA_CITY_NORMALIZED_SET = new Set(
  QUOTA_CITY_DEFS.map((city) => normalizeText(city.city_label)),
);

const inferTripType = (label) => {
  const n = normalizeText(label);
  if (n.includes("отпуск") || n.includes("длинная поездка")) return "vacation";
  if (n.includes("выходные") || n.includes("короткая")) return "weekend";
  return null;
};

const inferChildrenMode = (label) => {
  const n = normalizeText(label);
  if (n.includes("с детьми")) return "with_children";
  if (n.includes("без детей")) return "without_children";
  return null;
};

const inferGenderMode = (label) => {
  const n = normalizeText(label);
  if (n.includes("муж")) return "male";
  if (n.includes("жен")) return "female";
  return null;
};

const resolveAgeRange = (runState) => {
  const tripType = runState.trip_type;
  const childrenMode = runState.children_mode;
  if (!tripType || !childrenMode) return null;
  const range = AGE_RULES[tripType]?.[childrenMode];
  if (!range) return null;
  return { min: range[0], max: range[1] };
};

const resolveBranchChoice = (runState) => {
  if (BRANCH_POLICY === "always_yes") return "Да";
  if (BRANCH_POLICY === "always_no") return "Нет";
  return runState.branch_choice;
};

const isOrganizerQuestion = (questionText) => {
  const n = normalizeText(questionText);
  return (
    n.includes("были ли вы организатором своего последнего путешествия") &&
    n.includes("бронирование отелей и билетов")
  );
};

const isServiceMethodQuestion = (questionText) => {
  const n = normalizeText(questionText);
  return (
    n.includes("каким основным способом") &&
    (n.includes("покупали билеты") ||
      n.includes("бронировали жиль") ||
      n.includes("бронировали отел"))
  );
};

const isIncomeQuestion = (questionText) => {
  const n = normalizeText(questionText);
  return (
    n.includes("среднемесячный доход на одного члена семьи") ||
    (n.includes("среднемесячный доход") && n.includes("члена семьи")) ||
    (n.includes("ваш доход") && n.includes("члена семьи"))
  );
};

const isYandexTravelLabel = (label) => {
  const n = normalizeText(label);
  return n.includes("яндекс путешеств") || n.includes("yandex travel");
};

const isOtherProviderAllowedLabel = (label) => {
  const n = normalizeText(label);
  if (!n) return false;
  if (n.includes("другое")) return false;
  return !isYandexTravelLabel(n);
};

const isProviderAggregatorQuestion = (questionText, labels) => {
  const normalizedQuestion = normalizeRuleText(questionText);
  const normalizedLabels = labels.map((label) => normalizeRuleText(label));
  const hasYandex = normalizedLabels.some((label) => isYandexTravelLabel(label));
  const hasOtherProvider = normalizedLabels.some((label) =>
    isOtherProviderAllowedLabel(label),
  );
  if (!hasYandex || !hasOtherProvider) return false;

  // Method-level options (aggregator/direct/operator) are not a concrete provider list.
  const hasMethodLevelOptions = normalizedLabels.some(
    (label) =>
      label.includes("через онлайн агрегатор сервис") ||
      label.includes("напрямую") ||
      label.includes("туроператор") ||
      label.includes("турфирм"),
  );
  if (hasMethodLevelOptions) return false;

  const providerKeywordRegex =
    /(onetwotrip|aviasales|tutu|ostrovok|ozon\s*travel|т-?путешеств|яндекс путешеств)/i;
  const hasProviderKeywords = normalizedLabels.some((label) =>
    providerKeywordRegex.test(label),
  );
  if (!hasProviderKeywords) return false;

  const questionLooksProvider =
    normalizedQuestion.includes("агрегатор") ||
    normalizedQuestion.includes("бронирова") ||
    normalizedQuestion.includes("покупали билеты") ||
    normalizedQuestion.includes("жиль") ||
    normalizedQuestion.includes("отел");
  return questionLooksProvider;
};

const isConcreteProviderListQuestion = (questionText, labels) =>
  isProviderAggregatorQuestion(questionText, labels);

function findQuotaDrivenIndex({ questionText, labels, runState }) {
  if (runState.run_mode !== "quota" || !runState.target_cell) return null;
  const quotaCell = runState.target_cell;
  const normalizedLabels = labels.map((label) => normalizeText(label));
  const maleIndex = normalizedLabels.findIndex((label) => inferGenderMode(label) === "male");
  const femaleIndex = normalizedLabels.findIndex(
    (label) => inferGenderMode(label) === "female",
  );
  if (maleIndex !== -1 && femaleIndex !== -1) {
    return pickBalancedGender({
      maleIndex,
      femaleIndex,
      runState,
    });
  }

  const cityOptionCount = normalizedLabels.filter((label) =>
    QUOTA_CITY_NORMALIZED_SET.has(label),
  ).length;
  if (cityOptionCount >= 2) {
    const targetCity = normalizeText(
      quotaCell.city_label ?? QUOTA_CITY_LABEL_BY_KEY[quotaCell.city_key] ?? "",
    );
    const cityIndex = normalizedLabels.findIndex((label) => label === targetCity);
    if (cityIndex !== -1) {
      return { index: cityIndex, rule_id: "quota_city" };
    }
  }

  if (isOrganizerQuestion(questionText)) {
    const yesIndex = normalizedLabels.findIndex((label) => label === "да");
    if (yesIndex !== -1) {
      return { index: yesIndex, rule_id: "quota_organizer_yes" };
    }
  }

  const tripCandidates = labels.map((label) => inferTripType(label));
  if (tripCandidates.some(Boolean)) {
    const tripIndex = tripCandidates.findIndex(
      (tripType) => tripType === quotaCell.trip_type,
    );
    if (tripIndex !== -1) {
      return { index: tripIndex, rule_id: "quota_trip_type" };
    }
  }

  const childrenCandidates = labels.map((label) => inferChildrenMode(label));
  if (childrenCandidates.some(Boolean)) {
    const childrenIndex = childrenCandidates.findIndex(
      (mode) => mode === quotaCell.children_mode,
    );
    if (childrenIndex !== -1) {
      return { index: childrenIndex, rule_id: "quota_children_mode" };
    }
  }

  if (isServiceMethodQuestion(questionText)) {
    const aggregatorIndex = normalizedLabels.findIndex((label) =>
      label.includes("онлайн-агрегатор/сервис"),
    );
    if (aggregatorIndex !== -1) {
      return { index: aggregatorIndex, rule_id: "quota_service_method_aggregator" };
    }
  }

  const yandexIndex = normalizedLabels.findIndex((label) =>
    isYandexTravelLabel(label),
  );
  const otherIndex = normalizedLabels.findIndex((label) =>
    isOtherProviderAllowedLabel(label),
  );
  if (yandexIndex !== -1 && otherIndex !== -1) {
    if (quotaCell.provider_bucket === "yp") {
      return { index: yandexIndex, rule_id: "quota_provider_yp" };
    }
    return { index: otherIndex, rule_id: "quota_provider_other" };
  }

  return null;
}

function findLikertDrivenDecision({ questionText, labels, runState }) {
  if (!runState?.likert_rules?.length) return null;

  const rule = findLikertTargetRule(questionText, runState.likert_rules);
  if (!rule) return null;
  const segmentKey = runState.likert_segment_key ?? "__global__";
  const progressKey = `${segmentKey}|${rule.id}`;
  const targetMeanFromSegment =
    runState?.likert_segment_targets?.[segmentKey]?.[rule.id] ??
    runState?.likert_segment_targets?.__global__?.[rule.id] ??
    null;
  const targetMean =
    toFiniteNumberOrNull(targetMeanFromSegment) ??
    toFiniteNumberOrNull(rule.target_mean);
  if (!Number.isFinite(targetMean)) return null;
  const expectedTotal =
    runState?.likert_segment_expected_total?.[segmentKey] ??
    runState?.likert_expected_total ??
    null;

  const scoreToIndex = buildLikertScoreToIndexMap(labels);
  if (!scoreToIndex.size) return null;

  const alreadyChosen = runState.likert_answers_map?.get(progressKey);
  if (alreadyChosen) {
    const existingIndex = scoreToIndex.get(clampLikertScore(alreadyChosen.value));
    if (existingIndex !== undefined) {
      return {
        index: existingIndex,
        rule_id: rule.id,
        progress_key: progressKey,
        segment_key: segmentKey,
        target_mean: targetMean,
        score: clampLikertScore(alreadyChosen.value),
      };
    }
  }

  const progressItem = runState.likert_progress?.[progressKey] ?? {
    count: 0,
    sum: 0,
  };
  const desiredScore = pickLikertScoreByTarget({
    targetMean,
    currentCount: progressItem.count,
    currentSum: progressItem.sum,
    expectedTotal,
  });

  let selectedScore = desiredScore;
  if (!scoreToIndex.has(selectedScore)) {
    for (const candidate of [5, 4, 3, 2, 1]) {
      if (scoreToIndex.has(candidate)) {
        selectedScore = candidate;
        break;
      }
    }
  }
  const selectedIndex = scoreToIndex.get(selectedScore);
  if (selectedIndex === undefined) return null;

  return {
    index: selectedIndex,
    rule_id: rule.id,
    progress_key: progressKey,
    segment_key: segmentKey,
    target_mean: targetMean,
    score: selectedScore,
  };
}

function findRuleDrivenIndex({ pageNo, questionText, labels }) {
  const normalizedQuestion = normalizeRuleText(questionText);
  const normalizedLabels = labels.map((label) => normalizeRuleText(label));

  for (const rule of FIXED_RULES) {
    if (rule.page && String(rule.page) !== String(pageNo)) continue;

    if (Array.isArray(rule.question_includes) && rule.question_includes.length) {
      const allFound = rule.question_includes.every((part) =>
        normalizedQuestion.includes(normalizeRuleText(part)),
      );
      if (!allFound) continue;
    }

    if (rule.option_equals) {
      const idx = normalizedLabels.findIndex(
        (value) => value === normalizeRuleText(rule.option_equals),
      );
      if (idx !== -1) {
        return { index: idx, rule_id: rule.id ?? "unnamed_rule" };
      }
    }

    if (rule.option_includes) {
      const idx = normalizedLabels.findIndex((value) =>
        value.includes(normalizeRuleText(rule.option_includes)),
      );
      if (idx !== -1) {
        return { index: idx, rule_id: rule.id ?? "unnamed_rule" };
      }
    }
  }

  return null;
}

const classifyScheme = (selectionStats) => {
  const randomCount = selectionStats.random ?? 0;
  const ruleBasedCount = selectionStats.rule_based ?? 0;
  if (randomCount > 0 && ruleBasedCount > 0) return "mixed";
  if (ruleBasedCount > 0) return "rule_based";
  return "random";
};

const buildAttemptStatusCounter = (attempts) => {
  const counter = {};
  for (const attempt of attempts ?? []) {
    const status = String(attempt?.status ?? "unknown");
    counter[status] = sanitizeTargetValue(counter[status], 0) + 1;
  }
  return counter;
};

const makeQuotaCellKey = ({
  city_key,
  trip_type,
  children_mode,
  provider_bucket,
}) => `${city_key}|${trip_type}|${children_mode}|${provider_bucket}`;

const parseQuotaCellKey = (key) => {
  const [city_key, trip_type, children_mode, provider_bucket] = String(key).split("|");
  if (!city_key || !trip_type || !children_mode || !provider_bucket) return null;
  return { city_key, trip_type, children_mode, provider_bucket };
};

const buildDefaultQuotaTargets = () => {
  const targets = {};
  for (const city of QUOTA_CITY_DEFS) {
    for (const row of QUOTA_MATRIX) {
      for (const bucket of ["yp", "other"]) {
        const key = makeQuotaCellKey({
          city_key: city.city_key,
          trip_type: row.trip_type,
          children_mode: row.children_mode,
          provider_bucket: bucket,
        });
        targets[key] = row[bucket];
      }
    }
  }
  return targets;
};

const sanitizeTargetValue = (value, fallbackValue = 0) => {
  const n = Number(value);
  if (!Number.isFinite(n)) return fallbackValue;
  if (n < 0) return 0;
  return Math.floor(n);
};

const normalizeAttemptStatusCounter = (raw) => {
  const out = {};
  if (!raw || typeof raw !== "object") return out;
  for (const [key, value] of Object.entries(raw)) {
    const k = String(key ?? "").trim();
    if (!k) continue;
    out[k] = sanitizeTargetValue(value, 0);
  }
  return out;
};

const incrementCounterKey = (counter, key, amount = 1) => {
  const normalizedKey = String(key ?? "").trim() || "unknown";
  const inc = sanitizeTargetValue(amount, 0);
  if (!counter || typeof counter !== "object") return;
  counter[normalizedKey] = sanitizeTargetValue(counter[normalizedKey], 0) + inc;
};

const stableStringifyObject = (value) => {
  if (Array.isArray(value)) {
    return `[${value.map((item) => stableStringifyObject(item)).join(",")}]`;
  }
  if (value && typeof value === "object") {
    const keys = Object.keys(value).sort();
    return `{${keys
      .map((key) => `${JSON.stringify(key)}:${stableStringifyObject(value[key])}`)
      .join(",")}}`;
  }
  return JSON.stringify(value);
};

const computeQuotaTargetsHash = (targetsMap) =>
  createHash("sha256").update(stableStringifyObject(targetsMap)).digest("hex");

const normalizeFormUrlForState = (url) => String(url ?? "").replace(/[/?#]+$/, "");

const isProcessAlive = (pid) => {
  if (!Number.isFinite(Number(pid)) || Number(pid) <= 0) return false;
  try {
    process.kill(Number(pid), 0);
    return true;
  } catch (error) {
    if (error?.code === "EPERM") return true;
    return false;
  }
};

const acquireStateLock = (lockPath, lockPayload) => {
  if (fs.existsSync(lockPath)) {
    try {
      const existingRaw = JSON.parse(fs.readFileSync(lockPath, "utf8"));
      const existingPid = Number(existingRaw?.pid);
      if (isProcessAlive(existingPid)) {
        return {
          acquired: false,
          reason: "active",
          existing: existingRaw,
        };
      }
    } catch {
      // stale or malformed lock; will overwrite.
    }
    try {
      fs.unlinkSync(lockPath);
    } catch {
      // best effort
    }
  }
  writeJsonAtomic(lockPath, lockPayload);
  return { acquired: true };
};

const releaseStateLock = (lockPath, ownerPid) => {
  if (!fs.existsSync(lockPath)) return;
  try {
    const raw = JSON.parse(fs.readFileSync(lockPath, "utf8"));
    const lockPid = Number(raw?.pid);
    if (Number.isFinite(lockPid) && lockPid > 0 && lockPid !== Number(ownerPid)) {
      return;
    }
  } catch {
    // proceed with unlink for malformed lock.
  }
  try {
    fs.unlinkSync(lockPath);
  } catch {
    // best effort
  }
};

const loadQuotaTargets = () => {
  const defaults = buildDefaultQuotaTargets();
  if (!QUOTA_TARGETS_PATH) return defaults;
  if (!fs.existsSync(QUOTA_TARGETS_PATH)) return defaults;

  const raw = JSON.parse(fs.readFileSync(QUOTA_TARGETS_PATH, "utf8"));
  const candidateMap = raw?.targets && typeof raw.targets === "object" ? raw.targets : raw;
  if (!candidateMap || typeof candidateMap !== "object") return defaults;

  const merged = { ...defaults };
  for (const [key, value] of Object.entries(candidateMap)) {
    if (!(key in merged)) continue;
    merged[key] = sanitizeTargetValue(value, merged[key]);
  }
  return merged;
};

const buildQuotaCells = (targetsMap) => {
  const cityLabelByKey = Object.fromEntries(
    QUOTA_CITY_DEFS.map((city) => [city.city_key, city.city_label]),
  );
  const keys = Object.keys(targetsMap).sort();
  const cells = [];
  for (const key of keys) {
    const parsed = parseQuotaCellKey(key);
    if (!parsed) continue;
    cells.push({
      key,
      ...parsed,
      city_label: cityLabelByKey[parsed.city_key] ?? parsed.city_key,
      target: sanitizeTargetValue(targetsMap[key], 0),
    });
  }
  return cells;
};

const normalizeCompletedMap = (targetsMap, completedRaw) => {
  const normalized = {};
  for (const [key, target] of Object.entries(targetsMap)) {
    normalized[key] = sanitizeTargetValue(completedRaw?.[key], 0);
    if (normalized[key] > sanitizeTargetValue(target, 0)) {
      normalized[key] = sanitizeTargetValue(target, 0);
    }
  }
  return normalized;
};

const loadQuotaState = (targetsMap) => {
  const makeInitialState = () => {
    const completed = normalizeCompletedMap(targetsMap, {});
    return {
      completed,
      attempts_total: 0,
      success_total: 0,
      submit_click_total: 0,
      submit_click_unconfirmed_total: 0,
      attempt_status_counter: {},
      last_attempt_id: 0,
      form_url: normalizeFormUrlForState(FORM_URL),
      quota_targets_hash: computeQuotaTargetsHash(targetsMap),
      success_mode: SUCCESS_MODE,
      runner_version: RUNNER_VERSION,
      answer_counters: deriveAnswerCountersFromCompleted(completed),
      gender_counter: normalizeGenderCounter({}),
      likert_progress: {},
      last_updated: null,
    };
  };

  if (RESET_QUOTA_STATE) {
    return makeInitialState();
  }

  if (!fs.existsSync(QUOTA_STATE_PATH)) {
    return makeInitialState();
  }

  try {
    const raw = JSON.parse(fs.readFileSync(QUOTA_STATE_PATH, "utf8"));
    const completed = normalizeCompletedMap(targetsMap, raw?.completed ?? {});
    const completedTotal = Object.values(completed).reduce(
      (sum, value) => sum + sanitizeTargetValue(value, 0),
      0,
    );
    return {
      completed,
      attempts_total: sanitizeTargetValue(raw?.attempts_total, 0),
      success_total: Math.max(
        completedTotal,
        sanitizeTargetValue(raw?.success_total, completedTotal),
      ),
      submit_click_total: sanitizeTargetValue(raw?.submit_click_total, 0),
      submit_click_unconfirmed_total: sanitizeTargetValue(
        raw?.submit_click_unconfirmed_total,
        0,
      ),
      attempt_status_counter: normalizeAttemptStatusCounter(raw?.attempt_status_counter),
      last_attempt_id: sanitizeTargetValue(raw?.last_attempt_id, 0),
      form_url: String(raw?.form_url ?? ""),
      quota_targets_hash: String(raw?.quota_targets_hash ?? ""),
      success_mode: String(raw?.success_mode ?? ""),
      runner_version: String(raw?.runner_version ?? ""),
      answer_counters: deriveAnswerCountersFromCompleted(completed),
      gender_counter: normalizeGenderCounter(raw?.gender_counter),
      likert_progress: normalizeLikertProgress(raw?.likert_progress),
      last_updated: raw?.last_updated ?? null,
    };
  } catch {
    return makeInitialState();
  }
};

const writeJsonAtomic = (filePath, value) => {
  const tmpPath = `${filePath}.tmp-${process.pid}-${Date.now()}`;
  fs.writeFileSync(tmpPath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
  fs.renameSync(tmpPath, filePath);
};

const saveQuotaState = (state) => {
  const derivedAnswerCounters = deriveAnswerCountersFromCompleted(state.completed ?? {});
  const normalizedLikertProgress = normalizeLikertProgress(state.likert_progress);
  const normalizedGenderCounter = normalizeGenderCounter(state.gender_counter);
  const now = new Date().toISOString();
  const answerCounters = state.answer_counters ?? derivedAnswerCounters;
  state.answer_counters = answerCounters;
  state.gender_counter = normalizedGenderCounter;
  state.likert_progress = normalizedLikertProgress;
  state.last_updated = now;
  const payload = {
    completed: state.completed,
    attempts_total: state.attempts_total,
    success_total: state.success_total,
    submit_click_total: sanitizeTargetValue(state.submit_click_total, 0),
    submit_click_unconfirmed_total: sanitizeTargetValue(
      state.submit_click_unconfirmed_total,
      0,
    ),
    attempt_status_counter: normalizeAttemptStatusCounter(state.attempt_status_counter),
    last_attempt_id: sanitizeTargetValue(state.last_attempt_id, 0),
    form_url: normalizeFormUrlForState(state.form_url || FORM_URL),
    quota_targets_hash: String(state.quota_targets_hash || ""),
    success_mode: String(state.success_mode || SUCCESS_MODE),
    runner_version: String(state.runner_version || RUNNER_VERSION),
    answer_counters: answerCounters,
    gender_counter: normalizedGenderCounter,
    likert_progress: normalizedLikertProgress,
    last_updated: now,
  };
  writeJsonAtomic(QUOTA_STATE_PATH, payload);
};

const buildQuotaRemaining = (targetsMap, completedMap) => {
  const remaining = {};
  let totalTarget = 0;
  let totalCompleted = 0;
  let totalRemaining = 0;
  for (const [key, targetRaw] of Object.entries(targetsMap)) {
    const target = sanitizeTargetValue(targetRaw, 0);
    const completed = sanitizeTargetValue(completedMap?.[key], 0);
    const rem = Math.max(0, target - completed);
    remaining[key] = rem;
    totalTarget += target;
    totalCompleted += Math.min(target, completed);
    totalRemaining += rem;
  }
  return {
    remaining,
    total_target: totalTarget,
    total_completed: totalCompleted,
    total_remaining: totalRemaining,
  };
};

const validateQuotaStateCompatibility = (quotaState, targetsMap) => {
  const expected = {
    form_url: normalizeFormUrlForState(FORM_URL),
    quota_targets_hash: computeQuotaTargetsHash(targetsMap),
    success_mode: SUCCESS_MODE,
  };
  const actual = {
    form_url: normalizeFormUrlForState(quotaState?.form_url || ""),
    quota_targets_hash: String(quotaState?.quota_targets_hash || ""),
    success_mode: String(quotaState?.success_mode || ""),
  };

  if (actual.form_url !== expected.form_url) {
    return { ok: false, reason: "form_url_mismatch", expected, actual };
  }
  if (actual.quota_targets_hash !== expected.quota_targets_hash) {
    return { ok: false, reason: "quota_targets_hash_mismatch", expected, actual };
  }
  if (actual.success_mode !== expected.success_mode) {
    return { ok: false, reason: "success_mode_mismatch", expected, actual };
  }
  return { ok: true, reason: null, expected, actual };
};

const pickNextQuotaCell = (quotaCells, remainingMap) => {
  const candidates = quotaCells
    .map((cell) => ({ cell, rem: sanitizeTargetValue(remainingMap?.[cell.key], 0) }))
    .filter((entry) => entry.rem > 0)
    .sort((a, b) => {
      if (b.rem !== a.rem) return b.rem - a.rem;
      return a.cell.key.localeCompare(b.cell.key);
    });
  return candidates.length ? candidates[0].cell : null;
};

const normalizeGenderCounter = (raw) => ({
  male: sanitizeTargetValue(raw?.male, 0),
  female: sanitizeTargetValue(raw?.female, 0),
});

const pickBalancedGender = ({ maleIndex, femaleIndex, runState }) => {
  const counter = normalizeGenderCounter(runState?.gender_counter);
  const expectedTotal = sanitizeTargetValue(runState?.gender_expected_total, 0);
  const maleTarget = expectedTotal > 0 ? Math.floor(expectedTotal / 2) : null;
  const femaleTarget =
    expectedTotal > 0 ? expectedTotal - Math.floor(expectedTotal / 2) : null;

  if (maleTarget !== null && femaleTarget !== null) {
    const maleRem = maleTarget - counter.male;
    const femaleRem = femaleTarget - counter.female;
    if (maleRem > femaleRem) {
      return { index: maleIndex, gender: "male", rule_id: "quota_gender_balance" };
    }
    if (femaleRem > maleRem) {
      return { index: femaleIndex, gender: "female", rule_id: "quota_gender_balance" };
    }
  }

  if (counter.male < counter.female) {
    return { index: maleIndex, gender: "male", rule_id: "quota_gender_balance" };
  }
  if (counter.female < counter.male) {
    return { index: femaleIndex, gender: "female", rule_id: "quota_gender_balance" };
  }

  const parity = sanitizeTargetValue(runState?.attempt_id, 1) % 2;
  if (parity === 0) {
    return { index: femaleIndex, gender: "female", rule_id: "quota_gender_balance" };
  }
  return { index: maleIndex, gender: "male", rule_id: "quota_gender_balance" };
};

const buildLikertSegmentExpectedTotalsFromQuota = (targetsMap) => {
  const out = {};
  if (!targetsMap || typeof targetsMap !== "object") return out;
  for (const [cellKey, targetRaw] of Object.entries(targetsMap)) {
    const parsed = parseQuotaCellKey(cellKey);
    if (!parsed) continue;
    const segmentKey = normalizeSegmentKey(
      `${parsed.city_key}|${parsed.trip_type}|${parsed.children_mode}`,
    );
    if (!segmentKey) continue;
    out[segmentKey] = sanitizeTargetValue(out[segmentKey], 0) + sanitizeTargetValue(targetRaw, 0);
  }
  return out;
};

const createEmptyAnswerCounters = () => ({
  city: Object.fromEntries(QUOTA_CITY_DEFS.map((city) => [city.city_key, 0])),
  trip_type: {
    vacation: 0,
    weekend: 0,
  },
  children_mode: {
    with_children: 0,
    without_children: 0,
  },
  provider_bucket: {
    yp: 0,
    other: 0,
  },
  organizer: {
    yes: 0,
    no: 0,
  },
});

const deriveAnswerCountersFromCompleted = (completedMap) => {
  const counters = createEmptyAnswerCounters();
  for (const [key, rawCount] of Object.entries(completedMap ?? {})) {
    const parsed = parseQuotaCellKey(key);
    if (!parsed) continue;
    const count = sanitizeTargetValue(rawCount, 0);
    counters.city[parsed.city_key] =
      sanitizeTargetValue(counters.city[parsed.city_key], 0) + count;
    counters.trip_type[parsed.trip_type] =
      sanitizeTargetValue(counters.trip_type[parsed.trip_type], 0) + count;
    counters.children_mode[parsed.children_mode] =
      sanitizeTargetValue(counters.children_mode[parsed.children_mode], 0) + count;
    counters.provider_bucket[parsed.provider_bucket] =
      sanitizeTargetValue(counters.provider_bucket[parsed.provider_bucket], 0) + count;
    counters.organizer.yes += count;
  }
  return counters;
};

const toCityLabelCounters = (cityCounter) =>
  Object.fromEntries(
    QUOTA_CITY_DEFS.map((city) => [
      city.city_label,
      sanitizeTargetValue(cityCounter?.[city.city_key], 0),
    ]),
  );

const mergeAnswerCountersWithGender = (baseCounters, genderCounter) => ({
  ...baseCounters,
  gender: {
    male: sanitizeTargetValue(genderCounter?.male, 0),
    female: sanitizeTargetValue(genderCounter?.female, 0),
  },
});

const normalizeLikertRuleEntry = (entry, index = 0) => {
  if (!entry || typeof entry !== "object" || Array.isArray(entry)) return null;
  const id = String(entry.id ?? `likert_${index + 1}`);
  const targetMeanRaw =
    entry.target_mean ?? entry.mean ?? entry.avg ?? entry.average ?? null;
  const targetMean = Number(targetMeanRaw);
  if (!Number.isFinite(targetMean)) return null;

  const includesRaw =
    entry.question_includes ??
    entry.match ??
    entry.contains ??
    entry.question ??
    entry.label ??
    null;
  const includes = Array.isArray(includesRaw)
    ? includesRaw
    : typeof includesRaw === "string"
      ? [includesRaw]
      : [];
  const normalizedIncludes = includes
    .map((value) => normalizeRuleText(value))
    .filter(Boolean);
  if (!normalizedIncludes.length) return null;

  return {
    id,
    target_mean: Math.min(5, Math.max(1, targetMean)),
    question_includes: normalizedIncludes,
  };
};

const makeLikertSegmentKeyFromCell = (cell) =>
  cell ? `${cell.city_key}|${cell.trip_type}|${cell.children_mode}` : "__global__";

const normalizeSegmentKey = (value) => {
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  const parts = raw.split("|").map((part) => part.trim());
  if (parts.length === 3) return `${parts[0]}|${parts[1]}|${parts[2]}`;
  if (parts.length === 4) return `${parts[0]}|${parts[1]}|${parts[2]}`;
  return raw;
};

const parseOrderedMeansInput = (value) => {
  if (Array.isArray(value)) return value;
  if (typeof value !== "string") return [];
  return value
    .split(/[\t;\n\r ]+/)
    .map((token) => token.trim())
    .filter(Boolean)
    .map((token) => {
      const normalized = token.replace(",", ".");
      const parsed = Number(normalized);
      return Number.isFinite(parsed) ? parsed : null;
    });
};

const toFiniteNumberOrNull = (value) => {
  if (value === null || value === undefined) return null;
  if (typeof value === "string" && value.trim() === "") return null;
  const n = Number(
    typeof value === "string" ? value.replace(",", ".").trim() : value,
  );
  return Number.isFinite(n) ? n : null;
};

const buildLikertTargetsByOrder = ({
  orderedValues,
  orderIds,
}) => {
  const normalizedValues = parseOrderedMeansInput(orderedValues);
  const out = {};
  for (let i = 0; i < orderIds.length; i += 1) {
    const id = orderIds[i];
    const value = normalizedValues[i];
    if (!Number.isFinite(Number(value))) continue;
    out[id] = Math.min(5, Math.max(1, Number(value)));
  }
  return out;
};

const loadLikertTargets = () => {
  if (!LIKERT_TARGETS_PATH || !fs.existsSync(LIKERT_TARGETS_PATH)) {
    return {
      expected_total: null,
      rules: [],
      segment_targets: {},
      segment_expected_total: {},
    };
  }

  try {
    const raw = JSON.parse(fs.readFileSync(LIKERT_TARGETS_PATH, "utf8"));
    const expectedTotal = sanitizeTargetValue(
      raw?.expected_total ?? raw?.total_expected ?? null,
      0,
    );

    let entries = [];
    if (Array.isArray(raw)) {
      entries = raw;
    } else if (Array.isArray(raw?.targets)) {
      entries = raw.targets;
    } else if (Array.isArray(raw?.rules)) {
      entries = raw.rules;
    } else if (
      raw?.targets &&
      typeof raw.targets === "object" &&
      !raw?.segment_targets &&
      !raw?.profiles
    ) {
      entries = Object.entries(raw.targets).map(([match, mean], idx) => ({
        id: `likert_map_${idx + 1}`,
        match,
        target_mean: mean,
      }));
    } else if (
      raw &&
      typeof raw === "object" &&
      !raw?.segment_targets &&
      !raw?.profiles
    ) {
      entries = Object.entries(raw).map(([match, mean], idx) => ({
        id: `likert_map_${idx + 1}`,
        match,
        target_mean: mean,
      }));
    }

    const rules = entries
      .map((entry, idx) => normalizeLikertRuleEntry(entry, idx))
      .filter(Boolean);
    const useDefaultRules =
      raw?.use_default_rules_9_13 === true ||
      Boolean(raw?.segment_targets) ||
      Boolean(raw?.profiles);
    const finalRules = useDefaultRules
      ? LIKERT_DEFAULT_RULES_9_13.map((rule) => ({
          id: rule.id,
          question_includes: rule.question_includes.map((x) => normalizeRuleText(x)),
          target_mean: null,
        }))
      : rules;

    const segmentTargetsRaw = raw?.segment_targets ?? raw?.profiles ?? {};
    const segmentTargets = {};
    if (segmentTargetsRaw && typeof segmentTargetsRaw === "object") {
      for (const [segmentKeyRaw, targetValue] of Object.entries(segmentTargetsRaw)) {
        const segmentKey = normalizeSegmentKey(segmentKeyRaw);
        if (!segmentKey) continue;
        let perRule = {};

        if (Array.isArray(targetValue) || typeof targetValue === "string") {
          const values = parseOrderedMeansInput(targetValue);
          const orderIds =
            values.length === LIKERT_ORDER_SKIP_10_IDS.length
              ? LIKERT_ORDER_SKIP_10_IDS
              : LIKERT_ORDER_ALL_IDS;
          perRule = buildLikertTargetsByOrder({
            orderedValues: values,
            orderIds,
          });
        } else if (targetValue && typeof targetValue === "object") {
          if (Array.isArray(targetValue?.ordered_values) || typeof targetValue?.ordered_values === "string") {
            const orderMode = String(targetValue?.order ?? raw?.profile_order ?? "").toLowerCase();
            const orderIds =
              orderMode === "skip_10"
                ? LIKERT_ORDER_SKIP_10_IDS
                : orderMode === "all"
                  ? LIKERT_ORDER_ALL_IDS
                  : parseOrderedMeansInput(targetValue.ordered_values).length ===
                      LIKERT_ORDER_SKIP_10_IDS.length
                    ? LIKERT_ORDER_SKIP_10_IDS
                    : LIKERT_ORDER_ALL_IDS;
            perRule = buildLikertTargetsByOrder({
              orderedValues: targetValue.ordered_values,
              orderIds,
            });
          } else {
            for (const [id, mean] of Object.entries(targetValue)) {
              const n = Number(String(mean).replace(",", "."));
              if (!Number.isFinite(n)) continue;
              perRule[String(id)] = Math.min(5, Math.max(1, n));
            }
          }
        }

        if (Object.keys(perRule).length) {
          segmentTargets[segmentKey] = perRule;
        }
      }
    }

    const segmentExpectedTotalRaw = raw?.segment_expected_total;
    const segmentExpectedTotal = {};
    if (segmentExpectedTotalRaw && typeof segmentExpectedTotalRaw === "object") {
      for (const [segmentKeyRaw, value] of Object.entries(segmentExpectedTotalRaw)) {
        const segmentKey = normalizeSegmentKey(segmentKeyRaw);
        if (!segmentKey) continue;
        const n = sanitizeTargetValue(value, 0);
        if (n > 0) {
          segmentExpectedTotal[segmentKey] = n;
        }
      }
    }

    return {
      expected_total: expectedTotal > 0 ? expectedTotal : null,
      rules: finalRules,
      segment_targets: segmentTargets,
      segment_expected_total: segmentExpectedTotal,
    };
  } catch {
    return {
      expected_total: null,
      rules: [],
      segment_targets: {},
      segment_expected_total: {},
    };
  }
};

const parseLikertScoreFromLabel = (label) => {
  const normalized = normalizeText(label);
  const match = /^([1-5])(?:\b|[^\d]|$)/.exec(normalized);
  if (!match) return null;
  return Number(match[1]);
};

const buildLikertScoreToIndexMap = (labels) => {
  const scoreToIndex = new Map();
  for (let index = 0; index < labels.length; index += 1) {
    const score = parseLikertScoreFromLabel(labels[index]);
    if (score === null) continue;
    if (!scoreToIndex.has(score)) {
      scoreToIndex.set(score, index);
    }
  }
  return scoreToIndex;
};

const pickScoreWithRandomness = (scores, preferredScore, randomness = 0.35) => {
  const candidates = Array.from(
    new Set(
      (scores ?? [])
        .map((v) => clampLikertScore(v))
        .filter((v) => Number.isFinite(v)),
    ),
  ).sort((a, b) => a - b);
  if (!candidates.length) return clampLikertScore(preferredScore);
  if (candidates.length === 1) return candidates[0];

  const preferred = clampLikertScore(preferredScore);
  const r = Math.min(1, Math.max(0, Number(randomness ?? 0)));
  if (r <= 0) {
    return candidates.sort((a, b) => {
      const diff = Math.abs(a - preferred) - Math.abs(b - preferred);
      if (diff !== 0) return diff;
      return b - a;
    })[0];
  }

  const weighted = candidates.map((score) => {
    const dist = Math.abs(score - preferred);
    const closeness = 1 / (1 + dist);
    const weight = closeness * (1 - r) + r;
    return { score, weight };
  });
  const totalWeight = weighted.reduce((acc, item) => acc + item.weight, 0);
  if (!Number.isFinite(totalWeight) || totalWeight <= 0) return candidates[0];
  let threshold = Math.random() * totalWeight;
  for (const item of weighted) {
    threshold -= item.weight;
    if (threshold <= 0) return item.score;
  }
  return weighted[weighted.length - 1].score;
};

const resolveLikertFallbackIndex = (
  labels,
  preferredScore = LIKERT_NO_TARGET_SCORE,
) => {
  const scoreToIndex = buildLikertScoreToIndexMap(labels);
  if (!scoreToIndex.size) return -1;
  const pickedScore = pickScoreWithRandomness(
    [...scoreToIndex.keys()],
    preferredScore,
    LIKERT_FALLBACK_RANDOMNESS,
  );
  return scoreToIndex.get(pickedScore) ?? -1;
};

const findLikertTargetRule = (questionText, rules) => {
  if (!Array.isArray(rules) || !rules.length) return null;
  const normalizedQuestion = normalizeRuleText(questionText);
  for (const rule of rules) {
    if (!rule?.question_includes?.length) continue;
    const ok = rule.question_includes.every((needle) =>
      normalizedQuestion.includes(needle),
    );
    if (ok) return rule;
  }
  return null;
};

const pickLikertScoreByTarget = ({
  targetMean,
  currentCount,
  currentSum,
  expectedTotal,
}) => {
  const safeTarget = Math.min(5, Math.max(1, Number(targetMean)));
  if (!Number.isFinite(safeTarget)) return 3;

  const n = sanitizeTargetValue(currentCount, 0);
  const sum = Number(currentSum ?? 0);
  if (!Number.isFinite(sum)) return clampLikertScore(safeTarget);

  if (!expectedTotal || expectedTotal <= 0) {
    const preferred = clampLikertScore(safeTarget * (n + 1) - sum);
    return pickScoreWithRandomness([1, 2, 3, 4, 5], preferred, LIKERT_TARGET_RANDOMNESS);
  }

  const total = sanitizeTargetValue(expectedTotal, 0);
  const remainingAfter = Math.max(0, total - (n + 1));
  const targetSum = safeTarget * total;
  let score = clampLikertScore(safeTarget * (n + 1) - sum);

  const minFeasible = Math.max(
    1,
    Math.ceil(targetSum - sum - remainingAfter * 5),
  );
  const maxFeasible = Math.min(
    5,
    Math.floor(targetSum - sum - remainingAfter * 1),
  );
  if (minFeasible <= maxFeasible) {
    if (score < minFeasible) score = minFeasible;
    if (score > maxFeasible) score = maxFeasible;
    const feasibleScores = [];
    for (let s = minFeasible; s <= maxFeasible; s += 1) {
      feasibleScores.push(s);
    }
    return pickScoreWithRandomness(feasibleScores, score, LIKERT_TARGET_RANDOMNESS);
  }

  return clampLikertScore(score);
};

const resolveLikertTargetMeanByRuleId = (ruleId, runState) => {
  const segmentKey = runState?.likert_segment_key ?? "__global__";
  const fromSegment = runState?.likert_segment_targets?.[segmentKey]?.[ruleId];
  const segmentValue = toFiniteNumberOrNull(fromSegment);
  if (Number.isFinite(segmentValue)) return segmentValue;
  const fromGlobal = runState?.likert_segment_targets?.__global__?.[ruleId];
  const globalValue = toFiniteNumberOrNull(fromGlobal);
  if (Number.isFinite(globalValue)) return globalValue;
  return null;
};

const resolveLikertFallbackOrder = (runState) => {
  const segmentKey = runState?.likert_segment_key ?? "__global__";
  const targetMap = runState?.likert_segment_targets?.[segmentKey] ?? {};
  const targetIds = Object.keys(targetMap);
  const has10 = targetIds.some((id) => String(id).startsWith("10:"));
  if (!has10 && targetIds.length > 0 && targetIds.length <= LIKERT_ORDER_SKIP_10_IDS.length) {
    return LIKERT_ORDER_SKIP_10_IDS;
  }
  return LIKERT_ORDER_ALL_IDS;
};

const resolveQuestionIdFromSchema = (runState, questionText, fallbackQuestionId = "") => {
  const normalizedFallback = String(fallbackQuestionId ?? "").trim().toLowerCase();
  if (normalizedFallback) {
    return normalizedFallback.startsWith("q") ? normalizedFallback : `q${normalizedFallback}`;
  }
  if (SELECTOR_STRATEGY !== "text") return "";
  const map = runState?.schema?.question_text_to_id;
  if (!map || typeof map !== "object") return "";
  const key = normalizeRuleText(questionText);
  if (!key) return "";
  const hit = map[key];
  if (!Number.isFinite(Number(hit))) return "";
  return `q${Number(hit)}`;
};

const getLikertProgressItem = (runState, progressKey) =>
  runState?.likert_progress?.[progressKey] ?? {
    count: 0,
    sum: 0,
  };

const decideTableScaleScore = ({
  tableQuestionId,
  rowIndex,
  rowText,
  runState,
}) => {
  const segmentKey = runState?.likert_segment_key ?? "__global__";
  const rowMapKey = `${tableQuestionId}|${rowIndex}`;
  let rule = findLikertTargetRule(rowText, runState?.likert_rules ?? []);
  let mappedBy = "text";

  if (!rule) {
    const existing = runState?.table_row_rule_map?.get(rowMapKey);
    if (existing) {
      rule = { id: existing };
      mappedBy = "order_cache";
    } else {
      const fallbackOrder = resolveLikertFallbackOrder(runState);
      if (fallbackOrder.length) {
        const cursor = sanitizeTargetValue(runState?.table_row_fallback_cursor, 0);
        const selectedRuleId = fallbackOrder[cursor % fallbackOrder.length];
        runState.table_row_fallback_cursor = cursor + 1;
        runState.table_row_rule_map.set(rowMapKey, selectedRuleId);
        rule = { id: selectedRuleId };
        mappedBy = "order";
      }
    }
  }

  let ruleId = rule?.id ?? null;
  let source = "rule_based_likert";
  if (!ruleId) {
    ruleId = `unmapped:${tableQuestionId}:${rowIndex}`;
    source = "rule_based_likert_unmapped";
  }
  const progressKey = `${segmentKey}|${ruleId}`;
  const targetMean = resolveLikertTargetMeanByRuleId(ruleId, runState);
  let score = LIKERT_NO_TARGET_SCORE;

  if (Number.isFinite(Number(targetMean))) {
    const progressItem = getLikertProgressItem(runState, progressKey);
    const expectedTotal =
      runState?.likert_segment_expected_total?.[segmentKey] ??
      runState?.likert_expected_total ??
      null;
    score = pickLikertScoreByTarget({
      targetMean,
      currentCount: progressItem.count,
      currentSum: progressItem.sum,
      expectedTotal,
    });
  } else {
    score = pickScoreWithRandomness(
      [1, 2, 3, 4, 5],
      LIKERT_NO_TARGET_SCORE,
      LIKERT_FALLBACK_RANDOMNESS,
    );
    source =
      source === "rule_based_likert_unmapped"
        ? "rule_based_likert_unmapped"
        : "rule_based_likert_fallback_score";
  }

  return {
    rule_id: ruleId,
    progress_key: progressKey,
    segment_key: segmentKey,
    target_mean: targetMean,
    score: clampLikertScore(score),
    source,
    mapped_by: mappedBy,
  };
};

const createLikertValueCounter = () => ({
  1: 0,
  2: 0,
  3: 0,
  4: 0,
  5: 0,
});

const normalizeLikertProgress = (raw) => {
  const normalized = {};
  if (!raw || typeof raw !== "object") return normalized;
  for (const [id, item] of Object.entries(raw)) {
    if (!item || typeof item !== "object") continue;
    const count = sanitizeTargetValue(item.count, 0);
    const sum = Number(item.sum ?? 0);
    const targetMeanRaw = Number(item.target_mean ?? item.mean ?? 0);
    const targetMean = Number.isFinite(targetMeanRaw)
      ? Math.min(5, Math.max(1, targetMeanRaw))
      : null;
    const valueCounter = createLikertValueCounter();
    for (const score of [1, 2, 3, 4, 5]) {
      valueCounter[score] = sanitizeTargetValue(item?.value_counter?.[score], 0);
    }
    normalized[id] = {
      count,
      sum: Number.isFinite(sum) ? sum : 0,
      target_mean: targetMean,
      value_counter: valueCounter,
    };
  }
  return normalized;
};

const applyLikertAnswersToProgress = (progressRaw, answers) => {
  const progress = normalizeLikertProgress(progressRaw);
  if (!Array.isArray(answers)) return progress;

  for (const answer of answers) {
    if (!answer || typeof answer !== "object") continue;
    const progressKey = String(answer.progress_key ?? answer.rule_id ?? answer.id ?? "");
    if (!progressKey) continue;
    const score = clampLikertScore(answer.value);
    if (!Number.isFinite(score)) continue;
    if (!progress[progressKey]) {
      progress[progressKey] = {
        count: 0,
        sum: 0,
        target_mean: Number.isFinite(Number(answer.target_mean))
          ? Math.min(5, Math.max(1, Number(answer.target_mean)))
          : null,
        value_counter: createLikertValueCounter(),
      };
    }
    const item = progress[progressKey];
    item.count += 1;
    item.sum += score;
    item.value_counter[score] += 1;
    if (Number.isFinite(Number(answer.target_mean))) {
      item.target_mean = Math.min(5, Math.max(1, Number(answer.target_mean)));
    }
  }

  return progress;
};

const buildLikertProgressSummary = (progressRaw) => {
  const progress = normalizeLikertProgress(progressRaw);
  const summary = {};
  for (const [progressKey, item] of Object.entries(progress)) {
    const parts = progressKey.split("|");
    const ruleId = parts.length >= 4 ? parts.slice(3).join("|") : progressKey;
    const segmentKey = parts.length >= 4 ? parts.slice(0, 3).join("|") : "__global__";
    const meanCurrent = item.count > 0 ? item.sum / item.count : null;
    summary[progressKey] = {
      segment_key: segmentKey,
      rule_id: ruleId,
      target_mean: item.target_mean,
      mean_current: meanCurrent,
      count: item.count,
      sum: item.sum,
      value_counter: item.value_counter,
    };
  }
  return summary;
};

async function detectCaptcha(page) {
  const bodyText = (await page.locator("body").innerText().catch(() => ""))
    .toLowerCase()
    .trim();
  if (/(captcha|капча|не робот|robot|challenge)/i.test(bodyText)) return true;

  const frames = await page.locator("iframe").all().catch(() => []);
  for (const frame of frames) {
    const src = (await frame.getAttribute("src").catch(() => "")) ?? "";
    if (/(captcha|challenge|recaptcha|hcaptcha|smartcaptcha)/i.test(src)) {
      return true;
    }
  }
  return false;
}

async function detectLoadErrorScreen(page) {
  const bodyText = (await page.locator("body").innerText().catch(() => "")).trim();
  return (
    /что-то пошло не так/i.test(bodyText) &&
    /не удалось загрузить данные/i.test(bodyText)
  );
}

async function loadPlaywright() {
  try {
    return await import("playwright");
  } catch (error) {
    const e = new Error(
      "Missing dependency: playwright. Install it with `npm install playwright --no-save`.",
    );
    e.cause = error;
    throw e;
  }
}

const buildLaunchOptions = () => {
  const launchOptions = { headless: HEADLESS };
  if (BROWSER_NAME === "chromium") {
    if (BROWSER_EXECUTABLE_PATH) {
      launchOptions.executablePath = BROWSER_EXECUTABLE_PATH;
    } else if (BROWSER_CHANNEL) {
      launchOptions.channel = BROWSER_CHANNEL;
    }
    if (INCOGNITO) {
      launchOptions.args = [...(launchOptions.args ?? []), "--incognito"];
    }
  }
  return launchOptions;
};

const getLabelText = (value) => {
  if (typeof value === "string") return value;
  if (value && typeof value === "object") {
    if (typeof value.text === "string") return value.text;
    if (typeof value.label === "string") return value.label;
  }
  return "";
};

const buildSchemaFromState = (state, storageKey = null) => {
  const pagesList = Array.isArray(state?.pages?.list) ? state.pages.list : [];
  const questionsRaw = state?.questions && typeof state.questions === "object"
    ? state.questions
    : {};

  const questionTextToId = {};
  const questionsById = {};
  const questionTypeHistogram = {};

  for (const [idRaw, q] of Object.entries(questionsRaw)) {
    const id = Number(idRaw);
    if (!Number.isFinite(id)) continue;
    const type = String(q?.type ?? "unknown");
    questionTypeHistogram[type] = sanitizeTargetValue(questionTypeHistogram[type], 0) + 1;
    const labelText = getLabelText(q?.label);
    const labelNorm = normalizeRuleText(labelText);
    if (labelNorm && !(labelNorm in questionTextToId)) {
      questionTextToId[labelNorm] = id;
    }

    const options = Array.isArray(q?.options)
      ? q.options.map((opt, idx) => {
          const optionText = getLabelText(opt?.label ?? opt);
          return {
            index: idx,
            id: opt?.id ?? null,
            label: optionText,
            label_norm: normalizeRuleText(optionText),
          };
        })
      : [];
    const optionsByText = {};
    for (const opt of options) {
      if (!opt.label_norm) continue;
      if (!(opt.label_norm in optionsByText)) {
        optionsByText[opt.label_norm] = opt.index;
      }
    }

    const rows = Array.isArray(q?.rows)
      ? q.rows.map((row, idx) => {
          const rowText = getLabelText(row?.label ?? row);
          return {
            index: idx,
            id: row?.id ?? null,
            label: rowText,
            label_norm: normalizeRuleText(rowText),
          };
        })
      : [];

    questionsById[id] = {
      id,
      type,
      required: Boolean(q?.required),
      label: labelText,
      label_norm: labelNorm,
      show_if: q?.showIf ?? null,
      options,
      options_by_text: optionsByText,
      rows,
      bound_questions: Array.isArray(q?.boundQuestions) ? q.boundQuestions : [],
      min_value: q?.minValue ?? null,
      max_value: q?.maxValue ?? null,
      no_opinion_enabled: Boolean(q?.noOpinionAnswer?.enabled),
    };
  }

  const pages = pagesList.map((page) => ({
    id: page?.id ?? null,
    question_ids: Array.isArray(page?.questionIds) ? page.questionIds : [],
    next_btn_label: page?.nextBtnLabel ?? null,
    show_if: page?.showIf ?? null,
  }));

  return {
    source: "localStorage",
    storage_key: storageKey,
    pages_count: pages.length,
    questions_count: Object.keys(questionsById).length,
    question_type_histogram: questionTypeHistogram,
    pages,
    questions_by_id: questionsById,
    question_text_to_id: questionTextToId,
  };
};

async function discoverSurveySchema(browserType) {
  if (SCHEMA_DISCOVERY_MODE !== "startup_localstorage") {
    return {
      source: "disabled",
      storage_key: null,
      pages_count: null,
      questions_count: null,
      question_type_histogram: {},
      pages: [],
      questions_by_id: {},
      question_text_to_id: {},
    };
  }

  const pollId = extractPollId(FORM_URL);
  const expectedKey = resolveSchemaStorageKey(pollId, SCHEMA_STORAGE_KEY);
  const launchOptions = buildLaunchOptions();
  let browser = null;
  let context = null;
  let page = null;

  try {
    browser = await browserType.launch(launchOptions);
    context = await browser.newContext();
    page = await context.newPage();
    await page.goto(`${FORM_URL}?fresh=schema-${Date.now()}`, {
      waitUntil: "domcontentloaded",
      timeout: 45_000,
    });
    await clickCookieBannerIfPresent(page);
    await page.waitForTimeout(300);

    const extracted = await page.evaluate(
      ({ expectedKey, pollId }) => {
        const keys = Object.keys(window.localStorage ?? {});
        const key =
          (expectedKey && keys.includes(expectedKey) && expectedKey) ||
          keys.find((item) => item.startsWith("pythia10-") && pollId && item.includes(pollId)) ||
          keys.find((item) => item.startsWith("pythia10-")) ||
          null;
        if (!key) {
          return {
            storage_key: null,
            state: null,
            local_keys: keys,
          };
        }
        const raw = window.localStorage.getItem(key);
        if (!raw) {
          return {
            storage_key: key,
            state: null,
            local_keys: keys,
          };
        }
        try {
          return {
            storage_key: key,
            state: JSON.parse(raw),
            local_keys: keys,
          };
        } catch {
          return {
            storage_key: key,
            state: null,
            local_keys: keys,
          };
        }
      },
      { expectedKey, pollId },
    );

    if (!extracted?.state || typeof extracted.state !== "object") {
      return {
        source: "localStorage_missing",
        storage_key: extracted?.storage_key ?? expectedKey ?? null,
        pages_count: null,
        questions_count: null,
        question_type_histogram: {},
        pages: [],
        questions_by_id: {},
        question_text_to_id: {},
      };
    }

    return buildSchemaFromState(extracted.state, extracted.storage_key ?? expectedKey ?? null);
  } catch {
    return {
      source: "localStorage_error",
      storage_key: expectedKey,
      pages_count: null,
      questions_count: null,
      question_type_histogram: {},
      pages: [],
      questions_by_id: {},
      question_text_to_id: {},
    };
  } finally {
    await context?.close().catch(() => {});
    await browser?.close().catch(() => {});
  }
}

const classifyAttemptErrorStatus = (error) => {
  const message = String(error?.message ?? error ?? "");
  if (
    /ERR_NETWORK_CHANGED|ERR_INTERNET_DISCONNECTED|ERR_NETWORK_IO_SUSPENDED|ERR_TUNNEL_CONNECTION_FAILED/i.test(
      message,
    )
  ) {
    return "network_changed";
  }
  if (
    /Target page, context or browser has been closed|Browser has been closed|Execution context was destroyed/i.test(
      message,
    )
  ) {
    return "browser_closed";
  }
  return "attempt_error";
};

async function detectSuccessMarker(page) {
  const bodyText = await page.locator("body").innerText().catch(() => "");
  const match = bodyText.match(
    /(Спасибо[^\n]*|Ваш ответ[^\n]*|ответ\s*(принят|записан|отправлен)[^\n]*|Thank you[^\n]*|Response has been recorded[^\n]*)/i,
  );
  return match?.[1]?.trim() ?? null;
}

async function clickCookieBannerIfPresent(page) {
  const cookieButton = page
    .getByRole("button", {
      name: /Allow all|Разрешить все|Allow essential cookies/i,
    })
    .first();

  if (await cookieButton.count().catch(() => 0)) {
    try {
      await cookieButton.click({ timeout: 1200 });
    } catch {
      // best effort only
    }
  }
}

async function hasStateQuestionValue(page, runState, questionId) {
  const normalizedQuestionId = String(questionId ?? "").toLowerCase();
  const numericId = /^q\d+$/.test(normalizedQuestionId)
    ? Number(normalizedQuestionId.slice(1))
    : Number(normalizedQuestionId);
  if (!Number.isFinite(numericId)) return true;

  const storageKey =
    runState?.schema?.storage_key ??
    resolveSchemaStorageKey(extractPollId(FORM_URL), SCHEMA_STORAGE_KEY);
  if (!storageKey) return true;

  const result = await page
    .evaluate(
      ({ storageKey, numericId }) => {
        const raw = window.localStorage?.getItem(storageKey);
        if (!raw) return null;
        try {
          const parsed = JSON.parse(raw);
          const question = parsed?.questions?.[String(numericId)];
          if (!question) return null;
          const value = question?.value;
          if (value === undefined || value === null || value === "") return null;
          if (Array.isArray(value) && value.length === 0) return null;
          return value;
        } catch {
          return null;
        }
      },
      { storageKey, numericId },
    )
    .catch(() => null);

  return result !== null;
}

const normalizeQuestionIds = (items) => {
  const arr = Array.isArray(items) ? items : [];
  return arr
    .map((item) => String(item ?? "").trim())
    .filter((item) => /^q\d+$/i.test(item))
    .sort((a, b) => Number(a.slice(1)) - Number(b.slice(1)));
};

const resolveFormEngine = async (page) => {
  if (FORM_DOM_MODE === "pythia") return "pythia";
  if (FORM_DOM_MODE === "legacy") return "legacy";
  const pythiaRoot = await page
    .locator(".pythia-surveys-root, .pythia-surveys-wrapper")
    .count()
    .catch(() => 0);
  if (pythiaRoot > 0) return "pythia";
  return "legacy";
};

async function detectPagePosition(page) {
  const data = await page
    .evaluate(() => {
      const normalize = (value) =>
        String(value ?? "")
          .replace(/\s+/g, " ")
          .trim();
      const ids = Array.from(document.querySelectorAll("fieldset[id]"))
        .map((el) => el.getAttribute("id") || "")
        .filter((id) => /^q\d+$/i.test(id));
      const body = normalize(document.body?.innerText || "");
      const buttons = Array.from(document.querySelectorAll("button"))
        .map((button) => normalize(button.textContent))
        .filter(Boolean);
      const pageNode = document.querySelector(".page");
      const pageClass = normalize(pageNode?.className || "");
      const isPageFirst = /\bpage_first_yes\b/.test(pageClass);
      const hasNextButton = buttons.some((text) => /^(далее|next)$/i.test(text));
      const hasIntroText =
        /привет/i.test(body) &&
        (/мы проводим этот опрос/i.test(body) || /будем очень благодарны/i.test(body));
      return {
        question_ids: ids,
        has_city_question:
          /в каком из следующих тр[её]х городов вы проживаете/i.test(body) ||
          /в каком из следующих трех городов вы проживаете/i.test(body),
        has_load_error:
          /что-то пошло не так/i.test(body) &&
          /не удалось загрузить данные/i.test(body),
        has_thanks: /спасибо за участие в опросе/i.test(body),
        has_finish_button: buttons.some((text) =>
          /^(завершить опрос|отправить|submit|готово)$/i.test(text),
        ),
        has_intro_page: (isPageFirst && hasNextButton) || hasIntroText,
      };
    })
    .catch(() => ({
      question_ids: [],
      has_city_question: false,
      has_load_error: false,
      has_thanks: false,
      has_finish_button: false,
      has_intro_page: false,
    }));

  const questionIds = normalizeQuestionIds(data.question_ids);
  const hasQ1 = questionIds.includes("q1");
  const hasQ59 = questionIds.includes("q59");
  const hasQ60 = questionIds.includes("q60");
  const hasQ61 = questionIds.includes("q61");
  const isLastPage =
    (hasQ59 || hasQ60 || hasQ61) && data.has_thanks && data.has_finish_button;
  const isFirstPage = hasQ1 || data.has_city_question || data.has_intro_page;
  return {
    question_ids: questionIds,
    is_first_page: isFirstPage,
    is_last_page: isLastPage,
    has_intro_page: Boolean(data.has_intro_page),
    has_load_error: Boolean(data.has_load_error),
    has_q1: hasQ1,
    has_q59: hasQ59,
    has_q60: hasQ60,
    has_q61: hasQ61,
  };
}

async function ensureStartsFromFirstPage(page, context, attemptId, runState) {
  const maxRecovery = Math.max(
    0,
    sanitizeTargetValue(START_PAGE_RECOVERY_ATTEMPTS, 3),
  );
  const pageReloadRecovery = START_PAGE_RECOVERY_MODE === "page_reload";
  let lastPosition = null;

  for (let recovery = 0; recovery <= maxRecovery; recovery += 1) {
    lastPosition = await detectPagePosition(page);
    appendPageDebug(runState, {
      phase: "start_page_check",
      recovery_attempt: recovery,
      question_ids: lastPosition.question_ids,
      is_first_page: lastPosition.is_first_page,
      is_last_page: lastPosition.is_last_page,
      has_intro_page: lastPosition.has_intro_page,
      has_load_error: lastPosition.has_load_error,
    });
    runState.last_page_position = lastPosition;

    if (lastPosition.is_first_page) {
      return { ok: true, position: lastPosition, reason: null };
    }
    if (recovery >= maxRecovery) {
      break;
    }
    if (!pageReloadRecovery) {
      appendPageDebug(runState, {
        phase: "start_page_recovery_browser_restart",
        recovery_attempt: recovery,
        note: "page_reload_skipped_wait_retry",
      });
      await page.waitForTimeout(Math.min(3_000, 700 + recovery * 350));
      await clickCookieBannerIfPresent(page);
      continue;
    }

    await context.clearCookies().catch(() => {});
    await page.goto("https://yandex.ru", {
      waitUntil: "domcontentloaded",
      timeout: 30_000,
    }).catch(() => {});
    await page.evaluate(async () => {
      try {
        window.localStorage?.clear();
      } catch {
        // noop
      }
      try {
        window.sessionStorage?.clear();
      } catch {
        // noop
      }
      try {
        if (window.caches?.keys) {
          const keys = await window.caches.keys();
          await Promise.all(keys.map((key) => window.caches.delete(key)));
        }
      } catch {
        // noop
      }
      try {
        if (window.indexedDB?.databases) {
          const dbs = await window.indexedDB.databases();
          await Promise.all(
            dbs
              .map((db) => db?.name)
              .filter(Boolean)
              .map(
                (name) =>
                  new Promise((resolve) => {
                    const req = window.indexedDB.deleteDatabase(name);
                    req.onsuccess = () => resolve();
                    req.onerror = () => resolve();
                    req.onblocked = () => resolve();
                  }),
              ),
          );
        }
      } catch {
        // noop
      }
    }).catch(() => {});

    await page.goto(
      `${FORM_URL}?reset=attempt-${Date.now()}&fresh=${attemptId}-${recovery + 1}`,
      {
        waitUntil: "domcontentloaded",
        timeout: 45_000,
      },
    );
    await page.waitForTimeout(300);
    await clickCookieBannerIfPresent(page);
  }

  return {
    ok: false,
    position: lastPosition,
    reason: lastPosition?.has_load_error ? "load_error_screen" : "start_page_invalid",
  };
}

const getRadioStrategyOrder = () => {
  if (RADIO_FALLBACK_STRATEGY === "check_only") return ["check"];
  if (RADIO_FALLBACK_STRATEGY === "standard") return ["check", "click"];
  if (RADIO_FALLBACK_STRATEGY === "native_first") {
    return ["native", "label", "check", "click"];
  }
  return ["check", "click", "native", "label"];
};

const appendSelectionError = (runState, entry) => {
  if (!runState?.selection_errors) return;
  if (runState.selection_errors.length >= 250) return;
  runState.selection_errors.push({
    ...entry,
    ts: new Date().toISOString(),
  });
};

const appendPageDebug = (runState, entry) => {
  if (!LOG_PAGE_DEBUG) return;
  if (!runState?.page_debug) return;
  if (runState.page_debug.length >= 200) return;
  runState.page_debug.push({
    ...entry,
    ts: new Date().toISOString(),
  });
};

const applySelectedRadioValue = ({
  runState,
  pageNo,
  selectedValue,
  selectionSource,
  answerSource,
  likertMeta,
  pushAnswer,
}) => {
  runState.selection_stats[selectionSource] += 1;
  const tripType = inferTripType(selectedValue);
  const childrenMode = inferChildrenMode(selectedValue);
  const genderMode = inferGenderMode(selectedValue);
  if (tripType) runState.trip_type = tripType;
  if (childrenMode) runState.children_mode = childrenMode;
  if (genderMode) runState.gender_choice = genderMode;
  pushAnswer({
    page: pageNo,
    type: "radio",
    value: selectedValue,
    source:
      answerSource ??
      (likertMeta ? `likert_rule:${likertMeta.rule_id}` : undefined),
  });
  if (likertMeta && runState.likert_answers_map) {
    runState.likert_answers_map.set(likertMeta.progress_key, {
      rule_id: likertMeta.rule_id,
      progress_key: likertMeta.progress_key,
      segment_key: likertMeta.segment_key,
      target_mean: likertMeta.target_mean,
      value: likertMeta.score,
    });
  }
};

const pickProviderIndexForBucket = (labels, providerBucket) => {
  const normalizedLabels = labels.map((label) => normalizeText(label));
  const yandexIndex = normalizedLabels.findIndex((label) =>
    isYandexTravelLabel(label),
  );
  const otherIndices = normalizedLabels
    .map((label, index) => ({ label, index }))
    .filter((item) => isOtherProviderAllowedLabel(item.label))
    .map((item) => item.index);
  if (providerBucket === "yp") {
    return yandexIndex;
  }
  if (providerBucket === "other") {
    if (!otherIndices.length) return -1;
    if (otherIndices.length === 1) return otherIndices[0];
    const randomPos = randInt(0, otherIndices.length - 1);
    return otherIndices[randomPos];
  }
  return -1;
};

const decideRadioSelection = ({
  pageNo,
  questionText,
  labels,
  runState,
  questionId = "",
  optionMeta = [],
}) => {
  const radioCount = labels.length;
  const yesIndex = labels.findIndex((v) => /^да$/i.test(v));
  const noIndex = labels.findIndex((v) => /^нет$/i.test(v));
  const normalizedQuestionId = String(questionId || "").toLowerCase();
  const normalizedLabels = labels.map((label) => normalizeRuleText(label));
  const normalizedOptionMeta = Array.isArray(optionMeta) ? optionMeta : [];
  let pick = STRICT_RULES_ONLY ? 0 : randInt(0, radioCount - 1);
  let selectionSource = STRICT_RULES_ONLY ? "rule_based" : "random";
  let selectedValue = labels[pick];
  let answerSource;
  const incomeQuestion = isIncomeQuestion(questionText);
  let likertMeta = null;
  const likertDecision = findLikertDrivenDecision({
    questionText,
    labels,
    runState,
  });
  const likertFallbackIndex = resolveLikertFallbackIndex(labels);
  const concreteProviderListQuestion = isConcreteProviderListQuestion(
    questionText,
    labels,
  );
  const quotaDecision = findQuotaDrivenIndex({
    questionText,
    labels,
    runState,
  });
  const ruleDecision = findRuleDrivenIndex({
    pageNo,
    questionText,
    labels,
  });

  const targetCell = runState?.target_cell ?? null;
  if (normalizedQuestionId === "q1" && targetCell?.city_label) {
    const cityIndex = normalizedLabels.findIndex(
      (label) => label === normalizeRuleText(targetCell.city_label),
    );
    if (cityIndex !== -1) {
      pick = cityIndex;
      selectionSource = "rule_based";
      selectedValue = labels[pick];
      answerSource = "rule_based:q1_city";
    }
  } else if (normalizedQuestionId === "q5" || isOrganizerQuestion(questionText)) {
    let organizerYesIndex = yesIndex;
    if (organizerYesIndex === -1 && normalizedOptionMeta.length) {
      organizerYesIndex = normalizedOptionMeta.findIndex((meta) => {
        const valueRaw = String(meta?.value ?? "").trim();
        const inputId = normalizeRuleText(meta?.input_id);
        const testId = normalizeRuleText(meta?.testid);
        return (
          valueRaw === "14" ||
          inputId.includes("q5o0") ||
          testId.includes("q5o0")
        );
      });
    }
    if (organizerYesIndex !== -1) {
      pick = organizerYesIndex;
      selectionSource = "rule_based";
      selectedValue = labels[pick];
      answerSource = "rule_based:q5_organizer_yes";
    }
  } else if (normalizedQuestionId === "q7" && targetCell?.trip_type) {
    const vacationIndex = normalizedLabels.findIndex(
      (label) => label.includes("отпуск") || label.includes("длинная поездка"),
    );
    const weekendIndex = normalizedLabels.findIndex(
      (label) => label.includes("выходн") || label.includes("короткая поездка"),
    );
    const mappedIndex =
      targetCell.trip_type === "vacation" ? vacationIndex : weekendIndex;
    if (mappedIndex !== -1) {
      pick = mappedIndex;
      selectionSource = "rule_based";
      selectedValue = labels[pick];
      answerSource = "rule_based:q7_trip_type";
    }
  } else if (normalizedQuestionId === "q8" && targetCell?.children_mode) {
    const withChildrenIndex = normalizedLabels.findIndex((label) =>
      label.includes("с детьми"),
    );
    const withoutChildrenIndex = normalizedLabels.findIndex((label) =>
      label.includes("без детей"),
    );
    const mappedIndex =
      targetCell.children_mode === "with_children"
        ? withChildrenIndex
        : withoutChildrenIndex;
    if (mappedIndex !== -1) {
      pick = mappedIndex;
      selectionSource = "rule_based";
      selectedValue = labels[pick];
      answerSource = "rule_based:q8_children_mode";
    }
  } else if (concreteProviderListQuestion && targetCell?.provider_bucket) {
    const providerIndex = pickProviderIndexForBucket(
      labels,
      targetCell.provider_bucket,
    );
    if (providerIndex !== -1) {
      pick = providerIndex;
      selectionSource = "rule_based";
      selectedValue = labels[pick];
      const providerTag = normalizedQuestionId || "provider";
      answerSource = `rule_based:${providerTag}_provider_${targetCell.provider_bucket}`;
    }
  } else if (incomeQuestion) {
    pick = randInt(0, radioCount - 1);
    selectionSource = "random";
    selectedValue = labels[pick];
    answerSource = "income_random";
  } else if (likertDecision) {
    pick = likertDecision.index;
    selectionSource = "rule_based";
    selectedValue = labels[pick];
    likertMeta = likertDecision;
  } else if (STRICT_RULES_ONLY && likertFallbackIndex !== -1) {
    // For 1..5 scales without an explicit target mean, avoid deterministic bias to "1".
    pick = likertFallbackIndex;
    selectionSource = "rule_based";
    selectedValue = labels[pick];
    answerSource = "rule_based:likert_no_target_fallback";
  } else if (quotaDecision) {
    pick = quotaDecision.index;
    selectionSource = "rule_based";
    selectedValue = labels[pick];
  } else if (ruleDecision) {
    pick = ruleDecision.index;
    selectionSource = "rule_based";
    selectedValue = labels[pick];
  } else if (yesIndex !== -1 && noIndex !== -1) {
    const branchValue = resolveBranchChoice(runState);
    runState.branch_choice = branchValue;
    pick = branchValue === "Да" ? yesIndex : noIndex;
    selectedValue = labels[pick];
  }

  if (pick < 0 || pick >= radioCount) {
    pick = Math.max(0, Math.min(radioCount - 1, pick));
  }
  selectedValue = labels[pick];

  if (yesIndex !== -1 && noIndex !== -1) {
    if (pick === yesIndex) runState.branch_choice = "Да";
    if (pick === noIndex) runState.branch_choice = "Нет";
  }

  return {
    pick,
    selectionSource,
    selectedValue,
    answerSource,
    likertMeta,
  };
};

const extractQuestionTextFromGroup = async (group) =>
  group
    .evaluate((el) => {
      const normalize = (value) =>
        String(value ?? "")
          .replace(/\s+/g, " ")
          .trim();
      const aria = normalize(el.getAttribute("aria-label"));
      if (aria) return aria;
      const selectors =
        "h1,h2,h3,h4,legend,[role='rowheader'],.QuestionLabel,[id$='-label']";
      const ownHeading = el.querySelector(selectors);
      const ownText = normalize(ownHeading?.textContent);
      if (ownText) return ownText;
      let current = el;
      for (let i = 0; i < 8 && current; i += 1) {
        const candidate = current.querySelector(selectors);
        const text = normalize(candidate?.textContent);
        if (text) return text;
        current = current.parentElement;
      }
      return "";
    })
    .catch(() => "");

const hasCheckedInRoleGroup = async (group) => {
  const checkedByRole = await group
    .getByRole("radio", { checked: true })
    .count()
    .catch(() => 0);
  if (checkedByRole > 0) return true;
  const checkedByInput = await group
    .locator('input[type="radio"]:checked')
    .count()
    .catch(() => 0);
  if (checkedByInput > 0) return true;
  const checkedByAria = await group
    .locator('[role="radio"][aria-checked="true"]')
    .count()
    .catch(() => 0);
  return checkedByAria > 0;
};

async function selectRoleRadioOption({
  page,
  group,
  radios,
  pick,
  runState,
  pageNo,
  groupId,
}) {
  const strategies = getRadioStrategyOrder();
  for (const strategy of strategies) {
    try {
      if (strategy === "check") {
        await radios.nth(pick).check({ force: true, timeout: 1500 });
      } else if (strategy === "click") {
        await radios.nth(pick).click({ force: true, timeout: 1500 });
      } else if (strategy === "native") {
        const nativeInputs = group.locator('input[type="radio"]');
        const nativeCount = await nativeInputs.count().catch(() => 0);
        if (nativeCount <= pick) {
          throw new Error("native_input_missing");
        }
        await nativeInputs
          .nth(pick)
          .check({ force: true, timeout: 1500 })
          .catch(async () => {
            await nativeInputs.nth(pick).click({ force: true, timeout: 1500 });
          });
      } else if (strategy === "label") {
        let clicked = false;
        const labels = group.locator("label");
        const labelCount = await labels.count().catch(() => 0);
        if (labelCount > pick) {
          await labels.nth(pick).click({ force: true, timeout: 1500 });
          clicked = true;
        }
        if (!clicked) {
          const roleOptions = group.locator('[role="radio"]');
          const roleCount = await roleOptions.count().catch(() => 0);
          if (roleCount > pick) {
            await roleOptions.nth(pick).click({ force: true, timeout: 1500 });
            clicked = true;
          }
        }
        if (!clicked) {
          throw new Error("label_target_missing");
        }
      }
      await page.waitForTimeout(80);
      if (await hasCheckedInRoleGroup(group)) {
        return { ok: true, strategy };
      }
      appendSelectionError(runState, {
        page: pageNo,
        group: groupId,
        strategy,
        error: "not_checked_after_strategy",
      });
    } catch (error) {
      appendSelectionError(runState, {
        page: pageNo,
        group: groupId,
        strategy,
        error: String(error?.message ?? error ?? "unknown_selection_error"),
      });
    }
  }
  return { ok: false, strategy: null };
}

async function collectPythiaChoiceGroups(page) {
  const groups = await page
    .evaluate(() => {
      const normalize = (value) =>
        String(value ?? "")
          .replace(/\s+/g, " ")
          .trim();
      const isVisible = (node) => {
        if (!node) return false;
        const style = window.getComputedStyle(node);
        const rect = node.getBoundingClientRect();
        return (
          style.display !== "none" &&
          style.visibility !== "hidden" &&
          rect.width > 0 &&
          rect.height > 0
        );
      };

      return Array.from(
        document.querySelectorAll(
          "fieldset.question:not(.question_type_tableScale), .page__component fieldset.question:not(.question_type_tableScale)",
        ),
      )
        .map((fieldset, index) => {
          const options = Array.from(fieldset.querySelectorAll(".choice__option"));
          if (!options.length) return null;

          const firstInput = options[0]?.querySelector('input[type="radio"][name]');
          const name = normalize(firstInput?.getAttribute("name"));
          const questionText = normalize(
            fieldset.querySelector(".question__label .markdown, .question__label")
              ?.textContent,
          );
          const optionMeta = options.map((option, optionIndex) => {
            const input = option.querySelector('input[type="radio"]');
            const textCandidates = [
              option.querySelector(".choice__option-text .markdown")?.textContent,
              option.querySelector(".choice__option-text")?.textContent,
              option.querySelector("label .markdown")?.textContent,
              option.querySelector("label")?.textContent,
              option.querySelector("[role='radio']")?.textContent,
              option.getAttribute("aria-label"),
              input?.getAttribute("aria-label"),
              option.textContent,
            ];
            const text = textCandidates
              .map((candidate) => normalize(candidate))
              .find((candidate) => candidate.length > 0);
            return {
              index: optionIndex,
              label: text || `option_${optionIndex}`,
              value: input?.getAttribute("value") ?? "",
              input_id: input?.getAttribute("id") ?? "",
              testid:
                input?.getAttribute("data-testid") ??
                option.getAttribute("data-testid") ??
                "",
            };
          });
          const labels = optionMeta.map((meta) => meta.label);
          const checked = options.some((option) => {
            const input = option.querySelector('input[type="radio"]');
            if (input?.checked) return true;
            return option.classList.contains("choice__option_checked");
          });
          const visible = isVisible(fieldset) || options.some((option) => isVisible(option));
          return {
            index,
            question_id: normalize(fieldset.getAttribute("id")),
            name,
            questionText,
            labels,
            option_meta: optionMeta,
            checked,
            visible,
          };
        })
        .filter(Boolean);
    })
    .catch(() => []);
  return Array.isArray(groups) ? groups : [];
}

async function selectPythiaChoiceOption({
  page,
  groupName,
  pick,
  runState,
  pageNo,
  groupId,
}) {
  const strategies = ["option_container", "label", "control"];
  const escapedName = String(groupName).replace(/\\/g, "\\\\").replace(/"/g, '\\"');
  const getInputs = () => page.locator(`input[type="radio"][name="${escapedName}"]`);

  const verifySelection = async () => {
    const result = await page
      .evaluate((name) => {
        const inputs = Array.from(document.querySelectorAll('input[type="radio"]')).filter(
          (input) => (input.getAttribute("name") || "") === name,
        );
        if (!inputs.length) return false;
        return inputs.some((input) => {
          if (input.checked) return true;
          const option = input.closest(".choice__option");
          return Boolean(option?.classList.contains("choice__option_checked"));
        });
      }, groupName)
      .catch(() => false);
    return Boolean(result);
  };

  for (const strategy of strategies) {
    try {
      const inputs = getInputs();
      const inputCount = await inputs.count().catch(() => 0);
      if (inputCount <= pick) {
        appendSelectionError(runState, {
          page: pageNo,
          group: groupId,
          strategy: `pythia_${strategy}`,
          error: "target_not_found",
        });
        continue;
      }

      const targetInput = inputs.nth(pick);
      const optionContainer = targetInput
        .locator("xpath=ancestor::*[contains(@class,'choice__option')][1]")
        .first();
      if (strategy === "option_container") {
        if (await optionContainer.count().catch(() => 0)) {
          await optionContainer.click({ force: true, timeout: 1500 });
        } else {
          await targetInput.click({ force: true, timeout: 1500 });
        }
      } else if (strategy === "label") {
        const inputId = await targetInput.getAttribute("id").catch(() => null);
        if (inputId) {
          const label = page.locator(`label[for="${inputId}"]`).first();
          if (await label.count().catch(() => 0)) {
            await label.click({ force: true, timeout: 1500 });
          } else if (await optionContainer.count().catch(() => 0)) {
            await optionContainer.click({ force: true, timeout: 1500 });
          } else {
            await targetInput.click({ force: true, timeout: 1500 });
          }
        } else {
          await targetInput.click({ force: true, timeout: 1500 });
        }
      } else if (await optionContainer.count().catch(() => 0)) {
        const control = optionContainer.locator(".choice__option-control").first();
        if (await control.count().catch(() => 0)) {
          await control.click({ force: true, timeout: 1500 });
        } else {
          await optionContainer.click({ force: true, timeout: 1500 });
        }
      } else {
        await targetInput.click({ force: true, timeout: 1500 });
      }
      await page.waitForTimeout(120);
      if (await verifySelection()) {
        return { ok: true, strategy };
      }

      appendSelectionError(runState, {
        page: pageNo,
        group: groupId,
        strategy: `pythia_${strategy}`,
        error: "not_checked_after_strategy",
      });
    } catch (error) {
      appendSelectionError(runState, {
        page: pageNo,
        group: groupId,
        strategy: `pythia_${strategy}`,
        error: String(error?.message ?? error ?? "pythia_selection_error"),
      });
    }
  }
  return { ok: false, strategy: null };
}

async function collectPythiaTableScaleRows(page) {
  const rows = await page
    .evaluate(() => {
      const normalize = (value) =>
        String(value ?? "")
          .replace(/\s+/g, " ")
          .trim();
      const isVisible = (node) => {
        if (!node) return false;
        const style = window.getComputedStyle(node);
        const rect = node.getBoundingClientRect();
        return (
          style.display !== "none" &&
          style.visibility !== "hidden" &&
          rect.width > 0 &&
          rect.height > 0
        );
      };

      const result = [];
      const tables = Array.from(
        document.querySelectorAll(
          "fieldset.question.question_type_tableScale, fieldset.question_type_tableScale",
        ),
      ).filter((field) => isVisible(field));

      for (const table of tables) {
        const tableQuestionId = normalize(table.getAttribute("id"));
        const tableQuestionText = normalize(
          table.querySelector(".question__label .markdown, .question__label")?.textContent,
        );
        const rowNodes = Array.from(table.querySelectorAll(".question-table-scale__row"));
        rowNodes.forEach((rowNode, rowIndex) => {
          const rowText = normalize(
            rowNode.querySelector("[role='rowheader'] .markdown, [role='rowheader']")
              ?.textContent,
          );
          const name =
            rowNode.querySelector('input[type="radio"][name]')?.getAttribute("name") ?? "";
          const scoreNodes = Array.from(
            rowNode.querySelectorAll(".scale__option-with-checkbox[data-value]"),
          );
          const availableScores = scoreNodes
            .map((node) => Number(node.getAttribute("data-value")))
            .filter((value) => Number.isFinite(value));
          const checkedInput = rowNode.querySelector('input[type="radio"]:checked');
          const checkedContainer = checkedInput?.closest(".scale__option-with-checkbox");
          const checkedScoreRaw = checkedContainer?.getAttribute("data-value");
          const checkedScore = Number(checkedScoreRaw);
          const noOpinionChecked =
            (checkedInput?.id || "").includes("no-opinion-answer") ||
            (checkedInput?.dataset?.testid || "").includes("no-opinion-answer");

          result.push({
            table_question_id: tableQuestionId,
            table_question_text: tableQuestionText,
            row_index: rowIndex,
            row_text: rowText,
            group_name: normalize(name),
            available_scores: availableScores,
            checked_score: Number.isFinite(checkedScore) ? checkedScore : null,
            no_opinion_checked: Boolean(noOpinionChecked),
          });
        });
      }
      return result;
    })
    .catch(() => []);
  return Array.isArray(rows) ? rows : [];
}

async function selectPythiaTableScaleRowScore({
  page,
  groupName,
  score,
  pageNo,
  runState,
  groupId,
}) {
  const strategies = ["option_container", "label", "input_click"];

  const verify = async () =>
    page
      .evaluate(
        ({ name, score }) => {
          const inputs = Array.from(document.querySelectorAll('input[type="radio"]')).filter(
            (input) => (input.getAttribute("name") || "") === name,
          );
          if (!inputs.length) return false;
          const checked = inputs.find((input) => input.checked);
          if (!checked) return false;
          const container = checked.closest(".scale__option-with-checkbox");
          const value = Number(container?.getAttribute("data-value"));
          return Number.isFinite(value) && value === Number(score);
        },
        { name: groupName, score },
      )
      .catch(() => false);

  for (const strategy of strategies) {
    try {
      const actionResult = await page.evaluate(
        ({ name, score, strategy }) => {
          const inputs = Array.from(document.querySelectorAll('input[type="radio"]')).filter(
            (input) => (input.getAttribute("name") || "") === name,
          );
          if (!inputs.length) return { ok: false, error: "row_inputs_not_found" };

          let target = null;
          for (const input of inputs) {
            const container = input.closest(".scale__option-with-checkbox");
            const value = Number(container?.getAttribute("data-value"));
            if (Number.isFinite(value) && value === Number(score)) {
              target = input;
              break;
            }
          }
          if (!target) return { ok: false, error: "score_input_not_found" };

          const clickNode = (node) => {
            if (!node) return false;
            try {
              node.dispatchEvent(
                new MouseEvent("click", {
                  bubbles: true,
                  cancelable: true,
                  view: window,
                }),
              );
            } catch {
              // noop
            }
            try {
              node.click();
              return true;
            } catch {
              return false;
            }
          };

          const option = target.closest(".choice__option");
          const label = option?.querySelector("label") || document.querySelector(`label[for="${target.id}"]`);
          if (strategy === "option_container") {
            if (!clickNode(option) && !clickNode(label) && !clickNode(target)) {
              return { ok: false, error: "option_click_failed" };
            }
          } else if (strategy === "label") {
            if (!clickNode(label) && !clickNode(option) && !clickNode(target)) {
              return { ok: false, error: "label_click_failed" };
            }
          } else if (!clickNode(target) && !clickNode(option) && !clickNode(label)) {
            return { ok: false, error: "input_click_failed" };
          }
          return { ok: true };
        },
        { name: groupName, score, strategy },
      );
      if (!actionResult?.ok) {
        appendSelectionError(runState, {
          page: pageNo,
          group: groupId,
          strategy: `table_${strategy}`,
          error: String(actionResult?.error ?? "table_action_failed"),
        });
        continue;
      }

      await page.waitForTimeout(80);
      if (await verify()) {
        return { ok: true, strategy };
      }
      appendSelectionError(runState, {
        page: pageNo,
        group: groupId,
        strategy: `table_${strategy}`,
        error: "not_checked_after_strategy",
      });
    } catch (error) {
      appendSelectionError(runState, {
        page: pageNo,
        group: groupId,
        strategy: `table_${strategy}`,
        error: String(error?.message ?? error ?? "table_scale_selection_error"),
      });
    }
  }
  return { ok: false, strategy: null };
}

async function collectDomRadioGroups(page, skipNames = new Set()) {
  const skip = Array.from(skipNames ?? []);
  const groups = await page
    .evaluate((skipNames) => {
      const normalize = (value) =>
        String(value ?? "")
          .replace(/\s+/g, " ")
          .trim();
      const skipSet = new Set(skipNames ?? []);
      const map = new Map();
      const inputs = Array.from(document.querySelectorAll('input[type="radio"][name]'));

      const isVisible = (node) => {
        if (!node) return false;
        const style = window.getComputedStyle(node);
        const rect = node.getBoundingClientRect();
        return (
          style.display !== "none" &&
          style.visibility !== "hidden" &&
          rect.width > 0 &&
          rect.height > 0
        );
      };

      const getQuestionText = (input) => {
        const pythiaFieldset = input.closest("fieldset.question");
        if (pythiaFieldset) {
          const pythiaLabel = pythiaFieldset.querySelector(
            ".question__label .markdown, .question__label",
          );
          const text = normalize(pythiaLabel?.textContent);
          if (text) return text;
        }

        const selectors = "h1,h2,h3,legend,.QuestionLabel,[id$='-label']";
        let current = input.closest("fieldset,.Question,.QuestionMarkup,section,article,div");
        for (let i = 0; i < 10 && current; i += 1) {
          const heading = current.querySelector(selectors);
          const headingText = normalize(heading?.textContent);
          if (headingText) return headingText;
          current = current.parentElement;
        }
        return "";
      };

      inputs.forEach((input, index) => {
        const name = normalize(input.getAttribute("name"));
        if (!name || skipSet.has(name)) return;
        if (!map.has(name)) {
          const fieldset = input.closest("fieldset[id]");
          const questionId = normalize(fieldset?.getAttribute("id"));
          map.set(name, {
            name,
            question_id: questionId,
            firstIndex: index,
            questionText: "",
            labels: [],
            checked: false,
            visible: false,
          });
        }
        const group = map.get(name);
        if (!group.questionText) {
          group.questionText = getQuestionText(input);
        }
        const labelNode = input.closest("label");
        const roleNode = input.closest("[role='radio']");
        const optionText = normalize(
          input.getAttribute("aria-label") ||
            labelNode?.textContent ||
            roleNode?.textContent ||
            input.value ||
            `option_${group.labels.length}`,
        );
        group.labels.push(optionText || `option_${group.labels.length}`);
        if (input.checked) {
          group.checked = true;
        }
        if (roleNode?.getAttribute("aria-checked") === "true") {
          group.checked = true;
        }
        const pythiaOption = input.closest(".choice__option");
        const candidate = pythiaOption || labelNode || roleNode || input;
        if (isVisible(candidate)) {
          group.visible = true;
        }
      });

      return Array.from(map.values())
        .sort((a, b) => a.firstIndex - b.firstIndex)
        .map((item) => ({
          name: item.name,
          question_id: item.question_id,
          questionText: item.questionText,
          labels: item.labels,
          checked: item.checked,
          visible: item.visible,
        }));
    }, skip)
    .catch(() => []);
  return Array.isArray(groups) ? groups : [];
}

async function selectDomRadioOption({
  page,
  groupName,
  pick,
  runState,
  pageNo,
  groupId,
}) {
  const strategies = getRadioStrategyOrder();
  for (const strategy of strategies) {
    try {
      const result = await page.evaluate(
        ({ name, pickIndex, strategy }) => {
          const inputs = Array.from(document.querySelectorAll('input[type="radio"]')).filter(
            (input) => (input.getAttribute("name") || "") === name,
          );
          const target = inputs[pickIndex];
          if (!target) {
            return { ok: false, error: "target_not_found" };
          }

          const isChecked = () =>
            inputs.some((input) => input.checked) ||
            inputs.some((input) => {
              const role = input.closest("[role='radio']");
              return role?.getAttribute("aria-checked") === "true";
            });

          const emitNativeEvents = (node) => {
            node.dispatchEvent(new Event("input", { bubbles: true }));
            node.dispatchEvent(new Event("change", { bubbles: true }));
          };

          const clickNode = (node) => {
            if (!node) return;
            try {
              node.dispatchEvent(
                new MouseEvent("click", {
                  bubbles: true,
                  cancelable: true,
                  view: window,
                }),
              );
            } catch {
              // noop
            }
            try {
              node.click();
            } catch {
              // noop
            }
          };

          if (strategy === "check") {
            target.checked = true;
            emitNativeEvents(target);
          } else if (strategy === "click") {
            clickNode(target);
          } else if (strategy === "native") {
            target.focus();
            target.checked = true;
            emitNativeEvents(target);
            clickNode(target);
          } else if (strategy === "label") {
            const labelNode = target.closest("label");
            const roleNode = target.closest("[role='radio']");
            clickNode(labelNode || roleNode || target.parentElement || target);
            target.checked = true;
            emitNativeEvents(target);
          }

          return { ok: isChecked(), error: "not_checked_after_strategy" };
        },
        {
          name: groupName,
          pickIndex: pick,
          strategy,
        },
      );

      if (result?.ok) {
        await page.waitForTimeout(80);
        return { ok: true, strategy };
      }

      appendSelectionError(runState, {
        page: pageNo,
        group: groupId,
        strategy,
        error: String(result?.error ?? "not_checked_after_strategy"),
      });
    } catch (error) {
      appendSelectionError(runState, {
        page: pageNo,
        group: groupId,
        strategy,
        error: String(error?.message ?? error ?? "dom_selection_error"),
      });
    }
  }
  return { ok: false, strategy: null };
}

async function detectValidationErrors(page) {
  const errors = [];
  const visibleErrorMeta = await page
    .evaluate(() => {
      const isVisible = (node) => {
        if (!node) return false;
        const style = window.getComputedStyle(node);
        const rect = node.getBoundingClientRect();
        return (
          style.display !== "none" &&
          style.visibility !== "hidden" &&
          rect.width > 0 &&
          rect.height > 0
        );
      };
      const requiredBlocks = Array.from(
        document.querySelectorAll(
          "fieldset.question .question__content_error_yes, .question__content.question__content_error_yes",
        ),
      ).filter((node) => isVisible(node));
      const explicitErrorNodes = Array.from(
        document.querySelectorAll(
          "fieldset.question .QuestionLayout-ErrorContainer, fieldset.question .question__error, fieldset.question [aria-live='polite']",
        ),
      ).filter((node) => isVisible(node));

      return {
        required_error_visible: requiredBlocks.length > 0,
        explicit_error_visible: explicitErrorNodes.some((node) => {
          const text = String(node.textContent ?? "")
            .replace(/\s+/g, " ")
            .trim();
          return text.length > 0;
        }),
      };
    })
    .catch(() => ({
      required_error_visible: false,
      explicit_error_visible: false,
    }));

  if (visibleErrorMeta.required_error_visible) {
    errors.push("required_unanswered");
  }
  if (visibleErrorMeta.explicit_error_visible) {
    errors.push("fields_invalid");
  }

  if (!VALIDATION_VISIBLE_ONLY || !errors.length) {
    const bodyText = (await page.locator("body").innerText().catch(() => ""))
      .replace(/\s+/g, " ")
      .trim();
    if (/не оставляйте этот вопрос без ответа/i.test(bodyText)) {
      if (!errors.includes("required_unanswered")) {
        errors.push("required_unanswered");
      }
    }
    if (/не все поля заполнены верно/i.test(bodyText)) {
      if (!errors.includes("fields_invalid")) {
        errors.push("fields_invalid");
      }
    }
  }
  return errors;
}

async function fillVisibleFields(page, runState) {
  const pageNo = pageNoFromUrl(page.url());
  if (!runState.pages_visited.includes(pageNo)) {
    runState.pages_visited.push(pageNo);
  }
  const formEngine = await resolveFormEngine(page);
  runState.form_engine = formEngine;
  const pagePosition = await detectPagePosition(page);
  runState.last_page_position = pagePosition;
  const pageStats = {
    page: pageNo,
    engine: formEngine,
    question_ids: pagePosition.question_ids,
    schema_question_ids_on_page: [],
    is_first_page: pagePosition.is_first_page,
    is_last_page: pagePosition.is_last_page,
    pythia_groups_found: 0,
    pythia_groups_selected: 0,
    table_scale_rows_found: 0,
    table_scale_rows_selected: 0,
    table_scale_rows_unresolved: 0,
    table_scale_unmapped_rows: 0,
    role_groups_found: 0,
    dom_groups_found: 0,
    groups_found: 0,
    groups_selected: 0,
    validation_errors: 0,
    selection_unresolved: false,
    required_unanswered_groups: 0,
  };

  const pushAnswer = (entry) => {
    const key = JSON.stringify({
      page: entry.page,
      type: entry.type,
      value: entry.value,
      source: entry.source ?? null,
      question_id: entry.question_id ?? null,
      row_index: entry.row_index ?? null,
      group_name: entry.group_name ?? null,
    });
    if (runState.answer_keys.has(key)) return;
    runState.answer_keys.add(key);
    runState.answers_summary.push(entry);
  };

  const processedNames = new Set();
  const schemaQuestionIdsSet = new Set(
    Array.isArray(pagePosition.question_ids) ? pagePosition.question_ids : [],
  );
  const criticalQuestionIds = new Set(["q1", "q5", "q7", "q8"]);
  const ensureCriticalStateValue = async (questionId) => {
    if (!criticalQuestionIds.has(String(questionId ?? "").toLowerCase())) return true;
    if (await hasStateQuestionValue(page, runState, questionId)) return true;
    await page.waitForTimeout(80);
    return hasStateQuestionValue(page, runState, questionId);
  };

  const applyLikertTableRowAnswer = (decision, row) => {
    runState.selection_stats.rule_based += 1;
    if (decision.source === "rule_based_likert_fallback_score") {
      runState.likert_no_target_fallback_count += 1;
    }
    if (decision.source === "rule_based_likert_unmapped") {
      runState.likert_unmapped_rows_count += 1;
      pageStats.table_scale_unmapped_rows += 1;
    }
    pushAnswer({
      page: pageNo,
      type: "likert_table_row",
      question_id: row.table_question_id,
      row_index: row.row_index,
      row_text: row.row_text,
      value: String(decision.score),
      rule_id: decision.rule_id,
      source: decision.source,
      group_name: row.group_name,
    });
    runState.likert_answers_map.set(decision.progress_key, {
      rule_id: decision.rule_id,
      progress_key: decision.progress_key,
      segment_key: decision.segment_key,
      target_mean: decision.target_mean,
      value: decision.score,
    });
  };

  if (formEngine === "pythia") {
    const tableRows = await collectPythiaTableScaleRows(page);
    pageStats.table_scale_rows_found = tableRows.length;
    pageStats.groups_found += tableRows.length;
    const tableRowsToSelect = [];
    const tableResolvedScores = [];

    for (const row of tableRows) {
      if (row.group_name) {
        processedNames.add(row.group_name);
      }
      if (row.table_question_id) {
        schemaQuestionIdsSet.add(row.table_question_id.toLowerCase());
      }
      const availableScores = Array.isArray(row.available_scores)
        ? row.available_scores.filter((value) => Number.isFinite(value) && value >= 1 && value <= 5)
        : [];
      if (!row.group_name || !availableScores.length) {
        pageStats.table_scale_rows_unresolved += 1;
        continue;
      }

      const scoreAlreadyValid =
        Number.isFinite(row.checked_score) &&
        row.checked_score >= 1 &&
        row.checked_score <= 5 &&
        !(TABLE_SCALE_NO_OPINION_POLICY === "never" && row.no_opinion_checked);
      if (scoreAlreadyValid) {
        tableResolvedScores.push(clampLikertScore(row.checked_score));
        continue;
      }
      tableRowsToSelect.push({ row, availableScores });
    }

    let remainingTableRowsToSelect = tableRowsToSelect.length;
    for (const item of tableRowsToSelect) {
      const { row, availableScores } = item;
      const tableQuestionId = String(row.table_question_id ?? "").toLowerCase();
      const decision = decideTableScaleScore({
        tableQuestionId,
        rowIndex: row.row_index,
        rowText: row.row_text,
        runState,
      });
      const desiredScore = decision.score;
      let scoreToSet = desiredScore;
      if (!availableScores.includes(scoreToSet)) {
        scoreToSet = availableScores.reduce((best, candidate) => {
          if (Math.abs(candidate - desiredScore) < Math.abs(best - desiredScore)) {
            return candidate;
          }
          return best;
        }, availableScores[0]);
      }

      const nonOneScores = availableScores.filter((value) => value > 1);
      const isLastTableRow = remainingTableRowsToSelect === 1;
      const allResolvedOnes =
        tableResolvedScores.length > 0 && tableResolvedScores.every((value) => value === 1);
      if (scoreToSet === 1 && isLastTableRow && allResolvedOnes && nonOneScores.length) {
        const preferredNonOne = Number.isFinite(Number(decision.target_mean))
          ? Math.max(2, Math.min(5, Math.round(Number(decision.target_mean))))
          : Math.max(2, LIKERT_NO_TARGET_SCORE);
        scoreToSet = pickScoreWithRandomness(
          nonOneScores,
          preferredNonOne,
          LIKERT_FALLBACK_RANDOMNESS,
        );
      }

      remainingTableRowsToSelect -= 1;
      const selection = await selectPythiaTableScaleRowScore({
        page,
        groupName: row.group_name,
        score: scoreToSet,
        pageNo,
        runState,
        groupId: `table:${row.table_question_id}:${row.row_index}`,
      });
      if (!selection.ok) {
        pageStats.table_scale_rows_unresolved += 1;
        continue;
      }
      tableResolvedScores.push(clampLikertScore(scoreToSet));
      pageStats.table_scale_rows_selected += 1;
      pageStats.groups_selected += 1;
      applyLikertTableRowAnswer(
        {
          ...decision,
          score: scoreToSet,
        },
        row,
      );
    }

    const pythiaGroups = await collectPythiaChoiceGroups(page);
    const visiblePythiaGroups = pythiaGroups.filter((group) => group.visible);
    pageStats.pythia_groups_found = visiblePythiaGroups.length;
    pageStats.groups_found += visiblePythiaGroups.length;

    for (let pi = 0; pi < visiblePythiaGroups.length; pi += 1) {
      const pythiaGroup = visiblePythiaGroups[pi];
      if (pythiaGroup.name) {
        processedNames.add(pythiaGroup.name);
      }
      if (!pythiaGroup.name) {
        appendSelectionError(runState, {
          page: pageNo,
          group: `pythia:${pi}`,
          strategy: "collect",
          error: "missing_group_name",
        });
        continue;
      }
      if (pythiaGroup.checked) continue;
      const labels = Array.isArray(pythiaGroup.labels) ? pythiaGroup.labels : [];
      if (!labels.length) continue;
      const resolvedQuestionId = resolveQuestionIdFromSchema(
        runState,
        pythiaGroup.questionText ?? "",
        pythiaGroup.question_id ?? "",
      );
      if (resolvedQuestionId) {
        schemaQuestionIdsSet.add(resolvedQuestionId.toLowerCase());
      }
      const decision = decideRadioSelection({
        pageNo,
        questionText: pythiaGroup.questionText ?? "",
        labels,
        runState,
        questionId: resolvedQuestionId,
        optionMeta: pythiaGroup.option_meta ?? [],
      });
      const selection = await selectPythiaChoiceOption({
        page,
        groupName: pythiaGroup.name,
        pick: decision.pick,
        runState,
        pageNo,
        groupId: `pythia:${pythiaGroup.name || pi}:${pi}`,
      });
      if (!selection.ok) continue;
      if (!(await ensureCriticalStateValue(resolvedQuestionId))) {
        appendSelectionError(runState, {
          page: pageNo,
          group: `pythia:${pythiaGroup.name || pi}:${pi}`,
          strategy: "state_verify",
          error: "critical_value_missing_in_state",
        });
        continue;
      }

      pageStats.pythia_groups_selected += 1;
      pageStats.groups_selected += 1;
      applySelectedRadioValue({
        runState,
        pageNo,
        selectedValue: decision.selectedValue,
        selectionSource: decision.selectionSource,
        answerSource: decision.answerSource,
        likertMeta: decision.likertMeta,
        pushAnswer,
      });
    }
  }

  const groups = page.getByRole("radiogroup");
  const groupCount = await groups.count().catch(() => 0);
  for (let gi = 0; gi < groupCount; gi += 1) {
    const group = groups.nth(gi);
    try {
      if (!(await group.isVisible({ timeout: 100 }))) continue;
    } catch {
      continue;
    }
    pageStats.role_groups_found += 1;
    pageStats.groups_found += 1;

    const radios = group.getByRole("radio");
    const radioCount = await radios.count().catch(() => 0);
    if (!radioCount) continue;

    const labels = [];
    for (let ri = 0; ri < radioCount; ri += 1) {
      const label =
        (await radios
          .nth(ri)
          .evaluate((el) => {
            const normalize = (value) =>
              String(value ?? "")
                .replace(/\s+/g, " ")
                .trim();
            const aria = (el.getAttribute("aria-label") || "").trim();
            if (aria) return aria;
            const own = normalize(el.textContent || "");
            if (own) return own;
            const input =
              el instanceof HTMLInputElement && el.type === "radio"
                ? el
                : el.querySelector?.('input[type="radio"]');
            if (input) {
              const optionLabel = input.closest("label");
              const labelText = normalize(
                optionLabel?.querySelector(
                  ".g-control-label__text .markdown, .g-control-label__text, .OptionContent-Text, .choice__option-text, .markdown, p",
                )?.textContent ||
                  optionLabel?.textContent ||
                  "",
              );
              if (labelText) return labelText;
              const inputId = input.getAttribute("id");
              if (inputId) {
                const linkedLabel = Array.from(document.querySelectorAll("label[for]")).find(
                  (node) => node.getAttribute("for") === inputId,
                );
                const linkedText = normalize(linkedLabel?.textContent || "");
                if (linkedText) return linkedText;
              }
            }
            const container = el.closest("label, .choice__option, [role='radio']");
            const containerText = normalize(container?.textContent || "");
            if (containerText) return containerText;
            const parent = el.parentElement;
            return (
              (parent?.textContent || "")
                .replace(/\s+/g, " ")
                .trim() || ""
            );
          })
          .catch(() => "")) || `option_${ri}`;
      labels.push(label);
    }

    const nativeNames = await group
      .locator('input[type="radio"][name]')
      .evaluateAll((elements) =>
        elements
          .map((el) => el.getAttribute("name") || "")
          .filter((name) => Boolean(name)),
      )
      .catch(() => []);
    for (const name of nativeNames) {
      processedNames.add(name);
    }

    if (await hasCheckedInRoleGroup(group)) continue;

    const questionText = await extractQuestionTextFromGroup(group);
    const resolvedQuestionId = resolveQuestionIdFromSchema(runState, questionText, "");
    if (resolvedQuestionId) {
      schemaQuestionIdsSet.add(resolvedQuestionId.toLowerCase());
    }
    const decision = decideRadioSelection({
      pageNo,
      questionText,
      labels,
      runState,
      questionId: resolvedQuestionId,
    });

    const selection = await selectRoleRadioOption({
      page,
      group,
      radios,
      pick: decision.pick,
      runState,
      pageNo,
      groupId: `role:${gi}`,
    });
    if (!selection.ok) continue;
    if (!(await ensureCriticalStateValue(resolvedQuestionId))) {
      appendSelectionError(runState, {
        page: pageNo,
        group: `role:${gi}`,
        strategy: "state_verify",
        error: "critical_value_missing_in_state",
      });
      continue;
    }

    pageStats.groups_selected += 1;
    applySelectedRadioValue({
      runState,
      pageNo,
      selectedValue: decision.selectedValue,
      selectionSource: decision.selectionSource,
      answerSource: decision.answerSource,
      likertMeta: decision.likertMeta,
      pushAnswer,
    });
  }

  const domGroups = await collectDomRadioGroups(page, processedNames);
  const visibleDomGroups = domGroups.filter((group) => group.visible);
  pageStats.dom_groups_found = visibleDomGroups.length;
  pageStats.groups_found += visibleDomGroups.length;
  for (let di = 0; di < visibleDomGroups.length; di += 1) {
    const domGroup = visibleDomGroups[di];
    if (domGroup.checked) continue;
    const labels = Array.isArray(domGroup.labels) ? domGroup.labels : [];
    if (!labels.length) continue;

    const decision = decideRadioSelection({
      pageNo,
      questionText: domGroup.questionText ?? "",
      labels,
      runState,
      questionId: resolveQuestionIdFromSchema(
        runState,
        domGroup.questionText ?? "",
        domGroup.question_id ?? "",
      ),
    });
    const resolvedQuestionId = resolveQuestionIdFromSchema(
      runState,
      domGroup.questionText ?? "",
      domGroup.question_id ?? "",
    );
    if (resolvedQuestionId) {
      schemaQuestionIdsSet.add(resolvedQuestionId.toLowerCase());
    }
    const selection = await selectDomRadioOption({
      page,
      groupName: domGroup.name,
      pick: decision.pick,
      runState,
      pageNo,
      groupId: `dom:${domGroup.name}:${di}`,
    });
    if (!selection.ok) continue;
    if (!(await ensureCriticalStateValue(resolvedQuestionId))) {
      appendSelectionError(runState, {
        page: pageNo,
        group: `dom:${domGroup.name}:${di}`,
        strategy: "state_verify",
        error: "critical_value_missing_in_state",
      });
      continue;
    }

    pageStats.groups_selected += 1;
    applySelectedRadioValue({
      runState,
      pageNo,
      selectedValue: decision.selectedValue,
      selectionSource: decision.selectionSource,
      answerSource: decision.answerSource,
      likertMeta: decision.likertMeta,
      pushAnswer,
    });
  }

  const numberInputs = page.locator(
    'input[type="number"], input[inputmode="numeric"], input[role="spinbutton"]',
  );
  const numberCount = await numberInputs.count().catch(() => 0);
  for (let i = 0; i < numberCount; i += 1) {
    const el = numberInputs.nth(i);
    try {
      if (!(await el.isVisible({ timeout: 100 }))) continue;
      const ageRuleRange = resolveAgeRange(runState);
      let value = String(STRICT_RULES_ONLY ? 30 : randInt(18, 70));
      let source = STRICT_RULES_ONLY ? "rule_based_fallback" : "random_fallback";
      if (ageRuleRange) {
        // Even in strict mode, age varies randomly within the rule-defined range.
        value = String(randInt(ageRuleRange.min, ageRuleRange.max));
        source = "rule_based_age";
        runState.selection_stats.rule_based += 1;
      } else {
        if (STRICT_RULES_ONLY) {
          runState.selection_stats.rule_based += 1;
        } else {
          runState.selection_stats.random += 1;
        }
      }
      await el.fill(value);
      pushAnswer({ page: pageNo, type: "number", value, source });
    } catch {
      // best effort
    }
  }

  const textInputs = page.locator('input[type="text"], textarea');
  const textCount = await textInputs.count().catch(() => 0);
  for (let i = 0; i < textCount; i += 1) {
    const el = textInputs.nth(i);
    try {
      if (!(await el.isVisible({ timeout: 100 }))) continue;
      const cur = await el.inputValue().catch(() => "");
      if (cur?.trim()) continue;
      const value = STRICT_RULES_ONLY
        ? `test-fixed-${runState.attempt_id}`
        : `test-${randInt(100, 999)}`;
      await el.fill(value);
      pushAnswer({ page: pageNo, type: "text", value });
    } catch {
      // best effort
    }
  }

  const checkboxes = page.locator('input[type="checkbox"]');
  const checkCount = await checkboxes.count().catch(() => 0);
  if (checkCount > 0) {
    let hasChecked = false;
    for (let i = 0; i < checkCount; i += 1) {
      if (await checkboxes.nth(i).isChecked().catch(() => false)) {
        hasChecked = true;
        break;
      }
    }
    if (!hasChecked) {
      const pick = STRICT_RULES_ONLY ? 0 : randInt(0, checkCount - 1);
      await checkboxes
        .nth(pick)
        .check({ force: true })
        .then(() =>
          pushAnswer({
            page: pageNo,
            type: "checkbox",
            value: `index_${pick}`,
          }),
        )
        .catch(() => {});
    }
  }

  const requiredUnansweredGroups = await page
    .evaluate((noOpinionPolicy) => {
      const isVisible = (node) => {
        if (!node) return false;
        const style = window.getComputedStyle(node);
        const rect = node.getBoundingClientRect();
        return (
          style.display !== "none" &&
          style.visibility !== "hidden" &&
          rect.width > 0 &&
          rect.height > 0
        );
      };

      const requiredFieldsets = Array.from(
        document.querySelectorAll(
          "fieldset.question[aria-required='true'], fieldset[aria-required='true']",
        ),
      ).filter((fieldset) => isVisible(fieldset));

      let unanswered = 0;
      for (const fieldset of requiredFieldsets) {
        const radios = Array.from(fieldset.querySelectorAll('input[type="radio"][name]'));
        if (!radios.length) continue;
        const names = Array.from(new Set(radios.map((input) => input.getAttribute("name"))))
          .filter(Boolean);
        for (const name of names) {
          const checked = fieldset.querySelector(`input[type="radio"][name="${name}"]:checked`);
          if (!checked) {
            unanswered += 1;
            continue;
          }
          if (noOpinionPolicy === "never") {
            const id = String(checked.id || "");
            const testid = String(checked.getAttribute("data-testid") || "");
            if (id.includes("no-opinion-answer") || testid.includes("no-opinion-answer")) {
              unanswered += 1;
            }
          }
        }
      }
      return unanswered;
    }, TABLE_SCALE_NO_OPINION_POLICY)
    .catch(() => 0);
  pageStats.required_unanswered_groups = sanitizeTargetValue(
    requiredUnansweredGroups,
    0,
  );
  pageStats.schema_question_ids_on_page = Array.from(schemaQuestionIdsSet).sort();

  runState.last_page_stats = pageStats;
  appendPageDebug(runState, pageStats);
  return pageStats;
}

async function ensurePageSatisfied(page, runState, retries = PAGE_VALIDATION_RETRIES) {
  const maxRetries = Math.max(0, sanitizeTargetValue(retries, PAGE_VALIDATION_RETRIES));
  let lastStats = null;
  let lastErrors = [];
  let unresolvedSelection = false;
  for (let attempt = 0; attempt <= maxRetries; attempt += 1) {
    lastStats = await fillVisibleFields(page, runState);
    lastErrors = await detectValidationErrors(page);
    if (lastStats) {
      lastStats.validation_errors = lastErrors.length;
      const looksUnresolved =
        sanitizeTargetValue(lastStats.groups_found, 0) > 0 &&
        (sanitizeTargetValue(lastStats.groups_selected, 0) === 0 ||
          sanitizeTargetValue(lastStats.table_scale_rows_unresolved, 0) > 0 ||
          sanitizeTargetValue(lastStats.required_unanswered_groups, 0) > 0) &&
        (lastErrors.includes("required_unanswered") ||
          sanitizeTargetValue(lastStats.table_scale_rows_unresolved, 0) > 0 ||
          sanitizeTargetValue(lastStats.required_unanswered_groups, 0) > 0);
      if (looksUnresolved) {
        lastStats.selection_unresolved = true;
        unresolvedSelection = true;
      }
    }
    runState.last_validation_errors = lastErrors;
    if (!lastErrors.length) {
      return {
        ok: true,
        errors: [],
        stats: lastStats,
        reason: null,
      };
    }
    runState.validation_error_seen = true;
    appendPageDebug(runState, {
      ...(lastStats ?? {}),
      phase: "validation_retry",
      retry: attempt + 1,
      validation_errors: lastErrors.length,
      validation_error_types: lastErrors.join(","),
    });
    if (attempt < maxRetries) {
      await page.waitForTimeout(250);
    }
  }
  return {
    ok: false,
    errors: lastErrors,
    stats: lastStats,
    reason: unresolvedSelection ? "selection_unresolved" : "validation_stuck",
  };
}

async function performAttempt(attemptId, browserType, runPlan = {}) {
  const start = new Date().toISOString();
  const targetCell = runPlan?.target_cell ?? null;
  const onSubmitClick =
    typeof runPlan?.on_submit_click === "function" ? runPlan.on_submit_click : null;
  const runState = {
    run_mode: runPlan?.run_mode ?? RUN_MODE,
    target_cell: targetCell,
    attempt_id: attemptId,
    branch_choice: targetCell
      ? "Да"
      : STRICT_RULES_ONLY
        ? BRANCH_POLICY === "always_no"
          ? "Нет"
          : "Да"
        : Math.random() < 0.5
          ? "Да"
          : "Нет",
    trip_type: targetCell?.trip_type ?? null,
    children_mode: targetCell?.children_mode ?? null,
    gender_choice: null,
    gender_counter: normalizeGenderCounter(runPlan?.gender_counter),
    gender_expected_total: sanitizeTargetValue(runPlan?.gender_expected_total, 0) || null,
    pages_visited: [],
    answers_summary: [],
    answer_keys: new Set(),
    likert_rules: Array.isArray(runPlan?.likert_rules) ? runPlan.likert_rules : [],
    likert_expected_total: sanitizeTargetValue(runPlan?.likert_expected_total, 0) || null,
    likert_segment_key:
      runPlan?.likert_segment_key ??
      makeLikertSegmentKeyFromCell(targetCell),
    likert_segment_targets:
      runPlan?.likert_segment_targets &&
      typeof runPlan.likert_segment_targets === "object"
        ? runPlan.likert_segment_targets
        : {},
    likert_segment_expected_total:
      runPlan?.likert_segment_expected_total &&
      typeof runPlan.likert_segment_expected_total === "object"
        ? runPlan.likert_segment_expected_total
        : {},
    likert_progress: normalizeLikertProgress(runPlan?.likert_progress),
    likert_answers_map: new Map(),
    selection_errors: [],
    validation_error_seen: false,
    last_validation_errors: [],
    last_page_position: null,
    form_engine: null,
    page_debug: [],
    last_page_stats: null,
    schema:
      runPlan?.schema && typeof runPlan.schema === "object"
        ? runPlan.schema
        : {
            source: "none",
            storage_key: null,
            pages_count: null,
            questions_count: null,
            question_type_histogram: {},
            pages: [],
            questions_by_id: {},
            question_text_to_id: {},
          },
    table_row_rule_map: new Map(),
    table_row_fallback_cursor: 0,
    likert_no_target_fallback_count: 0,
    likert_unmapped_rows_count: 0,
    submit_clicks_attempt: 0,
    selection_stats: {
      random: 0,
      rule_based: 0,
    },
  };
  const getSafeFinalUrl = () => {
    try {
      return page.url();
    } catch {
      return null;
    }
  };
  const makeAttemptResult = (status, overrides = {}) => ({
    attempt_id: attemptId,
    status,
    start_ts: start,
    end_ts: new Date().toISOString(),
    branch_choice: runState.branch_choice,
    pages_visited: runState.pages_visited,
    answers_summary: runState.answers_summary,
    selection_stats: {
      random: runState.selection_stats.random,
      rule_based: runState.selection_stats.rule_based,
    },
    scheme: classifyScheme(runState.selection_stats),
    success_marker: null,
    quota_cell_key: runState.target_cell?.key ?? null,
    gender_choice: runState.gender_choice,
    likert_answers: Array.from(runState.likert_answers_map.values()),
    selection_errors: runState.selection_errors,
    validation_error_seen: runState.validation_error_seen,
    validation_errors: runState.last_validation_errors,
    form_engine: runState.form_engine,
    page_position: runState.last_page_position,
    page_debug: runState.page_debug,
    likert_no_target_fallback_count: runState.likert_no_target_fallback_count,
    likert_unmapped_rows_count: runState.likert_unmapped_rows_count,
    submit_clicks_attempt: runState.submit_clicks_attempt,
    final_url: getSafeFinalUrl(),
    ...overrides,
  });

  const launchOptions = buildLaunchOptions();
  const browser = await browserType.launch(launchOptions);
  const context = await browser.newContext();
  const page = await context.newPage();

  try {
    await page.goto(`${FORM_URL}?reset=attempt-${Date.now()}`, {
      waitUntil: "domcontentloaded",
      timeout: 45_000,
    });
    await clickCookieBannerIfPresent(page);
    const startPageCheck = await ensureStartsFromFirstPage(
      page,
      context,
      attemptId,
      runState,
    );
    if (!startPageCheck.ok) {
      const startStatus =
        startPageCheck.reason === "load_error_screen"
          ? "load_error_screen"
          : "start_page_invalid";
      return makeAttemptResult(startStatus, {
        page_position: startPageCheck.position,
      });
    }

    for (let step = 0; step < 160; step += 1) {
      await page.waitForTimeout(250);
      await clickCookieBannerIfPresent(page);

      if (await detectLoadErrorScreen(page)) {
        return makeAttemptResult("load_error_screen");
      }

      if (await detectCaptcha(page)) {
        return makeAttemptResult("captcha_detected");
      }

      const ensureResult = await ensurePageSatisfied(page, runState);
      if (!ensureResult.ok) {
        return makeAttemptResult(
          ensureResult.reason === "selection_unresolved"
            ? "selection_unresolved"
            : "validation_stuck",
          {
          validation_errors: ensureResult.errors,
          },
        );
      }

      const submitBtn = page
        .getByRole("button", { name: /^(Отправить|Submit|Готово|Завершить опрос)$/i })
        .first();
      if (await submitBtn.count().catch(() => 0)) {
        if (await submitBtn.isVisible().catch(() => false)) {
          const clicked = await submitBtn
            .click({ timeout: 2_000 })
            .then(() => true)
            .catch(() => false);
          if (clicked) {
            runState.submit_clicks_attempt += 1;
            if (onSubmitClick) {
              await onSubmitClick({
                attempt_id: attemptId,
                click_ts: new Date().toISOString(),
                quota_cell_key: targetCell?.key ?? null,
              }).catch(() => {});
            }
          }
          for (let i = 0; i < 30; i += 1) {
            await page.waitForTimeout(500);
            if (await detectCaptcha(page)) {
              return makeAttemptResult("captcha_detected");
            }
            const successMarker = await detectSuccessMarker(page);
            if (successMarker) {
              return makeAttemptResult("success", {
                success_marker: successMarker,
              });
            }
            const submitEnsureResult = await ensurePageSatisfied(page, runState, 1);
            if (!submitEnsureResult.ok) {
              return makeAttemptResult(
                submitEnsureResult.reason === "selection_unresolved"
                  ? "selection_unresolved"
                  : "validation_stuck",
                {
                  validation_errors: submitEnsureResult.errors,
                },
              );
            }
          }
          if (clicked) {
            return makeAttemptResult("submit_unconfirmed");
          }
        }
      }

      const nextBtn = page
        .getByRole("button", { name: /^(Далее|Next)$/i })
        .first();
      if (await nextBtn.count().catch(() => 0)) {
        if (await nextBtn.isVisible().catch(() => false)) {
          await nextBtn.click({ timeout: 2_000 }).catch(() => {});
          await page.waitForTimeout(220);
          const postNextErrors = await detectValidationErrors(page);
          if (postNextErrors.length) {
            runState.validation_error_seen = true;
            runState.last_validation_errors = postNextErrors;
            appendPageDebug(runState, {
              page: pageNoFromUrl(page.url()),
              phase: "post_next_validation_error",
              validation_errors: postNextErrors.length,
              validation_error_types: postNextErrors.join(","),
            });
          }
          continue;
        }
      }

      const successMarker = await detectSuccessMarker(page);
      if (successMarker) {
        return makeAttemptResult("success", {
          success_marker: successMarker,
        });
      }
    }

    return makeAttemptResult("step_limit_reached");
  } catch (error) {
    return makeAttemptResult(classifyAttemptErrorStatus(error), {
      error_message: String(error?.message ?? error ?? ""),
    });
  } finally {
    await context.close().catch(() => {});
    await browser.close().catch(() => {});
  }
}

async function main() {
  const playwright = await loadPlaywright();
  const browserType = {
    chromium: playwright.chromium,
    firefox: playwright.firefox,
    webkit: playwright.webkit,
  }[BROWSER_NAME];

  if (!browserType) {
    throw new Error(
      `Unsupported BROWSER="${BROWSER_NAME}". Use chromium|firefox|webkit.`,
    );
  }
  if (SUCCESS_MODE !== "click_and_confirm") {
    throw new Error(
      `Unsupported SUCCESS_MODE="${SUCCESS_MODE}". Only click_and_confirm is supported.`,
    );
  }
  const schemaDiscovery = await discoverSurveySchema(browserType);

  const attempts = [];
  const successful_runs = [];
  const likertConfig = loadLikertTargets();
  const quotaTargetsMap = RUN_MODE === "quota" ? loadQuotaTargets() : null;
  const quotaCells = quotaTargetsMap ? buildQuotaCells(quotaTargetsMap) : [];
  const quotaStateFileExists = RUN_MODE === "quota" && fs.existsSync(QUOTA_STATE_PATH);
  const quotaState = quotaTargetsMap ? loadQuotaState(quotaTargetsMap) : null;
  const quotaTargetTotal = quotaTargetsMap
    ? Object.values(quotaTargetsMap).reduce(
        (sum, value) => sum + sanitizeTargetValue(value, 0),
        0,
      )
    : null;
  const likertSegmentExpectedFromQuota = quotaTargetsMap
    ? buildLikertSegmentExpectedTotalsFromQuota(quotaTargetsMap)
    : {};
  const likertSegmentExpectedTotal = {
    ...likertSegmentExpectedFromQuota,
    ...(likertConfig.segment_expected_total ?? {}),
  };
  const likertExpectedTotal =
    likertConfig.expected_total ??
    (RUN_MODE === "quota" ? quotaTargetTotal : TARGET_SUCCESSFUL_SUBMITS);
  let likertProgressMemory = {};
  const lockPath = `${QUOTA_STATE_PATH}.lock`;
  let lockAcquired = false;
  const releaseLock = () => {
    if (!lockAcquired) return;
    releaseStateLock(lockPath, process.pid);
    lockAcquired = false;
  };

  if (RUN_MODE === "quota" && quotaTargetsMap && quotaState) {
    quotaState.form_url = normalizeFormUrlForState(FORM_URL);
    quotaState.quota_targets_hash = computeQuotaTargetsHash(quotaTargetsMap);
    quotaState.success_mode = SUCCESS_MODE;
    quotaState.runner_version = RUNNER_VERSION;
  }

  if (
    RUN_MODE === "quota" &&
    quotaTargetsMap &&
    quotaState &&
    !RESET_QUOTA_STATE &&
    quotaStateFileExists
  ) {
    const compatibility = validateQuotaStateCompatibility(quotaState, quotaTargetsMap);
    if (!compatibility.ok) {
      const report = {
        config: {
          FORM_URL,
          RUN_MODE,
          QUOTA_STATE_PATH,
          QUOTA_TARGETS_PATH,
          SUCCESS_MODE,
        },
        summary: {
          successful_submits: sanitizeTargetValue(quotaState.success_total, 0),
          attempts_total: sanitizeTargetValue(quotaState.attempts_total, 0),
          run_exit_reason: "state_incompatible",
          state_compatibility: compatibility,
        },
        attempts: [],
        successful_runs: [],
      };
      fs.writeFileSync(OUT_PATH, `${JSON.stringify(report, null, 2)}\n`, "utf8");
      throw new Error(
        `State incompatible (${compatibility.reason}). Expected form/quota/success mode do not match current run.`,
      );
    }
  }

  if (RUN_MODE === "quota") {
    const lockResult = acquireStateLock(lockPath, {
      pid: process.pid,
      started_at: new Date().toISOString(),
      form_url: normalizeFormUrlForState(FORM_URL),
      quota_state_path: QUOTA_STATE_PATH,
      quota_targets_path: QUOTA_TARGETS_PATH || null,
      success_mode: SUCCESS_MODE,
      runner_version: RUNNER_VERSION,
    });
    if (!lockResult.acquired) {
      const report = {
        config: {
          FORM_URL,
          RUN_MODE,
          QUOTA_STATE_PATH,
          QUOTA_TARGETS_PATH,
          SUCCESS_MODE,
        },
        summary: {
          successful_submits: sanitizeTargetValue(quotaState?.success_total, 0),
          attempts_total: sanitizeTargetValue(quotaState?.attempts_total, 0),
          run_exit_reason: "lock_active",
          active_lock: lockResult.existing ?? null,
        },
        attempts: [],
        successful_runs: [],
      };
      fs.writeFileSync(OUT_PATH, `${JSON.stringify(report, null, 2)}\n`, "utf8");
      throw new Error("lock_active: another runner process is already working with this state");
    }
    lockAcquired = true;
  }

  if (
    RUN_MODE === "quota" &&
    quotaState &&
    (RESET_QUOTA_STATE || !fs.existsSync(QUOTA_STATE_PATH))
  ) {
    saveQuotaState(quotaState);
  }
  let attemptId =
    RUN_MODE === "quota" && quotaState
      ? sanitizeTargetValue(quotaState.last_attempt_id, 0)
      : 0;
  let runExitReason = "running";

  const composeReport = () => {
    const schemeCounterSuccessfulRuns = {
      random: 0,
      rule_based: 0,
      mixed: 0,
    };
    for (const run of successful_runs) {
      const scheme = run.scheme;
      if (scheme && scheme in schemeCounterSuccessfulRuns) {
        schemeCounterSuccessfulRuns[scheme] += 1;
      }
    }
    const successfulLikertTableAnswers = successful_runs.flatMap((run) =>
      Array.isArray(run?.answers_summary)
        ? run.answers_summary.filter((entry) => entry?.type === "likert_table_row")
        : [],
    );
    const likertTableRowsAnsweredTotal = successfulLikertTableAnswers.length;
    const likertTableRowsUnmappedTotal = successfulLikertTableAnswers.filter((entry) =>
      String(entry?.source ?? "").includes("unmapped"),
    ).length;
    const likertNoTargetFallbackCount = successfulLikertTableAnswers.filter(
      (entry) => String(entry?.source ?? "") === "rule_based_likert_fallback_score",
    ).length;

    const quotaSnapshotFinal =
      RUN_MODE === "quota" && quotaTargetsMap && quotaState
        ? buildQuotaRemaining(quotaTargetsMap, quotaState.completed)
        : null;
    const answerCountersFinal =
      RUN_MODE === "quota" && quotaState
        ? quotaState.answer_counters ??
          deriveAnswerCountersFromCompleted(quotaState.completed)
        : null;
    const genderCounterFinal =
      RUN_MODE === "quota" && quotaState
        ? normalizeGenderCounter(quotaState.gender_counter)
        : null;
    const answerCountersWithGenderFinal =
      answerCountersFinal && genderCounterFinal
        ? mergeAnswerCountersWithGender(answerCountersFinal, genderCounterFinal)
        : answerCountersFinal;
    const targetReached =
      RUN_MODE === "quota"
        ? Boolean(quotaSnapshotFinal && quotaSnapshotFinal.total_remaining === 0)
        : successful_runs.length === TARGET_SUCCESSFUL_SUBMITS;
    const likertProgressFinal =
      RUN_MODE === "quota" ? quotaState?.likert_progress : likertProgressMemory;
    const likertProgressSummary = buildLikertProgressSummary(likertProgressFinal);
    const attemptsTotalCurrent = attempts.length;
    const successfulSubmitsCurrent = successful_runs.length;
    const attemptsTotalCumulative =
      RUN_MODE === "quota"
        ? sanitizeTargetValue(quotaState?.attempts_total, attemptsTotalCurrent)
        : attemptsTotalCurrent;
    const successfulSubmitsCumulative =
      RUN_MODE === "quota"
        ? sanitizeTargetValue(quotaState?.success_total, successfulSubmitsCurrent)
        : successfulSubmitsCurrent;
    const submitClickTotalCumulative =
      RUN_MODE === "quota"
        ? sanitizeTargetValue(quotaState?.submit_click_total, 0)
        : null;
    const submitClickUnconfirmedTotalCumulative =
      RUN_MODE === "quota"
        ? sanitizeTargetValue(quotaState?.submit_click_unconfirmed_total, 0)
        : null;
    const attemptStatusCounterCurrent = buildAttemptStatusCounter(attempts);
    const attemptStatusCounterCumulative =
      RUN_MODE === "quota"
        ? normalizeAttemptStatusCounter(quotaState?.attempt_status_counter)
        : attemptStatusCounterCurrent;
    const progressPercent =
      quotaSnapshotFinal && quotaSnapshotFinal.total_target > 0
        ? Number(
            (
              (quotaSnapshotFinal.total_completed / quotaSnapshotFinal.total_target) *
              100
            ).toFixed(4),
          )
        : null;

    return {
      config: {
        FORM_URL,
        RUN_MODE,
        TARGET_SUCCESSFUL_SUBMITS,
        BRANCH_POLICY,
        CAPTCHA_POLICY,
        MAX_TOTAL_ATTEMPTS,
        HEADLESS,
        BROWSER_NAME,
        BROWSER_CHANNEL,
        BROWSER_EXECUTABLE_PATH,
        INCOGNITO,
        STRICT_RULES_ONLY,
        QUOTA_STATE_PATH,
        QUOTA_TARGETS_PATH,
        RESET_QUOTA_STATE,
        LIKERT_TARGETS_PATH,
        PAGE_VALIDATION_RETRIES,
        START_PAGE_RECOVERY_ATTEMPTS,
        START_PAGE_RECOVERY_MODE,
        FORM_DOM_MODE,
        VALIDATION_VISIBLE_ONLY,
        RADIO_FALLBACK_STRATEGY,
        LOG_PAGE_DEBUG,
        SCHEMA_DISCOVERY_MODE,
        SCHEMA_STORAGE_KEY,
        SELECTOR_STRATEGY,
        TABLE_SCALE_NO_OPINION_POLICY,
        LIKERT_NO_TARGET_SCORE,
        LIKERT_TARGET_RANDOMNESS,
        LIKERT_FALLBACK_RANDOMNESS,
        SUCCESS_MODE,
        PROGRESS_MODE,
        PROGRESS_EVERY_N,
        RUNNER_VERSION,
      },
      summary: {
        successful_submits: successfulSubmitsCumulative,
        attempts_total: attemptsTotalCumulative,
        attempt_status_counter: attemptStatusCounterCumulative,
        successful_submits_current_process: successfulSubmitsCurrent,
        attempts_total_current_process: attemptsTotalCurrent,
        successful_submits_cumulative: successfulSubmitsCumulative,
        attempts_total_cumulative: attemptsTotalCumulative,
        submit_click_total_cumulative: submitClickTotalCumulative,
        submit_click_unconfirmed_total_cumulative:
          submitClickUnconfirmedTotalCumulative,
        attempt_status_counter_cumulative: attemptStatusCounterCumulative,
        attempt_status_counter_current_process: attemptStatusCounterCurrent,
        success_mode: SUCCESS_MODE,
        progress_percent: progressPercent,
        target_reached: targetReached,
        run_exit_reason: runExitReason,
        scheme_counter_successful_runs: schemeCounterSuccessfulRuns,
        quota_target_total: quotaSnapshotFinal?.total_target ?? null,
        quota_completed_total: quotaSnapshotFinal?.total_completed ?? null,
        quota_remaining_total: quotaSnapshotFinal?.total_remaining ?? null,
        quota_completed: quotaState?.completed ?? null,
        quota_remaining: quotaSnapshotFinal?.remaining ?? null,
        quota_reached:
          RUN_MODE === "quota"
            ? Boolean(quotaSnapshotFinal && quotaSnapshotFinal.total_remaining === 0)
            : null,
        answer_counter_quota_successful_runs: answerCountersFinal,
        answer_counter_quota_successful_runs_with_gender:
          answerCountersWithGenderFinal,
        answer_counter_city_labels: answerCountersFinal
          ? toCityLabelCounters(answerCountersFinal.city)
          : null,
        gender_counter_successful_runs: genderCounterFinal,
        gender_balance_diff: genderCounterFinal
          ? Math.abs(genderCounterFinal.male - genderCounterFinal.female)
          : null,
        quota_state_last_updated: quotaState?.last_updated ?? null,
        likert_targets_count: likertConfig.rules.length,
        likert_expected_total: likertExpectedTotal,
        likert_progress: likertProgressSummary,
        likert_table_rows_answered_total: likertTableRowsAnsweredTotal,
        likert_table_rows_unmapped_total: likertTableRowsUnmappedTotal,
        likert_no_target_fallback_count: likertNoTargetFallbackCount,
        schema_discovery: {
          source: schemaDiscovery?.source ?? "none",
          storage_key: schemaDiscovery?.storage_key ?? null,
          pages_count: schemaDiscovery?.pages_count ?? null,
          questions_count: schemaDiscovery?.questions_count ?? null,
          question_type_histogram: schemaDiscovery?.question_type_histogram ?? {},
        },
      },
      attempts,
      successful_runs,
    };
  };

  const persistProgress = () => {
    if (RUN_MODE === "quota" && quotaState) {
      saveQuotaState(quotaState);
    }
    const snapshot = composeReport();
    fs.writeFileSync(OUT_PATH, `${JSON.stringify(snapshot, null, 2)}\n`, "utf8");
    return snapshot;
  };

  const logAttemptProgress = (result) => {
    const shouldLog = (() => {
      if (PROGRESS_MODE === "minimal") {
        return ["success", "captcha_detected", "network_changed", "submit_unconfirmed"].includes(
          String(result?.status ?? ""),
        );
      }
      if (PROGRESS_MODE === "every_n") {
        return (
          attemptId % PROGRESS_EVERY_N === 0 || String(result?.status ?? "") === "success"
        );
      }
      return true;
    })();
    if (!shouldLog) return;

    const lastPageDebug = Array.isArray(result?.page_debug)
      ? result.page_debug[result.page_debug.length - 1]
      : null;
    const pagePosition = result?.page_position ?? null;
    const questionIds = Array.isArray(pagePosition?.question_ids)
      ? pagePosition.question_ids.join(",")
      : "";
    const debugSuffix = lastPageDebug
      ? ` page=${lastPageDebug.page ?? "-"} engine=${result?.form_engine ?? lastPageDebug.engine ?? "-"} groups_found=${lastPageDebug.groups_found ?? 0} groups_selected=${lastPageDebug.groups_selected ?? 0} table_rows=${lastPageDebug.table_scale_rows_selected ?? 0}/${lastPageDebug.table_scale_rows_found ?? 0} table_unresolved=${lastPageDebug.table_scale_rows_unresolved ?? 0} validation_errors=${lastPageDebug.validation_errors ?? 0}`
      : ` engine=${result?.form_engine ?? "-"}`;
    const idsSuffix = questionIds ? ` page_ids=${questionIds}` : "";
    if (RUN_MODE === "quota" && quotaTargetsMap && quotaState) {
      const quotaSnapshot = buildQuotaRemaining(quotaTargetsMap, quotaState.completed);
      const cumulativeSuccess = sanitizeTargetValue(quotaState.success_total, 0);
      const submitClicks = sanitizeTargetValue(quotaState.submit_click_total, 0);
      const unconfirmedClicks = sanitizeTargetValue(
        quotaState.submit_click_unconfirmed_total,
        0,
      );
      const cellKey = result?.quota_cell_key ?? null;
      console.log(
        `attempt=${attemptId} status=${result.status} success=${cumulativeSuccess}/${quotaSnapshot.total_target} remaining=${quotaSnapshot.total_remaining} submit_clicks=${submitClicks} unconfirmed_clicks=${unconfirmedClicks} cell=${cellKey ?? "-"}${debugSuffix}${idsSuffix}`,
      );
      if (
        result.status === "success" &&
        Number.isFinite(Number(result?.quota_cell_before)) &&
        Number.isFinite(Number(result?.quota_cell_after))
      ) {
        console.log(
          `attempt=${attemptId} quota_progress cell=${cellKey ?? "-"} ${result.quota_cell_before}->${result.quota_cell_after}`,
        );
      }
      return;
    }
    console.log(
      `[attempt ${attemptId}] status=${result.status} successful=${successful_runs.length}/${TARGET_SUCCESSFUL_SUBMITS}${debugSuffix}${idsSuffix}`,
    );
  };

  const stopWithSignal = (signal) => {
    runExitReason = `stopped_by_${signal.toLowerCase()}`;
    try {
      persistProgress();
      console.log(`Progress snapshot saved: ${OUT_PATH}`);
    } catch (error) {
      console.error("Failed to persist progress snapshot", error);
    } finally {
      releaseLock();
    }
    process.exit(130);
  };

  const stopWithFatal = (kind, error) => {
    runExitReason = `fatal_${kind}`;
    try {
      if (error) {
        console.error(error);
      }
      persistProgress();
      console.log(`Progress snapshot saved: ${OUT_PATH}`);
    } catch (persistError) {
      console.error("Failed to persist progress snapshot", persistError);
    } finally {
      releaseLock();
    }
    process.exit(1);
  };

  process.once("SIGINT", () => stopWithSignal("SIGINT"));
  process.once("SIGTERM", () => stopWithSignal("SIGTERM"));
  process.once("uncaughtException", (error) => stopWithFatal("uncaught_exception", error));
  process.once("unhandledRejection", (reason) =>
    stopWithFatal("unhandled_rejection", reason),
  );

  try {
    const shouldContinue = () => {
      if (attemptId >= MAX_TOTAL_ATTEMPTS) return false;
      if (RUN_MODE !== "quota") {
        return successful_runs.length < TARGET_SUCCESSFUL_SUBMITS;
      }
      if (!quotaTargetsMap || !quotaState) return false;
      const quotaSnapshot = buildQuotaRemaining(quotaTargetsMap, quotaState.completed);
      return quotaSnapshot.total_remaining > 0;
    };

    while (shouldContinue()) {
      attemptId += 1;
      if (RUN_MODE === "quota" && quotaState) {
        quotaState.last_attempt_id = attemptId;
        saveQuotaState(quotaState);
      }
      let targetCell = null;
      if (RUN_MODE === "quota") {
        const quotaSnapshot = buildQuotaRemaining(quotaTargetsMap, quotaState.completed);
        targetCell = pickNextQuotaCell(quotaCells, quotaSnapshot.remaining);
        if (!targetCell) break;
      }

      const result = await performAttempt(attemptId, browserType, {
      run_mode: RUN_MODE,
      target_cell: targetCell,
      gender_counter: RUN_MODE === "quota" ? quotaState?.gender_counter : null,
      gender_expected_total:
        RUN_MODE === "quota" ? quotaTargetTotal : TARGET_SUCCESSFUL_SUBMITS,
      likert_rules: likertConfig.rules,
      likert_expected_total: likertExpectedTotal,
      likert_segment_key:
        RUN_MODE === "quota" && targetCell
          ? makeLikertSegmentKeyFromCell(targetCell)
          : "__global__",
      likert_segment_targets: likertConfig.segment_targets,
      likert_segment_expected_total: likertSegmentExpectedTotal,
      likert_progress:
        RUN_MODE === "quota"
          ? quotaState?.likert_progress
          : likertProgressMemory,
      schema: schemaDiscovery,
      on_submit_click:
        RUN_MODE === "quota" && quotaState
          ? async ({ quota_cell_key }) => {
              quotaState.submit_click_total =
                sanitizeTargetValue(quotaState.submit_click_total, 0) + 1;
              saveQuotaState(quotaState);
              if (LOG_PAGE_DEBUG) {
                console.log(
                  `attempt=${attemptId} submit_click_total=${quotaState.submit_click_total} cell=${quota_cell_key ?? "-"}`,
                );
              }
            }
          : null,
    });
      attempts.push(result);

      if (RUN_MODE === "quota" && quotaState) {
        quotaState.attempts_total += 1;
        incrementCounterKey(quotaState.attempt_status_counter, result.status, 1);
        if (
          result.status !== "success" &&
          sanitizeTargetValue(result.submit_clicks_attempt, 0) > 0
        ) {
          quotaState.submit_click_unconfirmed_total =
            sanitizeTargetValue(quotaState.submit_click_unconfirmed_total, 0) +
            sanitizeTargetValue(result.submit_clicks_attempt, 0);
        }
      }

      if (result.status === "success") {
        const runRecord = {
        run_id: successful_runs.length + 1,
        timestamp_start: result.start_ts,
        timestamp_end: result.end_ts,
        scheme: result.scheme,
        selection_stats: result.selection_stats,
        branch_choice: result.branch_choice,
        pages_visited: result.pages_visited,
        answers_summary: result.answers_summary,
        success_marker: result.success_marker,
        final_url: result.final_url,
      };

        if (RUN_MODE === "quota" && quotaState && quotaTargetsMap && targetCell) {
          const before = sanitizeTargetValue(quotaState.completed[targetCell.key], 0);
          const target = sanitizeTargetValue(quotaTargetsMap[targetCell.key], before);
          const after = Math.min(target, before + 1);
          quotaState.completed[targetCell.key] = after;
          quotaState.success_total += 1;
          const genderChoice = String(result.gender_choice ?? "").toLowerCase();
          if (genderChoice === "male" || genderChoice === "female") {
            const normalizedGenderCounter = normalizeGenderCounter(quotaState.gender_counter);
            normalizedGenderCounter[genderChoice] += 1;
            quotaState.gender_counter = normalizedGenderCounter;
          }
          quotaState.answer_counters = deriveAnswerCountersFromCompleted(
            quotaState.completed,
          );
          runRecord.quota_cell_key = targetCell.key;
          runRecord.quota_cell_before = before;
          runRecord.quota_cell_after = after;
          result.quota_cell_key = targetCell.key;
          result.quota_cell_before = before;
          result.quota_cell_after = after;
          runRecord.quota_dimensions = {
            city_key: targetCell.city_key,
            city_label: targetCell.city_label,
            trip_type: targetCell.trip_type,
            children_mode: targetCell.children_mode,
            provider_bucket: targetCell.provider_bucket,
          };
        }
        runRecord.gender_choice = result.gender_choice;
        runRecord.likert_answers = result.likert_answers;

        if (RUN_MODE === "quota" && quotaState) {
          quotaState.likert_progress = applyLikertAnswersToProgress(
            quotaState.likert_progress,
            result.likert_answers,
          );
        } else {
          likertProgressMemory = applyLikertAnswersToProgress(
            likertProgressMemory,
            result.likert_answers,
          );
        }

        successful_runs.push(runRecord);
        persistProgress();
        logAttemptProgress(result);
        continue;
      }

      if (result.status === "captcha_detected" && CAPTCHA_POLICY === "restart_browser") {
        await randomSleep(3_000, 8_000);
      }
      if (result.status === "load_error_screen") {
        await randomSleep(3_000, 8_000);
      }
      if (result.status === "network_changed") {
        await randomSleep(3_000, 8_000);
      }
      if (result.status === "browser_closed") {
        await randomSleep(2_000, 5_000);
      }
      if (result.status === "attempt_error") {
        await randomSleep(2_000, 5_000);
      }
      if (result.status === "validation_stuck") {
        await randomSleep(1_500, 3_500);
      }
      if (result.status === "selection_unresolved") {
        await randomSleep(1_500, 3_500);
      }
      if (result.status === "submit_unconfirmed") {
        await randomSleep(1_500, 3_500);
      }
      if (result.status === "start_page_invalid") {
        await randomSleep(2_000, 4_000);
      }
      persistProgress();
      logAttemptProgress(result);
    }
    if (
      RUN_MODE === "quota" &&
      quotaTargetsMap &&
      quotaState &&
      buildQuotaRemaining(quotaTargetsMap, quotaState.completed).total_remaining === 0
    ) {
      runExitReason = "quota_reached";
    } else if (attemptId >= MAX_TOTAL_ATTEMPTS) {
      runExitReason = "max_attempts_reached";
    } else if (RUN_MODE !== "quota" && successful_runs.length >= TARGET_SUCCESSFUL_SUBMITS) {
      runExitReason = "target_success_reached";
    } else {
      runExitReason = "stopped_without_target";
    }

    const report = persistProgress();
    console.log(`Report written: ${OUT_PATH}`);
    console.log(
      `successful_submits=${report.summary.successful_submits} attempts_total=${report.summary.attempts_total}`,
    );
  } finally {
    releaseLock();
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
