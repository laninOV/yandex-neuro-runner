#!/usr/bin/env node
import fs from "node:fs";
import path from "node:path";

const RUNNER_VERSION = "2026-03-14-fandom-scenarios-v1";
const DEFAULT_FORM_URL = "https://forms.yandex.ru/u/69b41052f47e7332deb465c4";
const DEFAULT_OUT_PATH = path.resolve(
  process.cwd(),
  "yandex-fandom-form-scenarios-report.json",
);
const DEFAULT_STATE_PATH = path.resolve(
  process.cwd(),
  "yandex-fandom-form-scenarios-state.json",
);

const INTRO_BUTTON_TEXT = "Далее";
const NEXT_BUTTON_TEXT = "Далее";
const SUBMIT_BUTTON_TEXT = "Отправить";
const SUCCESS_PATH_PART = "/success/";

const QUESTION_FANDOM =
  "Состоите ли вы сейчас в каком-либо фанатском сообществе (по сериалу, фильму, актеру, игре, музыкальной группе и др.)?";
const QUESTION_DURATION =
  "Как давно вы состоите или состояли в этом фанатском сообществе";
const QUESTION_AGE = "Укажите Ваш возраст";
const QUESTION_GENDER = "Укажите Ваш пол";
const QUESTION_OCCUPATION = "Укажите Ваш род занятий";

const FANDOM_OPTIONS = [
  "Да, состою сейчас",
  "Состоял(а) раньше, но сейчас нет",
  "Нет",
];

const DURATION_OPTIONS = [
  "Менее 6 месяцев",
  "6 месяцев – 1 год",
  "1–3 года",
  "Более 3 лет",
  "Затрудняюсь ответить",
];

const DEMOGRAPHIC_OPTIONS = {
  [QUESTION_AGE]: ["до 18 лет", "18–24", "25–34", "35–44", "45 и старше"],
  [QUESTION_GENDER]: ["Женский", "Мужской", "Предпочитаю не указывать"],
  [QUESTION_OCCUPATION]: [
    "Школьник / студент",
    "Работаю",
    "Совмещаю работу и учебу",
    "Временно не работаю",
  ],
};

const VALUES_IMPORTANCE = [
  "Мне важно продолжать мысленно возвращаться к событиям и миру истории после ее завершения",
  "Мне важно возвращаться к этой истории, потому что она напоминает мне о времени, когда я впервые ее смотрел(а)",
  "Мне важно обсуждать любимую историю и делиться своими впечатлениями с другими людьми",
  "Мне важно проживать и выражать эмоции, которые вызывает у меня эта история",
  "Мне важно чувствовать, что есть сообщество людей, которым нравится то же самое, что и мне",
  "Мне важно выражать свои идеи и интерпретации истории через творчество или теории",
  "Мне важно узнавать всё больше о мире любимой истории, потому что мне его всегда мало",
  "Мне важно разбираться в деталях сюжета и искать объяснения происходящему в истории",
];

const ACTIVITY_LABELS = [
  "Продолжать мысленно возвращаться к событиям и миру истории после ее завершения",
  "Возвращаться к этой истории, потому что она напоминает мне о времени, когда я впервые ее смотрел(а)",
  "Обсуждать любимую историю и делиться своими впечатлениями с другими людьми",
  "Проживать и выражать эмоции, которые вызывает у меня эта история",
  "Чувствовать, что есть сообщество людей, которым нравится то же самое, что и мне",
  "Выражать свои идеи и интерпретации истории через творчество или теории",
  "Узнавать всё больше о мире любимой истории, потому что мне его всегда мало",
  "Разбираться в деталях сюжета и искать объяснения происходящему в истории",
];

const SECTION_PAGES = [
  {
    page: 4,
    title: "Новости и официальные материалы (тизеры, посты, трейлеры)",
    scores: [1, 2, 4, 2, 1, 1, 2, 3],
  },
  {
    page: 5,
    title: "Фанатский контент (теории, арты, фанфики)",
    scores: [5, 3, 3, 4, 5, 4, 3, 2],
  },
  {
    page: 6,
    title: "Совместный просмотр",
    scores: [4, 3, 2, 3, 5, 4, 3, 4],
  },
  {
    page: 7,
    title: "Обсуждения серий и сюжета, построение теорий",
    scores: [4, 3, 5, 5, 5, 3, 3, 4],
  },
  {
    page: 8,
    title: "Интерактивные форматы вокруг сюжета (конкурсы, челленджи) от платформы",
    scores: [2, 1, 2, 2, 3, 3, 3, 2],
  },
  {
    page: 9,
    title: "Тизеры и закулисный контент от создателей проекта в социальных сетях",
    scores: [4, 2, 3, 4, 3, 3, 4, 5],
  },
  {
    page: 10,
    title: "Премьеры, фестивали и другие офлайн-события вокруг проекта",
    scores: [4, 2, 3, 4, 5, 3, 1, 2],
  },
];

const REQUIRED_SCHEMA_LABELS = [
  QUESTION_FANDOM,
  QUESTION_DURATION,
  ...VALUES_IMPORTANCE,
  ...ACTIVITY_LABELS,
  ...SECTION_PAGES.map((item) => item.title),
  QUESTION_AGE,
  QUESTION_GENDER,
  QUESTION_OCCUPATION,
];

const WEIGHTED_SCORES = [
  { score: 1, weight: 1 },
  { score: 2, weight: 2 },
  { score: 3, weight: 4 },
  { score: 4, weight: 2 },
  { score: 5, weight: 1 },
];
const STAR_SCALE_LABELS = [
  "Очень слабо",
  "Слабо",
  "Средне",
  "Сильно",
  "Очень сильно",
];
const RADIO_SELECTOR = '[role="radio"], input[type="radio"]';

const SHORT_BRANCH_OPTION = "Нет";

const parseArgs = (argv) => {
  const options = {
    scenario: "both",
    submit: false,
    headless: true,
    outPath: DEFAULT_OUT_PATH,
    statePath: DEFAULT_STATE_PATH,
    formUrl: DEFAULT_FORM_URL,
    browser: "chromium",
    channel: "chrome",
    randomBranch: "auto",
    seed: Date.now(),
    maxBrowserRestarts: 4,
    targetRandom: 0,
    targetRules: 0,
    resetState: false,
    retryDelayMs: 2_000,
    maxTotalAttempts: 0,
    recentAttemptsLimit: 200,
  };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    const next = argv[i + 1];

    if (arg === "--scenario" && next) {
      options.scenario = String(next).toLowerCase();
      i += 1;
      continue;
    }
    if (arg === "--submit") {
      options.submit = true;
      continue;
    }
    if (arg === "--submit=1") {
      options.submit = true;
      continue;
    }
    if (arg === "--submit=0") {
      options.submit = false;
      continue;
    }
    if (arg === "--headless" && next) {
      options.headless = next !== "0";
      i += 1;
      continue;
    }
    if (arg === "--out" && next) {
      options.outPath = path.resolve(process.cwd(), next);
      i += 1;
      continue;
    }
    if (arg === "--state-path" && next) {
      options.statePath = path.resolve(process.cwd(), next);
      i += 1;
      continue;
    }
    if (arg === "--form-url" && next) {
      options.formUrl = next;
      i += 1;
      continue;
    }
    if (arg === "--browser" && next) {
      options.browser = String(next).toLowerCase();
      i += 1;
      continue;
    }
    if (arg === "--channel" && next) {
      options.channel = next;
      i += 1;
      continue;
    }
    if (arg === "--random-branch" && next) {
      options.randomBranch = String(next).toLowerCase();
      i += 1;
      continue;
    }
    if (arg === "--seed" && next) {
      options.seed = Number(next);
      i += 1;
      continue;
    }
    if (arg === "--max-browser-restarts" && next) {
      options.maxBrowserRestarts = Number(next);
      i += 1;
      continue;
    }
    if (arg === "--target-random" && next) {
      options.targetRandom = Number(next);
      i += 1;
      continue;
    }
    if (arg === "--target-rules" && next) {
      options.targetRules = Number(next);
      i += 1;
      continue;
    }
    if (arg === "--reset-state") {
      options.resetState = true;
      continue;
    }
    if (arg === "--retry-delay-ms" && next) {
      options.retryDelayMs = Number(next);
      i += 1;
      continue;
    }
    if (arg === "--max-total-attempts" && next) {
      options.maxTotalAttempts = Number(next);
      i += 1;
      continue;
    }
    if (arg === "--recent-attempts-limit" && next) {
      options.recentAttemptsLimit = Number(next);
      i += 1;
      continue;
    }
  }

  return options;
};

const batchModeEnabled = (options) =>
  Number(options.targetRandom) > 0 || Number(options.targetRules) > 0;

const normalizeText = (value) =>
  String(value ?? "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[.*,!?:;()"'`]/g, "")
    .trim();

const pageNoFromUrl = (url) => {
  const match = /[?&]page=(\d+)/.exec(String(url ?? ""));
  return match ? Number(match[1]) : 1;
};

const mulberry32 = (seed) => {
  let t = seed >>> 0;
  return () => {
    t += 0x6d2b79f5;
    let r = Math.imul(t ^ (t >>> 15), t | 1);
    r ^= r + Math.imul(r ^ (r >>> 7), r | 61);
    return ((r ^ (r >>> 14)) >>> 0) / 4294967296;
  };
};

const pickOne = (rng, items) => {
  const index = Math.floor(rng() * items.length);
  return items[index];
};

const pickWeightedScore = (rng) => {
  const total = WEIGHTED_SCORES.reduce((sum, item) => sum + item.weight, 0);
  let target = rng() * total;
  for (const item of WEIGHTED_SCORES) {
    target -= item.weight;
    if (target <= 0) return item.score;
  }
  return 3;
};

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const closeBrowserSafely = async (browser) => {
  if (!browser) return;
  await Promise.race([browser.close().catch(() => {}), sleep(5_000)]);
};

const loadPlaywright = async () => {
  try {
    return await import("playwright");
  } catch (error) {
    throw new Error(
      "Missing dependency: playwright. Install it with `npm install playwright --no-save`.",
      { cause: error },
    );
  }
};

const readJsonIfExists = (filePath) => {
  if (!fs.existsSync(filePath)) return null;
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
};

const writeJsonAtomic = (filePath, data) => {
  ensureOutputDir(filePath);
  const tmpPath = `${filePath}.tmp`;
  fs.writeFileSync(tmpPath, `${JSON.stringify(data, null, 2)}\n`, "utf8");
  fs.renameSync(tmpPath, filePath);
};

const selectBrowserType = (playwright, browserName) => {
  const browserType = {
    chromium: playwright.chromium,
    firefox: playwright.firefox,
    webkit: playwright.webkit,
  }[browserName];

  if (!browserType) {
    throw new Error(`Unsupported browser: ${browserName}`);
  }

  return browserType;
};

const isCaptchaText = (value) => {
  const n = normalizeText(value);
  return (
    n.includes("вы не робот") ||
    n.includes("я не робот") ||
    n.includes("подтвердите что запросы отправляли вы а не робот")
  );
};

class CaptchaError extends Error {
  constructor(stage) {
    super(`Captcha detected at stage: ${stage}`);
    this.name = "CaptchaError";
    this.stage = stage;
  }
}

const isCaptchaError = (error) =>
  error instanceof CaptchaError ||
  (error instanceof Error && /Captcha detected at stage:/.test(error.message));

const ensureNoCaptcha = async (page, stage) => {
  const title = await page.title();
  const bodyText = await page.locator("body").innerText().catch(() => "");
  if (isCaptchaText(title) || isCaptchaText(bodyText)) {
    throw new CaptchaError(stage);
  }
};

const captureSurvey = async (page, formUrl) => {
  let captured = null;

  page.on("response", async (response) => {
    if (!response.url().includes("/u/gateway/root/form/getSurvey")) return;
    try {
      captured = await response.json();
    } catch {
      captured = null;
    }
  });

  await page.goto(formUrl, { waitUntil: "domcontentloaded" });
  await ensureNoCaptcha(page, "initial_load");

  const deadline = Date.now() + 10_000;
  while (!captured && Date.now() < deadline) {
    await sleep(100);
  }

  if (!captured) {
    throw new Error("Unable to capture survey schema from getSurvey response.");
  }

  return captured;
};

const validateSurveySchema = (survey) => {
  if (survey?.id !== "69b41052f47e7332deb465c4") {
    throw new Error(`Unexpected survey id: ${survey?.id ?? "unknown"}`);
  }

  const labels = new Set();
  for (const page of survey.pages ?? []) {
    for (const item of page.items ?? []) {
      if (item.label) labels.add(normalizeText(item.label));
      for (const row of item.rows ?? []) {
        if (row.label) labels.add(normalizeText(row.label));
      }
    }
  }

  const missing = REQUIRED_SCHEMA_LABELS.filter(
    (label) => !labels.has(normalizeText(label)),
  );

  if (missing.length > 0) {
    throw new Error(
      `Survey schema mismatch. Missing labels: ${missing.join(" | ")}`,
    );
  }
};

const resolveBranchChoice = (scenario, rng, randomBranch) => {
  if (scenario === "rules") return "Да, состою сейчас";
  if (randomBranch === "short") return SHORT_BRANCH_OPTION;
  if (randomBranch === "long") {
    return pickOne(rng, FANDOM_OPTIONS.filter((label) => label !== SHORT_BRANCH_OPTION));
  }
  return pickOne(rng, FANDOM_OPTIONS);
};

const buildRandomScaleAnswers = (rng, labels) =>
  Object.fromEntries(labels.map((label) => [label, pickWeightedScore(rng)]));

const buildRulesScaleAnswers = (labels, scores) =>
  Object.fromEntries(labels.map((label, index) => [label, scores[index]]));

const buildRunPlan = (scenario, rng, randomBranch) => {
  const fandomChoice = resolveBranchChoice(scenario, rng, randomBranch);
  const longBranch = fandomChoice !== SHORT_BRANCH_OPTION;

  const plan = {
    scenario,
    branch: longBranch ? "long" : "short",
    fandomChoice,
    pageAnswers: {
      2: {
        [QUESTION_FANDOM]: fandomChoice,
      },
    },
    selectedAnswers: {
      [QUESTION_FANDOM]: fandomChoice,
    },
  };

  if (longBranch) {
    const durationChoice =
      scenario === "rules" ? "Более 3 лет" : pickOne(rng, DURATION_OPTIONS);
    plan.pageAnswers[2][QUESTION_DURATION] = durationChoice;
    plan.selectedAnswers[QUESTION_DURATION] = durationChoice;

    const page3Answers =
      scenario === "rules"
        ? buildRulesScaleAnswers(VALUES_IMPORTANCE, [5, 3, 2, 3, 1, 4, 5, 2])
        : buildRandomScaleAnswers(rng, VALUES_IMPORTANCE);
    plan.pageAnswers[3] = page3Answers;
    Object.assign(plan.selectedAnswers, page3Answers);

    for (const section of SECTION_PAGES) {
      const answers =
        scenario === "rules"
          ? buildRulesScaleAnswers(ACTIVITY_LABELS, section.scores)
          : buildRandomScaleAnswers(rng, ACTIVITY_LABELS);
      plan.pageAnswers[section.page] = answers;
      plan.selectedAnswers[section.title] = answers;
    }

    const demographicAnswers = {
      [QUESTION_AGE]: pickOne(rng, DEMOGRAPHIC_OPTIONS[QUESTION_AGE]),
      [QUESTION_GENDER]: pickOne(rng, DEMOGRAPHIC_OPTIONS[QUESTION_GENDER]),
      [QUESTION_OCCUPATION]: pickOne(rng, DEMOGRAPHIC_OPTIONS[QUESTION_OCCUPATION]),
    };
    plan.pageAnswers[11] = demographicAnswers;
    Object.assign(plan.selectedAnswers, demographicAnswers);
  }

  return plan;
};

const summarizeError = (error) => ({
  name: error instanceof Error ? error.name : "Error",
  message: error instanceof Error ? error.message : String(error),
  stage:
    error && typeof error === "object" && "stage" in error ? error.stage : null,
});

const describeElementText = async (locator) =>
  locator.evaluate((element) => {
    const labelledBy = element.getAttribute("aria-labelledby");
    if (labelledBy) {
      return labelledBy
        .split(/\s+/)
        .filter(Boolean)
        .map((id) => document.getElementById(id)?.textContent ?? "")
        .join(" ")
        .trim();
    }

    return (
      element.getAttribute("aria-label") ??
      element.textContent ??
      ""
    ).trim();
  });

const findVisibleRadioGroup = async (page, label) => {
  const groups = page.getByRole("radiogroup");
  const target = normalizeText(label);
  const count = await groups.count();

  for (let i = 0; i < count; i += 1) {
    const group = groups.nth(i);
    if (!(await group.isVisible().catch(() => false))) continue;
    const name = normalizeText(await describeElementText(group));
    if (name === target || name.includes(target) || target.includes(name)) {
      return group;
    }
  }

  throw new Error(`Visible radiogroup not found for label: ${label}`);
};

const clickRadioInHeadingQuestion = async (page, label, optionLabel) => {
  const headings = page.locator("h3, h4");
  const target = normalizeText(label);
  const count = await headings.count();

  for (let i = 0; i < count; i += 1) {
    const heading = headings.nth(i);
    if (!(await heading.isVisible().catch(() => false))) continue;
    const text = normalizeText(await describeElementText(heading));
    if (!(text === target || text.includes(target) || target.includes(text))) {
      continue;
    }

    const result = await heading.evaluate(
      (element, payload) => {
        const normalize = (value) =>
          String(value ?? "")
            .toLowerCase()
            .replace(/\s+/g, " ")
            .replace(/[.*,!?:;()"'`]/g, "")
            .trim();

        let container = element.parentElement;
        while (container && !container.querySelector(payload.radioSelector)) {
          container = container.parentElement;
        }
        if (!container) {
          return { clicked: false, reason: "container_not_found" };
        }

        const radios = [...container.querySelectorAll(payload.radioSelector)].filter(
          (radio) => {
            const rect = radio.getBoundingClientRect();
            return (
              rect.width > 0 ||
              rect.height > 0 ||
              radio.getClientRects().length > 0
            );
          },
        );

        const names = radios.map((radio) =>
          normalize(radio.getAttribute("aria-label") ?? radio.textContent ?? ""),
        );

        let match = radios.find((radio) => {
          const radioText = normalize(
            radio.getAttribute("aria-label") ?? radio.textContent ?? "",
          );
          return (
            radioText === payload.target ||
            radioText.includes(payload.target) ||
            payload.target.includes(radioText)
          );
        });

        const numeric = Number(payload.rawOption);
        if (
          !match &&
          Number.isInteger(numeric) &&
          numeric >= 1 &&
          numeric <= radios.length
        ) {
          match = radios[numeric - 1];
        }

        if (!match && Number.isInteger(numeric) && payload.starScale[numeric - 1]) {
          const mappedTarget = normalize(payload.starScale[numeric - 1]);
          match = radios.find((radio) => {
            const radioText = normalize(
              radio.getAttribute("aria-label") ?? radio.textContent ?? "",
            );
            return (
              radioText === mappedTarget ||
              radioText.includes(mappedTarget) ||
              mappedTarget.includes(radioText)
            );
          });
        }

        if (!match) {
          return {
            clicked: false,
            reason: "option_not_found",
            names,
          };
        }

        match.click();
        return { clicked: true, names };
      },
      {
        rawOption: String(optionLabel),
        target: normalizeText(optionLabel),
        starScale: STAR_SCALE_LABELS,
        radioSelector: RADIO_SELECTOR,
      },
    );

    if (result?.clicked) {
      await sleep(120);
      return;
    }
  }

  throw new Error(`Question container not found for label: ${label}`);
};

const findRadioInContainer = async (container, optionLabel) => {
  const radios = container.getByRole("radio");
  const rawOption = String(optionLabel);
  const target = normalizeText(rawOption);
  const count = await radios.count();

  for (let i = 0; i < count; i += 1) {
    const radio = radios.nth(i);
    const text = normalizeText(await describeElementText(radio));
    if (text === target || text.includes(target) || target.includes(text)) {
      return radio;
    }
  }

  const numeric = Number(rawOption);
  if (
    Number.isInteger(numeric) &&
    numeric >= 1 &&
    numeric <= count
  ) {
    return radios.nth(numeric - 1);
  }

  const mappedStarLabel = STAR_SCALE_LABELS[numeric - 1];
  if (mappedStarLabel) {
    const mappedTarget = normalizeText(mappedStarLabel);
    for (let i = 0; i < count; i += 1) {
      const radio = radios.nth(i);
      const text = normalizeText(await describeElementText(radio));
      if (text === mappedTarget || text.includes(mappedTarget)) {
        return radio;
      }
    }
  }

  throw new Error(
    `Radio option not found in question for option: ${optionLabel}`,
  );
};

const chooseRadioOption = async (page, groupLabel, optionLabel) => {
  try {
    const container = await findVisibleRadioGroup(page, groupLabel);
    const radio = await findRadioInContainer(container, optionLabel);
    await radio.scrollIntoViewIfNeeded();
    await radio.click();
    await sleep(120);
    return;
  } catch {
    await clickRadioInHeadingQuestion(page, groupLabel, optionLabel);
  }
};

const isButtonVisible = async (page, label) =>
  page
    .getByRole("button", { name: label, exact: true })
    .first()
    .isVisible()
    .catch(() => false);

const clickPrimaryButton = async (page, label) => {
  const button = page.getByRole("button", { name: label, exact: true }).first();
  await button.waitFor({ state: "visible", timeout: 5_000 });
  const before = page.url();
  await button.scrollIntoViewIfNeeded();
  await button.click();
  await Promise.race([
    page.waitForURL((url) => String(url) !== before, { timeout: 5_000 }).catch(
      () => null,
    ),
    sleep(300),
  ]);
};

const answerCurrentPage = async (page, pageNo, answers) => {
  const entries = Object.entries(answers ?? {});
  for (const [questionLabel, answerValue] of entries) {
    const optionLabel = String(answerValue);
    await chooseRadioOption(page, questionLabel, optionLabel);
  }
};

const completeRun = async (page, runPlan, submit) => {
  await clickPrimaryButton(page, INTRO_BUTTON_TEXT);
  await ensureNoCaptcha(page, "after_intro");

  const seenPages = [];
  const stepLog = [];

  for (let guard = 0; guard < 30; guard += 1) {
    await ensureNoCaptcha(page, `loop_${guard}`);

    if (page.url().includes(SUCCESS_PATH_PART)) {
      return {
        status: "submitted",
        finalUrl: page.url(),
        seenPages,
        stepLog,
      };
    }

    const pageNo = pageNoFromUrl(page.url());
    if (!seenPages.includes(pageNo)) {
      seenPages.push(pageNo);
    }

    if (await isButtonVisible(page, SUBMIT_BUTTON_TEXT)) {
      if (!submit) {
        return {
          status: "ready_to_submit",
          finalUrl: page.url(),
          seenPages,
          stepLog,
        };
      }

      await clickPrimaryButton(page, SUBMIT_BUTTON_TEXT);
      await ensureNoCaptcha(page, "after_submit_click");
      await page.waitForURL(new RegExp(SUCCESS_PATH_PART), { timeout: 10_000 });
      return {
        status: "submitted",
        finalUrl: page.url(),
        seenPages,
        stepLog,
      };
    }

    const pageAnswers = runPlan.pageAnswers[pageNo];
    if (pageAnswers) {
      await answerCurrentPage(page, pageNo, pageAnswers);
      stepLog.push({
        page: pageNo,
        answers: pageAnswers,
      });
    }

    if (!(await isButtonVisible(page, NEXT_BUTTON_TEXT))) {
      throw new Error(`Next button is not visible on page ${pageNo}`);
    }

    await clickPrimaryButton(page, NEXT_BUTTON_TEXT);
    await ensureNoCaptcha(page, `after_next_page_${pageNo}`);
  }

  throw new Error("Loop guard exceeded while completing run.");
};

const launchBrowser = async (options) => {
  const playwright = await loadPlaywright();
  const browserType = selectBrowserType(playwright, options.browser);
  const launchOptions = {
    headless: options.headless,
    args: ["--disable-blink-features=AutomationControlled"],
  };

  if (options.browser === "chromium" && options.channel) {
    launchOptions.channel = options.channel;
  }

  try {
    return await browserType.launch(launchOptions);
  } catch (error) {
    if (!launchOptions.channel) throw error;
    delete launchOptions.channel;
    return browserType.launch(launchOptions);
  }
};

const createContext = async (browser) => {
  const context = await browser.newContext({
    locale: "ru-RU",
    viewport: { width: 1440, height: 1200 },
    userAgent:
      "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36",
  });

  await context.addInitScript(() => {
    Object.defineProperty(navigator, "webdriver", {
      configurable: true,
      get: () => undefined,
    });
  });

  return context;
};

const withBrowserRestarts = async (options, label, task) => {
  const maxAttempts = Math.max(1, Number(options.maxBrowserRestarts) + 1);

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    const browser = await launchBrowser(options);
    try {
      const value = await task(browser, attempt);
      return {
        value,
        browserRestartsUsed: attempt - 1,
      };
    } catch (error) {
      if (isCaptchaError(error) && attempt < maxAttempts) {
        console.error(
          `[captcha] ${label}: restarting browser (${attempt}/${maxAttempts - 1})`,
        );
        continue;
      }
      throw error;
    } finally {
      await closeBrowserSafely(browser);
    }
  }

  throw new Error(`Unable to finish ${label} after browser restarts.`);
};

const runScenario = async (browser, survey, options, scenario, seedValue) => {
  const context = await createContext(browser);
  const page = await context.newPage();
  const rng = mulberry32(seedValue >>> 0);

  try {
    const runPlan = buildRunPlan(scenario, rng, options.randomBranch);
    await page.goto(options.formUrl, { waitUntil: "domcontentloaded" });
    await ensureNoCaptcha(page, `${scenario}_initial_page`);

    const result = await completeRun(page, runPlan, options.submit);
    return {
      scenario,
      submit: options.submit,
      branch: runPlan.branch,
      fandom_choice: runPlan.fandomChoice,
      status: result.status,
      final_url: result.finalUrl,
      seen_pages: result.seenPages,
      step_log: result.stepLog,
      selected_answers: runPlan.selectedAnswers,
      seed_used: seedValue,
    };
  } finally {
    await context.close();
  }
};

const ensureOutputDir = (filePath) => {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
};

const writeReport = (filePath, report) => {
  writeJsonAtomic(filePath, report);
};

const createInitialState = (options) => ({
  runner_version: RUNNER_VERSION,
  created_at: new Date().toISOString(),
  updated_at: new Date().toISOString(),
  completed_at: null,
  form_url: options.formUrl,
  form_id: null,
  submit: options.submit,
  seed: options.seed,
  random_branch: options.randomBranch,
  max_browser_restarts: options.maxBrowserRestarts,
  targets: {
    random: Math.max(0, Math.floor(Number(options.targetRandom) || 0)),
    rules: Math.max(0, Math.floor(Number(options.targetRules) || 0)),
  },
  counts: {
    success: { random: 0, rules: 0, total: 0 },
    attempts: { random: 0, rules: 0, total: 0 },
    failures: { captcha: 0, other: 0, schema: 0, total: 0 },
    browser_restarts_total: 0,
  },
  last_attempt_id: 0,
  last_error: null,
  successful_runs: [],
  recent_attempts: [],
});

const normalizeState = (state, options) => {
  const base = createInitialState(options);
  const merged = {
    ...base,
    ...state,
    targets: {
      ...base.targets,
      ...(state?.targets ?? {}),
      random: Math.max(
        base.targets.random,
        Math.floor(Number(state?.targets?.random ?? base.targets.random) || 0),
      ),
      rules: Math.max(
        base.targets.rules,
        Math.floor(Number(state?.targets?.rules ?? base.targets.rules) || 0),
      ),
    },
    counts: {
      ...base.counts,
      ...(state?.counts ?? {}),
      success: { ...base.counts.success, ...(state?.counts?.success ?? {}) },
      attempts: { ...base.counts.attempts, ...(state?.counts?.attempts ?? {}) },
      failures: { ...base.counts.failures, ...(state?.counts?.failures ?? {}) },
    },
    successful_runs: Array.isArray(state?.successful_runs)
      ? state.successful_runs
      : [],
    recent_attempts: Array.isArray(state?.recent_attempts)
      ? state.recent_attempts
      : [],
  };

  merged.runner_version = RUNNER_VERSION;
  merged.form_url = options.formUrl;
  merged.submit = options.submit;
  merged.seed = options.seed;
  merged.random_branch = options.randomBranch;
  merged.max_browser_restarts = options.maxBrowserRestarts;
  merged.updated_at = new Date().toISOString();
  if (
    merged.counts.success.random < merged.targets.random ||
    merged.counts.success.rules < merged.targets.rules
  ) {
    merged.completed_at = null;
  }
  return merged;
};

const loadState = (options) => {
  if (options.resetState) {
    return createInitialState(options);
  }

  const loaded = readJsonIfExists(options.statePath);
  return loaded ? normalizeState(loaded, options) : createInitialState(options);
};

const pushRecentAttempt = (state, options, entry) => {
  state.recent_attempts.push(entry);
  const limit = Math.max(10, Math.floor(Number(options.recentAttemptsLimit) || 200));
  if (state.recent_attempts.length > limit) {
    state.recent_attempts.splice(0, state.recent_attempts.length - limit);
  }
};

const persistProgress = (state, options) => {
  state.updated_at = new Date().toISOString();
  writeReport(options.statePath, state);
  if (options.outPath !== options.statePath) {
    writeReport(options.outPath, state);
  }
};

const targetReached = (state, scenario) =>
  state.counts.success[scenario] >= state.targets[scenario];

const batchCompleted = (state) =>
  targetReached(state, "random") && targetReached(state, "rules");

const chooseNextScenario = (state) => {
  const remainingRandom = Math.max(0, state.targets.random - state.counts.success.random);
  const remainingRules = Math.max(0, state.targets.rules - state.counts.success.rules);

  if (remainingRandom === 0 && remainingRules === 0) return null;
  if (remainingRandom === 0) return "rules";
  if (remainingRules === 0) return "random";

  const randomRatio = state.counts.success.random / Math.max(1, state.targets.random);
  const rulesRatio = state.counts.success.rules / Math.max(1, state.targets.rules);
  if (randomRatio < rulesRatio) return "random";
  if (rulesRatio < randomRatio) return "rules";
  return remainingRules >= remainingRandom ? "rules" : "random";
};

const makeAttemptSeed = (options, attemptId, scenario) => {
  const scenarioOffset = scenario === "rules" ? 200_003 : 100_003;
  return (Number(options.seed) + attemptId * 9973 + scenarioOffset) >>> 0;
};

const logBatchProgress = (state) => {
  console.error(
    `[progress] random ${state.counts.success.random}/${state.targets.random} | rules ${state.counts.success.rules}/${state.targets.rules} | attempts ${state.counts.attempts.total} | failures ${state.counts.failures.total}`,
  );
};

const ensureSurveySchema = async (options, state) => {
  while (true) {
    try {
      const { value: survey, browserRestartsUsed } = await withBrowserRestarts(
        options,
        "schema_capture",
        async (browser) => {
          const schemaContext = await createContext(browser);
          try {
            const schemaPage = await schemaContext.newPage();
            const capturedSurvey = await captureSurvey(schemaPage, options.formUrl);
            validateSurveySchema(capturedSurvey);
            return capturedSurvey;
          } finally {
            await schemaContext.close().catch(() => {});
          }
        },
      );
      state.form_id = survey.id;
      state.counts.browser_restarts_total += browserRestartsUsed;
      persistProgress(state, options);
      return survey;
    } catch (error) {
      state.counts.failures.schema += 1;
      state.counts.failures.total += 1;
      state.last_error = summarizeError(error);
      pushRecentAttempt(state, options, {
        type: "schema_capture",
        at: new Date().toISOString(),
        status: "failed",
        error: summarizeError(error),
      });
      persistProgress(state, options);
      console.error(
        `[schema] failed: ${state.last_error.message}. Retrying in ${options.retryDelayMs}ms`,
      );
      await sleep(Math.max(0, Number(options.retryDelayMs) || 0));
    }
  }
};

const runBatchMode = async (options) => {
  if (!options.submit) {
    throw new Error("Batch mode requires --submit because counters increment only on real submits.");
  }

  const state = loadState(options);
  const survey = await ensureSurveySchema(options, state);

  if (batchCompleted(state)) {
    state.completed_at = state.completed_at ?? new Date().toISOString();
    persistProgress(state, options);
    return state;
  }

  while (!batchCompleted(state)) {
    if (Number(options.maxTotalAttempts) > 0 &&
        state.counts.attempts.total >= Number(options.maxTotalAttempts)) {
      throw new Error(
        `Reached max total attempts (${options.maxTotalAttempts}) before quotas were completed.`,
      );
    }

    const scenario = chooseNextScenario(state);
    if (!scenario) break;

    state.last_attempt_id += 1;
    const attemptId = state.last_attempt_id;
    const attemptSeed = makeAttemptSeed(options, attemptId, scenario);
    state.counts.attempts.total += 1;
    state.counts.attempts[scenario] += 1;
    persistProgress(state, options);

    try {
      const { value: run, browserRestartsUsed } = await withBrowserRestarts(
        options,
        `scenario_${scenario}_attempt_${attemptId}`,
        (browser) => runScenario(browser, survey, options, scenario, attemptSeed),
      );

      state.counts.browser_restarts_total += browserRestartsUsed;

      if (run.status !== "submitted") {
        throw new Error(`Unexpected run status: ${run.status}`);
      }

      state.counts.success.total += 1;
      state.counts.success[scenario] += 1;
      state.last_error = null;

      const successEntry = {
        attempt_id: attemptId,
        scenario,
        at: new Date().toISOString(),
        status: "submitted",
        branch: run.branch,
        fandom_choice: run.fandom_choice,
        final_url: run.final_url,
        seen_pages: run.seen_pages,
        browser_restarts_used: browserRestartsUsed,
        seed_used: run.seed_used,
        selected_answers: run.selected_answers,
      };

      state.successful_runs.push(successEntry);
      pushRecentAttempt(state, options, successEntry);
      persistProgress(state, options);
      logBatchProgress(state);
    } catch (error) {
      const errorSummary = summarizeError(error);
      state.last_error = errorSummary;
      state.counts.failures.total += 1;
      if (isCaptchaError(error)) {
        state.counts.failures.captcha += 1;
      } else {
        state.counts.failures.other += 1;
      }

      pushRecentAttempt(state, options, {
        attempt_id: attemptId,
        scenario,
        at: new Date().toISOString(),
        status: "failed",
        seed_used: attemptSeed,
        error: errorSummary,
      });
      persistProgress(state, options);

      console.error(
        `[attempt ${attemptId}] ${scenario} failed: ${errorSummary.message}. Retrying in ${options.retryDelayMs}ms`,
      );
      await sleep(Math.max(0, Number(options.retryDelayMs) || 0));
    }
  }

  state.completed_at = new Date().toISOString();
  persistProgress(state, options);
  return state;
};

const runSingleMode = async (options) => {
  const survey = await ensureSurveySchema(options, createInitialState(options));

  const scenarios =
    options.scenario === "both"
      ? ["random", "rules"]
      : [options.scenario];

  for (const scenario of scenarios) {
    if (!["random", "rules"].includes(scenario)) {
      throw new Error(`Unsupported scenario: ${scenario}`);
    }
  }

  const runs = [];
  for (let i = 0; i < scenarios.length; i += 1) {
    const scenario = scenarios[i];
    const seedValue = makeAttemptSeed(options, i + 1, scenario);
    const { value: run, browserRestartsUsed } = await withBrowserRestarts(
      options,
      `scenario_${scenario}`,
      (browser) => runScenario(browser, survey, options, scenario, seedValue),
    );
    runs.push({
      ...run,
      browser_restarts_used: browserRestartsUsed,
    });
  }

  const report = {
    runner_version: RUNNER_VERSION,
    created_at: new Date().toISOString(),
    form_url: options.formUrl,
    form_id: survey.id,
    submit: options.submit,
    scenario: options.scenario,
    random_branch: options.randomBranch,
    seed: options.seed,
    max_browser_restarts: options.maxBrowserRestarts,
    runs,
  };

  writeReport(options.outPath, report);
  if (options.statePath !== options.outPath) {
    writeReport(options.statePath, report);
  }
  return report;
};

const main = async () => {
  const options = parseArgs(process.argv.slice(2));
  const report = batchModeEnabled(options)
    ? await runBatchMode(options)
    : await runSingleMode(options);
  console.log(JSON.stringify(report, null, 2));
};

main().catch((error) => {
  const message = error instanceof Error ? error.stack ?? error.message : String(error);
  console.error(message);
  process.exitCode = 1;
});
