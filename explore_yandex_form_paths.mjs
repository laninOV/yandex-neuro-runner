#!/usr/bin/env node
import fs from "node:fs";

function parseCliArgs(argv) {
  const parsed = {};

  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) continue;

    const key = token.slice(2);
    const next = argv[index + 1];

    if (!next || next.startsWith("--")) {
      parsed[key] = true;
      continue;
    }

    parsed[key] = next;
    index += 1;
  }

  return parsed;
}

const CLI_ARGS = parseCliArgs(process.argv.slice(2));
const FORM_URL =
  CLI_ARGS.url ??
  process.env.FORM_URL ??
  "https://forms.yandex.ru/cloud/69b5f847068ff06856f78893";
const OUT_PATH = CLI_ARGS.out ?? process.env.OUT_PATH ?? "/tmp/yandex-form-paths.json";
const HEADLESS = String(CLI_ARGS.headless ?? process.env.HEADLESS ?? "1") !== "0";
const MAX_RUNS = Number(CLI_ARGS["max-runs"] ?? process.env.MAX_RUNS ?? "40");
const TARGET_SUCCESSFUL_RUNS = Number(
  CLI_ARGS["target-rules"] ??
    CLI_ARGS["target-successful-runs"] ??
    process.env.TARGET_SUCCESSFUL_RUNS ??
    "0",
);
const TARGET_RANDOM_RUNS = Number(
  CLI_ARGS["target-random"] ?? process.env.TARGET_RANDOM_RUNS ?? "0",
);
const RESUME =
  Boolean(CLI_ARGS.resume) || String(process.env.RESUME ?? "0") === "1";
const MAX_TOTAL_ATTEMPTS = Number(
  CLI_ARGS["max-total-attempts"] ??
    process.env.MAX_TOTAL_ATTEMPTS ??
    String(TARGET_SUCCESSFUL_RUNS > 0 ? Math.max(TARGET_SUCCESSFUL_RUNS, 1) : MAX_RUNS),
);
const SUBMIT_FINAL =
  Boolean(CLI_ARGS.submit) || String(process.env.SUBMIT_FINAL ?? "0") === "1";
const BROWSER = (CLI_ARGS.browser ?? process.env.BROWSER ?? "chromium").toLowerCase();
const CHANNEL = CLI_ARGS.channel ?? process.env.BROWSER_CHANNEL ?? "chrome";
const NAV_TIMEOUT_MS = Number(
  CLI_ARGS["nav-timeout-ms"] ?? process.env.NAV_TIMEOUT_MS ?? "60000",
);
const MAX_NAV_RETRIES = Number(
  CLI_ARGS["max-nav-retries"] ?? process.env.MAX_NAV_RETRIES ?? "3",
);
const BETWEEN_RUN_DELAY_MS = Number(
  CLI_ARGS["between-run-delay-ms"] ?? process.env.BETWEEN_RUN_DELAY_MS ?? "1200",
);
const RUN_RETRY_COUNT = Number(
  CLI_ARGS["run-retry-count"] ?? process.env.RUN_RETRY_COUNT ?? "3",
);
const ANSWER_RULES = [
  {
    pageNo: 2,
    label: "Сколько Вам лет?",
    type: "radio",
    mode: "fixed",
    value: "18-24",
  },
  {
    pageNo: 2,
    label: "Пользовались ли Вы нейросетями (ИИ-сервисами)?",
    type: "radio",
    mode: "fixed",
    value: "Да",
  },
  {
    pageNo: 2,
    label: "Пользовались ли Вы Нейро Алисой хотя бы один раз?",
    type: "radio",
    mode: "fixed",
    value: "Да",
  },
  {
    pageNo: 2,
    label: "Ваш пол",
    type: "radio",
    mode: "random",
  },
  {
    pageNo: 2,
    label: "В каком городе Вы проживаете?",
    type: "radio",
    mode: "fixed",
    value: "Москва",
  },
  {
    pageNo: 2,
    label: "Ваш род деятельности?",
    type: "radio",
    mode: "allowed",
    values: ["Студент", "Работаю", "Учусь и работаю"],
  },
  {
    pageNo: 3,
    label: "Как часто Вы используете нейросети?",
    type: "radio",
    mode: "random",
    values: ["Несколько раз в день", "Примерно раз в день"],
  },
  {
    pageNo: 3,
    label: "Какими нейросетями Вы пользовались?",
    type: "checkbox",
    mode: "weighted_multi",
    options: [
      { value: "Нейро Алиса", probability: 0.95 },
      { value: "ChatGPT", probability: 0.8 },
      { value: "DeepSeek", probability: 0.75 },
      { value: "GigaChat", probability: 0.35 },
    ],
    minSelected: 1,
  },
  {
    pageNo: 3,
    label: "Насколько хорошо Вы в целом разбираетесь в нейросетях?",
    type: "radio",
    mode: "allowed",
    values: ["Скорее хорошо", "Очень хорошо"],
  },
  {
    pageNo: 4,
    label: "Как давно Вы впервые попробовали Нейро Алису?",
    type: "radio",
    mode: "allowed",
    values: ["Менее недели назад", "От 1 недели до 1 месяца назад"],
  },
  {
    pageNo: 4,
    label: "Как часто Вы используете Нейро Алису?",
    type: "radio",
    mode: "allowed",
    values: [
      "Несколько раз в день",
      "Примерно раз в день",
      "Несколько раз в неделю",
    ],
  },
  {
    pageNo: 4,
    label: "Для каких задач Вы чаще всего обращаетесь к Нейро Алисе?",
    type: "checkbox",
    mode: "fixed_multi",
    values: ["Учеба", "Работа", "Написание текстов", "Генерация идей"],
  },
  {
    pageNo: 5,
    label: "Какие сервисы Яндекса Вы используете регулярно?",
    type: "checkbox",
    mode: "fixed_multi",
    values: [
      "Яндекс.Карты",
      "Яндекс.Музыка",
      "Яндекс.Почта",
      "Яндекс.Go / Такси",
      "Яндекс.Браузер",
      "Яндекс.Плюс",
    ],
  },
  {
    pageNo: 5,
    label: "Насколько для Вас важно, что Нейро Алиса встроена в экосистему Яндекса?",
    type: "radio",
    mode: "allowed",
    values: ["4", "5"],
  },
  {
    pageNo: 5,
    label: "Использовали бы Вы Нейро Алису так же часто, если бы она не была частью сервисов Яндекса?",
    type: "radio",
    mode: "allowed",
    values: ["Скорее нет", "Нет"],
  },
  {
    pageNo: 6,
    label: "Насколько Вы обеспокоены за безопасность данных при использовании нейросетей?",
    type: "radio",
    mode: "allowed",
    values: ["1", "2"],
  },
  {
    pageNo: 6,
    label: "Насколько Вы согласны с утверждением:",
    type: "radio",
    mode: "allowed",
    values: ["4", "5"],
  },
  {
    pageNo: 6,
    label: "Влияют ли опасения за безопасность данных на Ваше использование нейросетей?",
    type: "radio",
    mode: "allowed",
    values: ["Да, сильно влияют", "Да, немного влияют"],
  },
  {
    pageNo: 7,
    label: "Насколько Нейро Алиса помогает Вам решать задачи?",
    type: "radio",
    mode: "fixed",
    value: "1",
  },
  {
    pageNo: 7,
    label: "В каких задачах Вы используете Нейро Алису?",
    type: "checkbox",
    mode: "fixed_multi",
    values: [
      "Учеба",
      "Работа",
      "Написание текстов",
      "Поиск информации",
      "Идеи / мозговой штурм",
      "Развлечение",
    ],
  },
  {
    pageNo: 7,
    label: "Насколько Вы согласны с утверждением:",
    type: "radio",
    mode: "fixed",
    value: "1",
  },
  {
    pageNo: 8,
    label: "Насколько Вам удобно пользоваться Нейро Алисой?",
    type: "radio",
    mode: "fixed",
    value: "5",
  },
  {
    pageNo: 8,
    label: "Насколько корректно Нейро Алиса отвечает Вам при запросах?",
    type: "radio",
    mode: "fixed",
    value: "3",
  },
  {
    pageNo: 8,
    label: "Что Вам больше всего нравится в Нейро Алисе?",
    type: "checkbox",
    mode: "weighted_multi",
    requiredValues: [
      "Интеграция с сервисами Яндекса",
      "Доступность",
      "Стиль общения",
    ],
    options: [
      { value: "Простой интерфейс", probability: 0.3 },
      { value: "Быстрые ответы", probability: 0.25 },
    ],
    minSelected: 3,
  },
  {
    pageNo: 8,
    label: "Готовы ли Вы продолжать пользоваться Нейро Алисой в будущем?",
    type: "radio",
    mode: "allowed",
    values: [
      "Да, скорее всего буду использовать регулярно",
      "Возможно буду использовать",
    ],
  },
];

const normalize = (value) =>
  String(value ?? "")
    .replace(/\s+/g, " ")
    .replace(/^(\*|\s)+/, "")
    .replace(/Обязательное поле/g, "")
    .trim();

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const stableKey = (assignments) =>
  JSON.stringify(
    Object.entries(assignments)
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([key, value]) => [key, value]),
  );

const questionKey = (pageNo, label) => `p${pageNo}::${label}`;

function getAnswerRule(pageNo, question) {
  return (
    ANSWER_RULES.find(
      (rule) =>
        Number(rule.pageNo) === Number(pageNo) &&
        rule.label === question.label &&
        rule.type === question.type,
    ) ?? null
  );
}

function chooseRuleOption(rule, question, fallbackValue) {
  if (!rule) {
    return {
      option: fallbackValue,
      ruleMode: null,
      branchEnabled: true,
      allowedOptions: null,
    };
  }

  if (rule.mode === "random") {
    const options = Array.isArray(question.options) ? question.options : [];
    const candidateOptions = Array.isArray(rule.values) && rule.values.length
      ? options.filter((option) => rule.values.includes(option))
      : options;
    const option =
      candidateOptions[
        Math.floor(Math.random() * Math.max(1, candidateOptions.length))
      ] ?? fallbackValue;
    return {
      option,
      ruleMode: "random",
      branchEnabled: false,
      allowedOptions: null,
    };
  }

  if (rule.mode === "allowed") {
    const options = Array.isArray(question.options) ? question.options : [];
    const allowedOptions = options.filter((option) =>
      Array.isArray(rule.values) ? rule.values.includes(option) : false,
    );
    const option =
      allowedOptions.includes(fallbackValue) ? fallbackValue : allowedOptions[0] ?? fallbackValue;
    return {
      option,
      ruleMode: "allowed",
      branchEnabled: true,
      allowedOptions,
    };
  }

  return {
    option: rule.value ?? fallbackValue,
    ruleMode: rule.mode ?? "fixed",
    branchEnabled: false,
    allowedOptions: null,
  };
}

function pickRandom(values) {
  return values[Math.floor(Math.random() * Math.max(1, values.length))];
}

function sampleAssignmentsFromRules() {
  const assignments = {};

  for (const rule of ANSWER_RULES) {
    if (rule.mode !== "allowed") continue;
    if (!Array.isArray(rule.values) || !rule.values.length) continue;
    assignments[questionKey(rule.pageNo, rule.label)] = pickRandom(rule.values);
  }

  return assignments;
}

function isSuccessfulExit(exitReason) {
  return exitReason === "final-before-submit" || exitReason === "success";
}

function chooseCheckboxOptions(rule, question) {
  const options = Array.isArray(question.options) ? question.options : [];

  if (!rule) {
    return {
      options: options[0] ? [options[0]] : [],
      ruleMode: null,
    };
  }

  if (rule.mode === "fixed_multi") {
    const selected = (Array.isArray(rule.values) ? rule.values : []).filter((value) =>
      options.includes(value),
    );
    return {
      options: selected,
      ruleMode: "fixed_multi",
    };
  }

  if (rule.mode === "weighted_multi") {
    const specs = Array.isArray(rule.options) ? rule.options : [];
    const requiredValues = (Array.isArray(rule.requiredValues)
      ? rule.requiredValues
      : []
    ).filter((value) => options.includes(value));
    const selected = specs
      .filter((spec) => {
        if (!options.includes(spec?.value)) return false;
        if (requiredValues.includes(spec.value)) return false;
        const probability = Math.max(
          0,
          Math.min(1, Number(spec?.probability ?? spec?.weight ?? 0)),
        );
        return Math.random() < probability;
      })
      .map((spec) => spec.value);

    for (const value of requiredValues) {
      if (!selected.includes(value)) {
        selected.unshift(value);
      }
    }

    const minSelected = Math.max(0, Number(rule.minSelected ?? 0));
    if (selected.length < minSelected) {
      const fallbacks = [...specs]
        .filter((spec) => options.includes(spec?.value))
        .sort(
          (a, b) =>
            Number(b?.probability ?? b?.weight ?? 0) -
            Number(a?.probability ?? a?.weight ?? 0),
        )
        .map((spec) => spec.value);

      for (const value of fallbacks) {
        if (selected.includes(value)) continue;
        selected.push(value);
        if (selected.length >= minSelected) break;
      }
    }

    return {
      options: selected,
      ruleMode: "weighted_multi",
    };
  }

  if (rule.mode === "allowed_multi") {
    return {
      options: options.filter((value) =>
        Array.isArray(rule.values) ? rule.values.includes(value) : false,
      ),
      ruleMode: "allowed_multi",
    };
  }

  return {
    options: options[0] ? [options[0]] : [],
    ruleMode: rule.mode ?? null,
  };
}

async function loadPlaywright() {
  try {
    return await import("playwright");
  } catch (error) {
    throw new Error("Missing dependency `playwright`.", { cause: error });
  }
}

function selectBrowser(playwright) {
  const browserType = {
    chromium: playwright.chromium,
    firefox: playwright.firefox,
    webkit: playwright.webkit,
  }[BROWSER];

  if (!browserType) {
    throw new Error(`Unsupported browser: ${BROWSER}`);
  }

  return browserType;
}

async function createPage(browserType) {
  const launchOptions = {
    headless: HEADLESS,
    args: ["--disable-blink-features=AutomationControlled"],
  };

  if (BROWSER === "chromium" && CHANNEL) {
    launchOptions.channel = CHANNEL;
  }

  const browser = await browserType
    .launch(launchOptions)
    .catch(async () => browserType.launch({ headless: HEADLESS }));

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

  const page = await context.newPage();
  return { browser, context, page };
}

function pageNoFromUrl(url) {
  const match = /[?&]page=(\d+)/.exec(String(url ?? ""));
  return match ? Number(match[1]) : 1;
}

function screenTypeFromPage(pageInfo) {
  const hasOnlyStatements =
    pageInfo.questions.length > 0 &&
    pageInfo.questions.every((question) => question.type === "statement");

  if (pageInfo.submitVisible && (pageInfo.questions.length === 0 || hasOnlyStatements)) {
    return "final-before-submit";
  }
  if (
    pageInfo.pageNo === 1 &&
    pageInfo.nextVisible &&
    (pageInfo.questions.length === 0 || hasOnlyStatements)
  ) {
    return "intro";
  }
  if (
    !pageInfo.submitVisible &&
    !pageInfo.nextVisible &&
    (!pageInfo.questions.length || hasOnlyStatements)
  ) {
    return "end-screen";
  }
  return "questionnaire";
}

async function navigateToForm(page) {
  let lastError = null;

  for (let attempt = 1; attempt <= Math.max(1, MAX_NAV_RETRIES); attempt += 1) {
    try {
      await page.goto(FORM_URL, {
        waitUntil: "domcontentloaded",
        timeout: NAV_TIMEOUT_MS,
      });
      await page.waitForTimeout(1_500);

      const ready = await page.evaluate(() => {
        const bodyText = String(document.body?.innerText ?? "").toLowerCase();
        if (
          bodyText.includes("подтвердите, что запросы отправляли вы") ||
          bodyText.includes("вы не робот")
        ) {
          return "captcha";
        }
        if (document.querySelector(".SurveyPage-Name")) return "form";
        if ([...document.querySelectorAll("button")].some((button) =>
          String(button.innerText ?? "").includes("Далее"),
        )) {
          return "form";
        }
        return "unknown";
      });

      if (ready === "captcha") {
        throw new Error("Captcha detected while opening form");
      }
      if (ready === "form") {
        return;
      }
    } catch (error) {
      lastError = error;
    }

    await page.waitForTimeout(1_000 * attempt);
  }

  throw lastError ?? new Error("Unable to open form");
}

async function extractPage(page) {
  const pageInfo = await page.evaluate(() => {
    const normalizeText = (value) =>
      String(value ?? "")
        .replace(/\s+/g, " ")
        .replace(/^(\*|\s)+/, "")
        .replace(/Обязательное поле/g, "")
        .trim();

    const visible = (el) => {
      if (!el) return false;
      const style = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return (
        style.display !== "none" &&
        style.visibility !== "hidden" &&
        (rect.width > 0 || rect.height > 0)
      );
    };

    const pageNoMatch = /[?&]page=(\d+)/.exec(window.location.href);
    const pageNo = pageNoMatch ? Number(pageNoMatch[1]) : 1;
    const submitVisible = [...document.querySelectorAll("button")]
      .some((button) => normalizeText(button.innerText) === "Отправить");
    const nextVisible = [...document.querySelectorAll("button")]
      .some((button) => normalizeText(button.innerText) === "Далее");

    const questionNodes = [...document.querySelectorAll(".Question")]
      .filter(visible);

    const questions = questionNodes.map((node) => {
      const label =
        normalizeText(
          node.querySelector(".QuestionLabel-Text")?.innerText ||
            node.querySelector(".QuestionLabel")?.innerText ||
            node.querySelector("h3,h4")?.innerText ||
            "",
        );

      const radioInputs = [
        ...node.querySelectorAll('[role="radio"], input[type="radio"]'),
      ];
      const checkboxInputs = [
        ...node.querySelectorAll('[role="checkbox"], input[type="checkbox"]'),
      ];
      const textInputs = [...node.querySelectorAll('input[type="text"], input[type="number"]')];
      const textarea = node.querySelector("textarea");
      const tableRows = [...node.querySelectorAll("tr")];

      let type = "unknown";
      let options = [];
      let rows = [];
      let cols = [];

      if (tableRows.length > 1 && radioInputs.length > 0) {
        type = "matrix";
        cols = [...tableRows[0].querySelectorAll("th,td")]
          .slice(1)
          .map((cell) => normalizeText(cell.innerText))
          .filter(Boolean);
        rows = tableRows
          .slice(1)
          .map((row) => normalizeText(row.querySelector("th,td")?.innerText || ""))
          .filter(Boolean);
      } else if (radioInputs.length > 0) {
        type = "radio";
        options = [
          ...node.querySelectorAll(
            '.OptionContent-Text, .g-control-label__text, [role="radio"]',
          ),
        ]
          .map((el) =>
            normalizeText(
              el.getAttribute?.("aria-label") ??
                el.innerText ??
                el.textContent ??
                "",
            ),
          )
          .filter(Boolean)
          .filter((value, index, array) => array.indexOf(value) === index);
      } else if (checkboxInputs.length > 0) {
        type = "checkbox";
        options = [
          ...node.querySelectorAll(
            '.OptionContent-Text, .g-control-label__text, [role="checkbox"]',
          ),
        ]
          .map((el) =>
            normalizeText(
              el.getAttribute?.("aria-label") ??
                el.innerText ??
                el.textContent ??
                "",
            ),
          )
          .filter(Boolean)
          .filter((value, index, array) => array.indexOf(value) === index);
      } else if (textarea) {
        type = "textarea";
      } else if (textInputs.length > 0) {
        type = "text";
      } else if (label) {
        type = "statement";
      }

      return {
        label,
        type,
        required: Boolean(node.querySelector(".QuestionLabel-Required")),
        options,
        rows,
        cols,
      };
    });

    const pageHeading = normalizeText(
      document.querySelector(".SurveyPage-Name")?.innerText || document.title,
    );

    return {
      url: window.location.href,
      pageNo,
      pageHeading,
      nextVisible,
      submitVisible,
      buttonLabels: [...document.querySelectorAll("button")]
        .map((button) => normalizeText(button.innerText))
        .filter(Boolean),
      questions,
    };
  });

  return {
    ...pageInfo,
    screenType: screenTypeFromPage(pageInfo),
  };
}

async function clickIntroIfNeeded(page) {
  const introButton = page.getByRole("button", { name: "Далее", exact: true }).first();
  if (!(await introButton.isVisible().catch(() => false))) return;

  const pageInfo = await extractPage(page);
  if (pageInfo.pageNo !== 1 || pageInfo.questions.length > 0) return;

  await introButton.click();
  await sleep(300);
}

async function waitForReady(page) {
  await page.waitForLoadState("domcontentloaded");
  await page.waitForTimeout(1500);
}

async function answerRadio(page, label, option) {
  const ok = await page.evaluate(
    ({ label, option }) => {
      const normalizeText = (value) =>
        String(value ?? "")
          .replace(/\s+/g, " ")
          .replace(/^(\*|\s)+/, "")
          .replace(/Обязательное поле/g, "")
          .trim();

      const question = [...document.querySelectorAll(".Question")].find((node) => {
        const questionLabel = normalizeText(
          node.querySelector(".QuestionLabel-Text")?.innerText ||
            node.querySelector(".QuestionLabel")?.innerText ||
            node.querySelector("h3,h4")?.innerText ||
            "",
        );
        return questionLabel === label;
      });

      if (!question) return false;

      const labels = [...question.querySelectorAll("label")];
      const labelledTarget = labels.find((item) => {
        const text = normalizeText(item.innerText || item.textContent || "");
        return text === option;
      });

      const labelledInput = labelledTarget?.querySelector('input[type="radio"]');
      if (labelledInput) {
        labelledInput.click();
        return true;
      }

      const roleTarget = [...question.querySelectorAll('[role="radio"]')].find((item) => {
        const text = normalizeText(
          item.getAttribute("aria-label") ?? item.innerText ?? item.textContent ?? "",
        );
        return text === option;
      });

      if (roleTarget) {
        roleTarget.click();
        return true;
      }

      return false;
    },
    { label, option },
  );

  if (!ok) {
    throw new Error(`Unable to answer radio question "${label}" with "${option}"`);
  }
}

async function answerCheckbox(page, label, options) {
  const ok = await page.evaluate(
    ({ label, options }) => {
      const normalizeText = (value) =>
        String(value ?? "")
          .replace(/\s+/g, " ")
          .replace(/^(\*|\s)+/, "")
          .replace(/Обязательное поле/g, "")
          .trim();

      const question = [...document.querySelectorAll(".Question")].find((node) => {
        const questionLabel = normalizeText(
          node.querySelector(".QuestionLabel-Text")?.innerText ||
            node.querySelector(".QuestionLabel")?.innerText ||
            node.querySelector("h3,h4")?.innerText ||
            "",
        );
        return questionLabel === label;
      });

      if (!question) return false;

      const normalizedOptions = Array.isArray(options)
        ? options.map((option) => normalizeText(option)).filter(Boolean)
        : [];

      let clicked = 0;
      for (const option of normalizedOptions) {
        const labels = [...question.querySelectorAll("label")];
        const labelledTarget = labels.find((item) => {
          const text = normalizeText(item.innerText || item.textContent || "");
          return text === option;
        });

        const labelledInput = labelledTarget?.querySelector('input[type="checkbox"]');
        if (labelledInput) {
          if (!labelledInput.checked) labelledInput.click();
          clicked += 1;
          continue;
        }

        const roleTarget = [...question.querySelectorAll('[role="checkbox"]')].find((item) => {
          const text = normalizeText(
            item.getAttribute("aria-label") ?? item.innerText ?? item.textContent ?? "",
          );
          return text === option;
        });

        if (roleTarget) {
          roleTarget.click();
          clicked += 1;
        }
      }

      return clicked === normalizedOptions.length && normalizedOptions.length > 0;
    },
    { label, options },
  );

  if (!ok) {
    throw new Error(`Unable to answer checkbox question "${label}"`);
  }
}

async function answerText(page, label, value) {
  const ok = await page.evaluate(
    ({ label, value }) => {
      const normalizeText = (valueToNormalize) =>
        String(valueToNormalize ?? "")
          .replace(/\s+/g, " ")
          .replace(/^(\*|\s)+/, "")
          .replace(/Обязательное поле/g, "")
          .trim();

      const question = [...document.querySelectorAll(".Question")].find((node) => {
        const questionLabel = normalizeText(
          node.querySelector(".QuestionLabel-Text")?.innerText ||
            node.querySelector(".QuestionLabel")?.innerText ||
            node.querySelector("h3,h4")?.innerText ||
            "",
        );
        return questionLabel === label;
      });

      if (!question) return false;

      const input =
        question.querySelector("textarea") ||
        question.querySelector('input[type="text"]') ||
        question.querySelector('input[type="number"]');

      if (!input) return false;
      input.focus();
      input.value = value;
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(new Event("change", { bubbles: true }));
      return true;
    },
    { label, value },
  );

  if (!ok) {
    throw new Error(`Unable to fill text question "${label}"`);
  }
}

async function answerMatrix(page, label) {
  const ok = await page.evaluate((label) => {
    const normalizeText = (value) =>
      String(value ?? "")
        .replace(/\s+/g, " ")
        .replace(/^(\*|\s)+/, "")
        .replace(/Обязательное поле/g, "")
        .trim();

    const question = [...document.querySelectorAll(".Question")].find((node) => {
      const questionLabel = normalizeText(
        node.querySelector(".QuestionLabel-Text")?.innerText ||
          node.querySelector(".QuestionLabel")?.innerText ||
          node.querySelector("h3,h4")?.innerText ||
          "",
      );
      return questionLabel === label;
    });

    if (!question) return false;

    const rowGroups = [...question.querySelectorAll("tr")]
      .slice(1)
      .map((row) => [...row.querySelectorAll('input[type="radio"]')].filter((input) => !input.disabled));

    for (const inputs of rowGroups) {
      if (!inputs.length) continue;
      inputs[0].click();
    }

    return true;
  }, label);

  if (!ok) {
    throw new Error(`Unable to answer matrix question "${label}"`);
  }
}

async function clickNext(page, submitVisible) {
  const label = submitVisible ? "Отправить" : "Далее";
  const button = page.getByRole("button", { name: label, exact: true }).first();
  await button.click();
  await sleep(400);
}

async function discoverDynamicQuestions(
  page,
  assignments,
  branchOptions,
  uniqueQuestions,
  optionEffects,
) {
  const answered = new Set();
  const chosenAnswers = {};
  const questionSnapshots = [];

  for (let guard = 0; guard < 20; guard += 1) {
    const pageInfo = await extractPage(page);
    questionSnapshots.push(pageInfo);
    let changed = false;

    for (const question of pageInfo.questions) {
      uniqueQuestions.push({
        pageNo: pageInfo.pageNo,
        ...question,
      });

      const qKey = questionKey(pageInfo.pageNo, question.label);
      if (answered.has(qKey)) continue;

      if (question.type === "radio") {
        const answerRule = getAnswerRule(pageInfo.pageNo, question);
        const decision = chooseRuleOption(
          answerRule,
          question,
          assignments[qKey] ?? question.options[0],
        );
        const option = decision.option;
        if (decision.branchEnabled) {
          const candidateOptions = decision.allowedOptions ?? question.options;
          for (const candidate of candidateOptions) {
            if (candidate === option) continue;
            branchOptions.push({ key: qKey, option: candidate });
          }
        }

        const beforeLabels = new Set(pageInfo.questions.map((item) => item.label));
        await answerRadio(page, question.label, option);
        const afterPageInfo = await extractPage(page);
        const revealedQuestions = afterPageInfo.questions
          .filter((item) => !beforeLabels.has(item.label))
          .map((item) => item.label);

        optionEffects.push({
          pageNo: pageInfo.pageNo,
          question: question.label,
          type: question.type,
          option,
          rule_mode: decision.ruleMode,
          revealed_same_page_questions: revealedQuestions,
        });

        answered.add(qKey);
        chosenAnswers[qKey] = option;
        changed = true;
        break;
      }

      if (question.type === "checkbox" && question.options.length) {
        const decision = chooseCheckboxOptions(
          getAnswerRule(pageInfo.pageNo, question),
          question,
        );
        await answerCheckbox(page, question.label, decision.options);
        answered.add(qKey);
        chosenAnswers[qKey] = decision.options;
        optionEffects.push({
          pageNo: pageInfo.pageNo,
          question: question.label,
          type: question.type,
          option: decision.options,
          rule_mode: decision.ruleMode,
          revealed_same_page_questions: [],
        });
        changed = true;
        break;
      }

      if (question.type === "textarea") {
        await answerText(page, question.label, "тест");
        answered.add(qKey);
        chosenAnswers[qKey] = "тест";
        changed = true;
        break;
      }

      if (question.type === "text") {
        await answerText(page, question.label, "1");
        answered.add(qKey);
        chosenAnswers[qKey] = "1";
        changed = true;
        break;
      }

      if (question.type === "matrix") {
        await answerMatrix(page, question.label);
        answered.add(qKey);
        chosenAnswers[qKey] = question.cols[0] ?? "first-column";
        changed = true;
        break;
      }

      answered.add(qKey);
    }

    if (!changed) {
      return {
        pageInfo,
        chosenAnswers,
        questionSnapshots,
      };
    }
  }

  return {
    pageInfo: await extractPage(page),
    chosenAnswers,
    questionSnapshots,
  };
}

async function runPath(browserType, assignments) {
  const { browser, context, page } = await createPage(browserType);
  const visitedPages = [];
  const branchOptions = [];
  const uniqueQuestions = [];
  const optionEffects = [];
  const steps = [];
  let exitReason = "unknown";

  try {
    await navigateToForm(page);
    await waitForReady(page);
    await clickIntroIfNeeded(page);
    await waitForReady(page);

    for (let guard = 0; guard < 30; guard += 1) {
      const {
        pageInfo,
        chosenAnswers,
        questionSnapshots,
      } = await discoverDynamicQuestions(
        page,
        assignments,
        branchOptions,
        uniqueQuestions,
        optionEffects,
      );

      visitedPages.push(pageInfo);
      steps.push({
        pageNo: pageInfo.pageNo,
        screenType: pageInfo.screenType,
        chosenAnswers,
        questions: pageInfo.questions,
        snapshots_seen: questionSnapshots.map((snapshot) => ({
          pageNo: snapshot.pageNo,
          questions: snapshot.questions.map((question) => question.label),
        })),
      });

      if (!pageInfo.questions.length && !pageInfo.nextVisible && !pageInfo.submitVisible) {
        exitReason = page.url().includes("/success/") ? "success" : "no_questions";
        break;
      }

      if (pageInfo.submitVisible) {
        if (!SUBMIT_FINAL) {
          exitReason = "final-before-submit";
          break;
        }

        await clickNext(page, true);
        await page.waitForTimeout(1200);

        if (page.url().includes("/success/")) {
          exitReason = "success";
          break;
        }

        const afterSubmitPageInfo = await extractPage(page);
        if (
          !afterSubmitPageInfo.questions.length &&
          !afterSubmitPageInfo.nextVisible &&
          !afterSubmitPageInfo.submitVisible
        ) {
          exitReason = "success";
          break;
        }

        exitReason = "submit_failed";
        break;
      }

      if (!pageInfo.nextVisible) {
        exitReason = "blocked_without_next";
        break;
      }

      const beforeUrl = page.url();
      await clickNext(page, false);
      await page.waitForTimeout(800);

      if (page.url() === beforeUrl) {
        const nextPageInfo = await extractPage(page);
        if (JSON.stringify(nextPageInfo.questions) === JSON.stringify(pageInfo.questions)) {
          exitReason = "stuck_same_page";
          break;
        }
      }

      if (page.url().includes("/success/")) {
        exitReason = "success";
        break;
      }
    }
  } finally {
    await context.close().catch(() => {});
    await browser.close().catch(() => {});
  }

  return {
    assignments,
    assignment_key: stableKey(assignments),
    exit_reason: exitReason,
    visited_pages: visitedPages,
    branch_options: branchOptions,
    unique_questions: uniqueQuestions,
    option_effects: optionEffects,
    steps,
  };
}

async function runPathWithRetries(browserType, assignments) {
  let lastError = null;

  for (let attempt = 1; attempt <= Math.max(1, RUN_RETRY_COUNT); attempt += 1) {
    try {
      return await runPath(browserType, assignments);
    } catch (error) {
      lastError = error;
      await sleep(1_500 * attempt);
    }
  }

  return {
    assignments,
    assignment_key: stableKey(assignments),
    exit_reason: "run_error",
    error_message: String(lastError?.message ?? lastError ?? "Unknown error"),
    visited_pages: [],
    branch_options: [],
    unique_questions: [],
    option_effects: [],
    steps: [],
  };
}

function collectUniqueQuestions(runs) {
  const map = new Map();

  for (const run of runs) {
    for (const question of run.unique_questions ?? []) {
      const key = questionKey(question.pageNo, question.label);
      if (!map.has(key)) {
        map.set(key, {
          pageNo: question.pageNo,
          label: question.label,
          type: question.type,
          required: question.required,
          options: question.options ?? [],
          rows: question.rows ?? [],
          cols: question.cols ?? [],
        });
        continue;
      }

      const current = map.get(key);
      current.options = [...new Set([...(current.options ?? []), ...(question.options ?? [])])];
      current.rows = [...new Set([...(current.rows ?? []), ...(question.rows ?? [])])];
      current.cols = [...new Set([...(current.cols ?? []), ...(question.cols ?? [])])];
    }
  }

  return [...map.values()].sort((a, b) =>
    a.pageNo === b.pageNo ? a.label.localeCompare(b.label) : a.pageNo - b.pageNo,
  );
}

function collectOptionOutcomes(runs) {
  const map = new Map();

  for (const run of runs) {
    for (const effect of run.option_effects ?? []) {
      const key = stableKey({
        pageNo: effect.pageNo,
        question: effect.question,
        option: effect.option,
      });

      if (!map.has(key)) {
        map.set(key, {
          pageNo: effect.pageNo,
          question: effect.question,
          type: effect.type,
          option: effect.option,
          revealed_same_page_questions: new Set(),
          routes: new Set(),
        });
      }

      const item = map.get(key);
      for (const label of effect.revealed_same_page_questions ?? []) {
        item.revealed_same_page_questions.add(label);
      }
      item.routes.add(run.assignment_key);
    }
  }

  return [...map.values()]
    .map((item) => ({
      pageNo: item.pageNo,
      question: item.question,
      type: item.type,
      option: item.option,
      revealed_same_page_questions: [...item.revealed_same_page_questions].sort(),
      observed_in_routes: item.routes.size,
    }))
    .sort((a, b) =>
      a.pageNo === b.pageNo
        ? `${a.question} ${a.option}`.localeCompare(`${b.question} ${b.option}`)
        : a.pageNo - b.pageNo,
    );
}

function summarizePaths(runs) {
  return runs.map((run, index) => ({
    run_id: index + 1,
    assignments: run.assignments,
    exit_reason: run.exit_reason,
    pages: run.steps.map((step) => ({
      pageNo: step.pageNo,
      screen_type: step.screenType,
      chosen_answers: step.chosenAnswers,
      questions: step.questions.map((question) => ({
        label: question.label,
        type: question.type,
      })),
    })),
  }));
}

function buildReport(runs) {
  const questionRegistry = collectUniqueQuestions(runs);
  const optionOutcomes = collectOptionOutcomes(runs);

  return {
    form_url: FORM_URL,
    generated_at: new Date().toISOString(),
    max_runs: MAX_RUNS,
    target_successful_runs: TARGET_SUCCESSFUL_RUNS,
    target_random_runs: TARGET_RANDOM_RUNS,
    max_total_attempts: MAX_TOTAL_ATTEMPTS,
    submit_final: SUBMIT_FINAL,
    runs_executed: runs.length,
    successful_runs: runs.filter((run) => isSuccessfulExit(run.exit_reason)).length,
    failed_runs: runs.filter((run) => !isSuccessfulExit(run.exit_reason)).length,
    answer_rules: ANSWER_RULES,
    question_registry: questionRegistry,
    option_outcomes: optionOutcomes,
    routes: summarizePaths(runs),
  };
}

function writeReport(runs) {
  const report = buildReport(runs);
  fs.writeFileSync(OUT_PATH, `${JSON.stringify(report, null, 2)}\n`, "utf8");
  return report;
}

function logRunProgress(runIndex, run, stats) {
  const { successfulRuns, failedRuns, targetRuns, maxAttempts } = stats;
  const targetSuffix = targetRuns > 0 ? ` success=${successfulRuns}/${targetRuns}` : "";
  const attemptsSuffix = maxAttempts > 0 ? ` attempts=${runIndex}/${maxAttempts}` : "";
  const failedSuffix = ` failed=${failedRuns}`;
  console.log(
    `[run ${runIndex}] exit=${run.exit_reason}${targetSuffix}${attemptsSuffix}${failedSuffix}`,
  );
}

function loadExistingRoutesReport() {
  if (!RESUME || TARGET_SUCCESSFUL_RUNS <= 0) return null;
  if (!fs.existsSync(OUT_PATH)) return null;

  const raw = fs.readFileSync(OUT_PATH, "utf8");
  const parsed = JSON.parse(raw);
  const routes = Array.isArray(parsed?.routes) ? parsed.routes : [];

  return {
    runs: routes.map((route) => ({
      assignments: route.assignments ?? {},
      assignment_key: stableKey(route.assignments ?? {}),
      exit_reason: route.exit_reason ?? "unknown",
      visited_pages: [],
      branch_options: [],
      unique_questions: [],
      option_effects: [],
      steps: Array.isArray(route.pages)
        ? route.pages.map((page) => ({
            pageNo: page.pageNo,
            screenType: page.screen_type,
            chosenAnswers: page.chosen_answers ?? {},
            questions: Array.isArray(page.questions)
              ? page.questions.map((question) => ({
                  label: question.label,
                  type: question.type,
                }))
              : [],
            snapshots_seen: [],
          }))
        : [],
    })),
  };
}

async function main() {
  const playwright = await loadPlaywright();
  const browserType = selectBrowser(playwright);
  const resumeState = loadExistingRoutesReport();
  const runs = resumeState?.runs ?? [];

  if (TARGET_RANDOM_RUNS > 0) {
    console.log(
      `Note: --target-random=${TARGET_RANDOM_RUNS} received, but this script currently executes only rule-based runs.`,
    );
  }

  if (RESUME) {
    if (resumeState) {
      const resumedSuccesses = runs.filter((run) => isSuccessfulExit(run.exit_reason)).length;
      console.log(
        `Resume loaded: runs=${runs.length} successful=${resumedSuccesses} from ${OUT_PATH}`,
      );
    } else {
      console.log(`Resume requested, but no usable report found at ${OUT_PATH}. Starting fresh.`);
    }
  }

  if (TARGET_SUCCESSFUL_RUNS > 0) {
    let attempts = runs.length;
    let successfulRuns = runs.filter((run) => isSuccessfulExit(run.exit_reason)).length;

    while (
      attempts < Math.max(1, MAX_TOTAL_ATTEMPTS) &&
      successfulRuns < Math.max(1, TARGET_SUCCESSFUL_RUNS)
    ) {
      attempts += 1;
      const result = await runPathWithRetries(browserType, sampleAssignmentsFromRules());
      runs.push(result);
      if (isSuccessfulExit(result.exit_reason)) {
        successfulRuns += 1;
      }
      logRunProgress(runs.length, result, {
        successfulRuns,
        failedRuns: runs.length - successfulRuns,
        targetRuns: Math.max(1, TARGET_SUCCESSFUL_RUNS),
        maxAttempts: Math.max(1, MAX_TOTAL_ATTEMPTS),
      });
      writeReport(runs);

      if (BETWEEN_RUN_DELAY_MS > 0) {
        await sleep(BETWEEN_RUN_DELAY_MS);
      }
    }
  } else {
    const queue = [{}];
    const enqueued = new Set([stableKey({})]);
    const completed = new Set();

    while (queue.length && runs.length < MAX_RUNS) {
      const assignments = queue.shift();
      const assignmentKey = stableKey(assignments);
      if (completed.has(assignmentKey)) continue;

      const result = await runPathWithRetries(browserType, assignments);
      runs.push(result);
      completed.add(assignmentKey);
      logRunProgress(runs.length, result, {
        successfulRuns: runs.filter((run) => isSuccessfulExit(run.exit_reason)).length,
        failedRuns: runs.filter((run) => !isSuccessfulExit(run.exit_reason)).length,
        targetRuns: 0,
        maxAttempts: MAX_RUNS,
      });
      writeReport(runs);

      for (const branch of result.branch_options) {
        const nextAssignments = {
          ...assignments,
          [branch.key]: branch.option,
        };
        const nextKey = stableKey(nextAssignments);
        if (!enqueued.has(nextKey) && !completed.has(nextKey)) {
          queue.push(nextAssignments);
          enqueued.add(nextKey);
        }
      }

      if (BETWEEN_RUN_DELAY_MS > 0) {
        await sleep(BETWEEN_RUN_DELAY_MS);
      }
    }
  }

  const report = writeReport(runs);
  console.log(`Report written: ${OUT_PATH}`);
  console.log(`runs_executed=${runs.length}`);
  console.log(`successful_runs=${report.successful_runs}`);
  console.log(`failed_runs=${report.failed_runs}`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
