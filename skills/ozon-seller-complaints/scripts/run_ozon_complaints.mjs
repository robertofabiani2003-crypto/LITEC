import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";
import { chromium } from "playwright";

const DEFAULT_START_URL = "https://seller.ozon.ru/app/dashboard/main";

const DEFAULT_FLOW = {
  helpEntry: "Не нашли ответ на свой вопрос?",
  categoryProducts: "Товары и Цены",
  categoryQuality: "Контроль качества",
  categoryViolation: "Нарушение правил площадки другим продавцом",
  categoryContent: "Использование моих фото, видео, текста",
  actionComplain: "Пожаловаться на другой товар",
  promptOwnSku: "Пришлите SKU вашего товара",
  promptTargetSku: "Перейдите в карточку товара",
  promptSingleValue: "Пришлите только одно скопированное значение",
  successPrefix: "Скрыли товар",
  refresh: "Обновить",
  loadError: "Что-то пошло не так",
  chatListError: "Не удалось загрузить список чатов",
};

function parseArgs(argv) {
  const options = {
    configPath: null,
    cdpUrl: null,
    startIndex: 0,
    limit: null,
  };

  for (const arg of argv) {
    if (arg.startsWith("--config=")) {
      options.configPath = arg.slice("--config=".length);
    } else if (arg.startsWith("--cdp-url=")) {
      options.cdpUrl = arg.slice("--cdp-url=".length);
    } else if (arg.startsWith("--start-index=")) {
      options.startIndex = Number(arg.slice("--start-index=".length));
    } else if (arg.startsWith("--limit=")) {
      options.limit = Number(arg.slice("--limit=".length));
    }
  }

  if (!options.configPath) {
    throw new Error("Pass --config=/absolute/or/relative/path/to/config.json");
  }

  return options;
}

function resolveMaybeRelative(baseDir, value) {
  if (!value) {
    return null;
  }
  return path.isAbsolute(value) ? value : path.resolve(baseDir, value);
}

function loadJsonConfig(configPath) {
  const resolvedPath = path.resolve(configPath);
  const baseDir = path.dirname(resolvedPath);
  const config = JSON.parse(fs.readFileSync(resolvedPath, "utf8"));

  const runtimeDir = resolveMaybeRelative(baseDir, config.runtimeDir) ?? baseDir;
  const flow = { ...DEFAULT_FLOW, ...(config.flow ?? {}) };

  return {
    configPath: resolvedPath,
    baseDir,
    runtimeDir,
    cdpUrl: config.cdpUrl ?? "http://127.0.0.1:9223",
    startUrl: config.startUrl ?? DEFAULT_START_URL,
    ownSku: String(config.ownSku ?? "").trim(),
    attachments: (config.attachments ?? []).map((item) => resolveMaybeRelative(baseDir, item)),
    skuListPath: resolveMaybeRelative(baseDir, config.skuListPath),
    excelPath: resolveMaybeRelative(baseDir, config.excelPath),
    pythonPath: resolveMaybeRelative(baseDir, config.pythonPath),
    excelColumnIndex: Number(config.excelColumnIndex ?? 0),
    progressPath:
      resolveMaybeRelative(baseDir, config.progressPath) ?? path.join(runtimeDir, "ozon-complaints-progress.json"),
    logPath:
      resolveMaybeRelative(baseDir, config.logPath) ?? path.join(runtimeDir, "ozon-complaints-log.jsonl"),
    storageStatePath:
      resolveMaybeRelative(baseDir, config.storageStatePath) ?? path.join(runtimeDir, "ozon-storage-state.json"),
    maxAttemptsPerSku: Number(config.maxAttemptsPerSku ?? 3),
    closeBrowserOnFinish: Boolean(config.closeBrowserOnFinish ?? false),
    saveStorageStateOnFinish: Boolean(config.saveStorageStateOnFinish ?? true),
    flow,
  };
}

function validateConfig(config) {
  if (!config.ownSku) {
    throw new Error("Config field ownSku is required");
  }

  if (!config.skuListPath && !config.excelPath) {
    throw new Error("Provide either skuListPath or excelPath in config");
  }

  if (config.excelPath && !config.pythonPath) {
    throw new Error("excelPath requires pythonPath with pandas installed");
  }

  for (const filePath of config.attachments) {
    if (!fs.existsSync(filePath)) {
      throw new Error(`Attachment not found: ${filePath}`);
    }
  }

  if (config.skuListPath && !fs.existsSync(config.skuListPath)) {
    throw new Error(`SKU list file not found: ${config.skuListPath}`);
  }

  if (config.excelPath && !fs.existsSync(config.excelPath)) {
    throw new Error(`Excel file not found: ${config.excelPath}`);
  }

  if (config.pythonPath && !fs.existsSync(config.pythonPath)) {
    throw new Error(`Python executable not found: ${config.pythonPath}`);
  }

  fs.mkdirSync(config.runtimeDir, { recursive: true });
}

function readSkus(config) {
  if (config.skuListPath) {
    return fs
      .readFileSync(config.skuListPath, "utf8")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);
  }

  const py = `
import json
import pandas as pd
import sys

path = sys.argv[1]
column_index = int(sys.argv[2])
df = pd.read_excel(path)
values = [str(v).strip() for v in df.iloc[:, column_index].tolist() if str(v).strip() and str(v).strip() != 'nan']
print(json.dumps(values, ensure_ascii=False))
`;

  return JSON.parse(
    execFileSync(config.pythonPath, ["-", config.excelPath, String(config.excelColumnIndex)], {
      input: py,
      encoding: "utf8",
    })
  );
}

function loadProgress(progressPath) {
  if (!fs.existsSync(progressPath)) {
    return {
      completed: [],
      failed: [],
      startedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
  }

  const raw = fs.readFileSync(progressPath, "utf8").replace(/^\uFEFF/, "");
  return JSON.parse(raw);
}

function saveProgress(progressPath, progress) {
  progress.updatedAt = new Date().toISOString();
  fs.writeFileSync(progressPath, JSON.stringify(progress, null, 2));
}

function appendLog(logPath, entry) {
  fs.appendFileSync(logPath, `${JSON.stringify(entry)}\n`);
}

async function getPrimaryPage(browser) {
  for (const context of browser.contexts()) {
    for (const page of context.pages()) {
      if (page.url().includes("seller.ozon.ru")) {
        return { context, page };
      }
    }
  }

  const context = browser.contexts()[0];
  if (!context) {
    throw new Error("Browser context not found");
  }

  const page = context.pages()[0];
  if (!page) {
    throw new Error("No page available in browser context");
  }

  return { context, page };
}

async function saveStorageState(storageStatePath, context) {
  const state = await context.storageState();
  fs.writeFileSync(storageStatePath, JSON.stringify(state, null, 2));
}

async function waitForTail(page, predicate, timeoutMs = 90000, pollMs = 1500) {
  const deadline = Date.now() + timeoutMs;
  let tail = "";

  while (Date.now() < deadline) {
    tail = await page.evaluate(() => document.body?.innerText.slice(-4000) ?? "");
    if (predicate(tail)) {
      return tail;
    }
    await page.waitForTimeout(pollMs);
  }

  throw new Error(`Timeout waiting for page state. Last tail:\n${tail}`);
}

async function clickText(page, text, timeoutMs = 30000) {
  const locator = page.getByText(text, { exact: true }).last();
  await locator.waitFor({ state: "visible", timeout: timeoutMs });
  await locator.scrollIntoViewIfNeeded().catch(() => {});
  await locator.click({ timeout: timeoutMs }).catch(async () => {
    await locator.evaluate((node) => {
      const clickable = node.closest("button, a, [role='button'], label, div, span") ?? node;
      clickable.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, view: window }));
    });
  });
}

async function recoverChatLoadError(page, flow) {
  for (let attempt = 0; attempt < 3; attempt += 1) {
    const tail = await page.evaluate(() => document.body?.innerText.slice(-3500) ?? "");
    if (!tail.includes(flow.loadError) && !tail.includes(flow.chatListError)) {
      return;
    }
    await clickText(page, flow.refresh);
    await page.waitForTimeout(3000);
  }
}

async function prepareComplaintFlow(page, config) {
  const { flow } = config;
  const tail = await page.evaluate(() => document.body?.innerText.slice(-3500) ?? "");
  if (tail.includes(flow.actionComplain) || tail.includes(flow.promptOwnSku)) {
    return;
  }

  const helpButton = page.locator('button[data-onboarding-target="floating-help-button"]').first();
  await helpButton.waitFor({ state: "visible", timeout: 30000 });
  await helpButton.click();
  await page.waitForTimeout(1500);

  await clickText(page, flow.helpEntry);
  await page.waitForTimeout(2500);

  await recoverChatLoadError(page, flow);

  const afterHelp = await page.evaluate(() => document.body?.innerText.slice(-3500) ?? "");
  if (afterHelp.includes(flow.promptOwnSku) || afterHelp.includes(flow.actionComplain)) {
    return;
  }

  await clickText(page, flow.categoryProducts);
  await page.waitForTimeout(1500);
  await clickText(page, flow.categoryQuality);
  await page.waitForTimeout(1500);
  await clickText(page, flow.categoryViolation);
  await page.waitForTimeout(1500);
  await clickText(page, flow.categoryContent);

  await waitForTail(
    page,
    (nextTail) => nextTail.includes(flow.promptOwnSku) || nextTail.includes(flow.actionComplain),
    60000,
    1500
  );
}

function visibleTextarea(page) {
  return page.locator("textarea:visible").first();
}

function composerButtons(page) {
  return page.locator("div.om_17inde--oQE4A button:visible");
}

async function clickComplaintAction(page, flow) {
  await clickText(page, flow.actionComplain);
  const textarea = visibleTextarea(page);
  await textarea.waitFor({ state: "visible", timeout: 30000 });
  await page.waitForFunction(() => {
    const nodes = Array.from(document.querySelectorAll("textarea"));
    const visible = nodes.find((node) => {
      const style = window.getComputedStyle(node);
      return style.visibility !== "hidden" && style.display !== "none" && node.offsetParent !== null;
    });
    return Boolean(visible && !visible.disabled);
  }, null, { timeout: 30000 });
}

async function sendComposerText(page, value) {
  const textarea = visibleTextarea(page);
  const sendButton = composerButtons(page).last();
  await textarea.waitFor({ state: "visible", timeout: 30000 });
  await textarea.fill(value);
  await page.waitForTimeout(400);
  await sendButton.waitFor({ state: "visible", timeout: 30000 });
  await sendButton.click();
}

async function sendAttachments(page, attachments) {
  const buttons = composerButtons(page);
  const attachButton = buttons.first();
  const sendButton = buttons.last();
  const chooserPromise = page.waitForEvent("filechooser");
  await attachButton.click();
  const chooser = await chooserPromise;
  await chooser.setFiles(attachments);
  await page.waitForTimeout(5000);
  await sendButton.click();
}

async function processSkuAttempt(page, sku, config) {
  const { flow, ownSku, attachments } = config;

  await prepareComplaintFlow(page, config);

  const tail = await page.evaluate(() => document.body?.innerText.slice(-3500) ?? "");
  if (tail.includes(flow.actionComplain)) {
    await clickComplaintAction(page, flow);
    await page.waitForTimeout(1000);
  }

  await sendComposerText(page, ownSku);

  await waitForTail(
    page,
    (nextTail) => nextTail.includes(flow.promptTargetSku) || nextTail.includes(flow.promptSingleValue),
    45000
  );

  await sendComposerText(page, sku);

  await waitForTail(
    page,
    (nextTail) => nextTail.includes(`товар с артикулом ${sku}`) || nextTail.includes(`${flow.successPrefix} ${sku}`),
    60000
  );

  const preAttachTail = await page.evaluate(() => document.body?.innerText.slice(-4000) ?? "");
  if (!preAttachTail.includes(`${flow.successPrefix} ${sku}`) && attachments.length > 0) {
    await sendAttachments(page, attachments);
  }

  return await waitForTail(
    page,
    (nextTail) => nextTail.includes(`${flow.successPrefix} ${sku}`),
    120000,
    2500
  );
}

async function processSku(page, sku, config, logPath) {
  let lastError = null;

  for (let attempt = 1; attempt <= config.maxAttemptsPerSku; attempt += 1) {
    try {
      return await processSkuAttempt(page, sku, config);
    } catch (error) {
      lastError = error;
      const message = error instanceof Error ? error.message : String(error);
      const shouldRetry =
        message.includes("Timeout waiting for page state") &&
        message.includes(String(sku)) &&
        attempt < config.maxAttemptsPerSku;

      if (!shouldRetry) {
        throw error;
      }

      appendLog(logPath, {
        event: "retry",
        sku,
        attempt,
        reason: "hidden-confirmation-timeout",
        error: message,
        retryAt: new Date().toISOString(),
      });

      await page.goto(config.startUrl, { waitUntil: "domcontentloaded" });
      await page.waitForTimeout(3000);
    }
  }

  throw lastError;
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const config = loadJsonConfig(options.configPath);
  if (options.cdpUrl) {
    config.cdpUrl = options.cdpUrl;
  }

  validateConfig(config);

  const skus = readSkus(config);
  const progress = loadProgress(config.progressPath);
  const completedSet = new Set(progress.completed.map((item) => item.sku));
  const failedSet = new Set(progress.failed.map((item) => item.sku));
  const browser = await chromium.connectOverCDP(config.cdpUrl);

  try {
    const { context, page } = await getPrimaryPage(browser);
    await page.bringToFront();
    await page.goto(config.startUrl, { waitUntil: "domcontentloaded" });
    await page.waitForTimeout(3000);

    let processed = 0;
    for (let index = options.startIndex; index < skus.length; index += 1) {
      if (options.limit !== null && processed >= options.limit) {
        break;
      }

      const sku = skus[index];
      if (completedSet.has(sku) || failedSet.has(sku)) {
        continue;
      }

      const startedAt = new Date().toISOString();
      appendLog(config.logPath, { event: "start", index, sku, startedAt });

      try {
        const finalTail = await processSku(page, sku, config, config.logPath);
        const finishedAt = new Date().toISOString();
        const record = { index, sku, startedAt, finishedAt };
        progress.completed.push(record);
        completedSet.add(sku);
        saveProgress(config.progressPath, progress);
        appendLog(config.logPath, {
          event: "success",
          index,
          sku,
          finishedAt,
          tail: finalTail.slice(-1200),
        });
      } catch (error) {
        const finishedAt = new Date().toISOString();
        const record = {
          index,
          sku,
          startedAt,
          finishedAt,
          error: error instanceof Error ? error.message : String(error),
        };
        progress.failed.push(record);
        failedSet.add(sku);
        saveProgress(config.progressPath, progress);
        appendLog(config.logPath, { event: "failure", ...record });
        throw error;
      }

      processed += 1;
    }

    if (config.saveStorageStateOnFinish) {
      await saveStorageState(config.storageStatePath, context);
    }
  } finally {
    if (config.closeBrowserOnFinish) {
      await browser.close().catch(() => {});
    }
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
