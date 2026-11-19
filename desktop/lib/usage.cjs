const ISO_DATETIME = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})$/;
const USAGE_TIMEOUT_MS = 15_000;
const MAX_USAGE_ITEMS = 8;

function isObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function isIsoDateTime(value) {
  return typeof value === "string" && ISO_DATETIME.test(value) && Number.isFinite(Date.parse(value));
}

function copyNumber(source, key, target, { nullable = false, required = false } = {}) {
  if (!Object.hasOwn(source, key)) return !required;
  const value = source[key];
  if (nullable && value === null) return true;
  if (typeof value !== "number" || !Number.isFinite(value)) return false;
  target[key] = value;
  return true;
}

function copyDate(source, key, target, { nullable = true } = {}) {
  if (!Object.hasOwn(source, key)) return true;
  const value = source[key];
  if (nullable && value === null) return true;
  if (!isIsoDateTime(value)) return false;
  target[key] = value;
  return true;
}

function copyString(source, key, target, { maxLength = 120, nullable = false } = {}) {
  if (!Object.hasOwn(source, key)) return true;
  const value = source[key];
  if (nullable && value === null) return true;
  if (typeof value !== "string" || value.length > maxLength) return false;
  target[key] = value;
  return true;
}

function copyBoolean(source, key, target) {
  if (!Object.hasOwn(source, key)) return true;
  if (typeof source[key] !== "boolean") return false;
  target[key] = source[key];
  return true;
}

function sanitizeQuota(value) {
  if (!isObject(value)) return null;
  const result = {};
  if (!copyNumber(value, "limit", result, { required: true }) ||
      !copyNumber(value, "used", result, { required: true }) ||
      !copyNumber(value, "remaining", result, { required: true }) ||
      !copyString(value, "unit", result, { maxLength: 12 })) return null;
  if (result.unit !== undefined && result.unit !== "USD") return null;
  return result;
}

function sanitizeRateLimit(value) {
  if (!isObject(value) || !["5h", "1d", "7d"].includes(value.window)) return null;
  const result = { window: value.window };
  if (!copyNumber(value, "limit", result, { required: true }) ||
      !copyNumber(value, "used", result, { required: true }) ||
      !copyNumber(value, "remaining", result, { required: true }) ||
      !copyDate(value, "window_start", result) ||
      !copyDate(value, "reset_at", result) ||
      !copyDate(value, "expires_at", result) ||
      !copyNumber(value, "days_until_expiry", result) ||
      !copyString(value, "status", result, { maxLength: 80 }) ||
      !copyBoolean(value, "isValid", result)) return null;
  return result;
}

function sanitizeSubscription(value) {
  if (!isObject(value)) return null;
  const result = {};
  const fields = [
    ["daily_usage_usd", false], ["weekly_usage_usd", false], ["monthly_usage_usd", false],
    ["daily_limit_usd", true], ["weekly_limit_usd", true], ["monthly_limit_usd", true]
  ];
  for (const [key, nullable] of fields) {
    if (!copyNumber(value, key, result, { nullable })) return null;
  }
  for (const [usageKey, limitKey] of [
    ["daily_usage_usd", "daily_limit_usd"],
    ["weekly_usage_usd", "weekly_limit_usd"],
    ["monthly_usage_usd", "monthly_limit_usd"]
  ]) {
    if (typeof result[limitKey] === "number" && result[limitKey] > 0 && typeof result[usageKey] !== "number") return null;
  }
  if (!copyDate(value, "weekly_window_start", result) || !copyDate(value, "expires_at", result)) return null;
  return result;
}

function sanitizeUsage(payload) {
  if (!isObject(payload) || !["quota_limited", "unrestricted"].includes(payload.mode)) return null;
  const result = { mode: payload.mode };
  if (!copyBoolean(payload, "isValid", result) ||
      !copyString(payload, "status", result, { maxLength: 80 }) ||
      !copyString(payload, "planName", result) ||
      !copyString(payload, "unit", result, { maxLength: 12 }) ||
      !copyDate(payload, "expires_at", result) ||
      !copyNumber(payload, "days_until_expiry", result) ||
      !copyNumber(payload, "remaining", result) ||
      !copyNumber(payload, "balance", result)) return null;
  if (result.unit !== undefined && result.unit !== "USD") return null;

  if (Object.hasOwn(payload, "quota")) {
    result.quota = sanitizeQuota(payload.quota);
    if (!result.quota) return null;
  }
  if (Object.hasOwn(payload, "rate_limits")) {
    if (!Array.isArray(payload.rate_limits) || payload.rate_limits.length > MAX_USAGE_ITEMS) return null;
    result.rate_limits = [];
    for (const item of payload.rate_limits) {
      const cleaned = sanitizeRateLimit(item);
      if (!cleaned) return null;
      result.rate_limits.push(cleaned);
    }
  }
  if (Object.hasOwn(payload, "subscription")) {
    result.subscription = sanitizeSubscription(payload.subscription);
    if (!result.subscription) return null;
  }
  if (result.planName === "钱包余额" && typeof result.balance !== "number") return null;
  return result;
}

async function requestUsage(credential, fetchImpl = globalThis.fetch, timeoutMs = USAGE_TIMEOUT_MS) {
  if (!credential || typeof credential.apiKey !== "string" || !credential.apiKey ||
      typeof credential.baseUrl !== "string" || typeof fetchImpl !== "function") {
    return { state: "unavailable" };
  }
  let response;
  try {
    const baseUrl = credential.baseUrl.replace(/\/+$/, "");
    response = await fetchImpl(`${baseUrl}/usage`, {
      headers: { Authorization: `Bearer ${credential.apiKey}`, Accept: "application/json" },
      redirect: "error",
      signal: AbortSignal.timeout(timeoutMs)
    });
  } catch {
    return { state: "unavailable" };
  }
  if (response.status === 401 || response.status === 403) return { state: "unauthorized" };
  if (response.status >= 500) return { state: "unavailable" };
  if (!response.ok) return { state: "invalid" };
  let payload;
  try { payload = await response.json(); } catch { return { state: "invalid" }; }
  const usage = sanitizeUsage(payload);
  return usage ? { state: "available", usage } : { state: "invalid" };
}

function exhaustedWindows(usage) {
  const values = [];
  for (const item of usage?.rate_limits || []) {
    if (item.remaining <= 0) values.push({ window: item.window, resetAt: item.reset_at });
  }
  if (usage?.quota && usage.quota.remaining <= 0) values.push({ window: "quota" });
  const subscription = usage?.subscription;
  for (const [window, usedKey, limitKey] of [
    ["1d", "daily_usage_usd", "daily_limit_usd"],
    ["7d", "weekly_usage_usd", "weekly_limit_usd"],
    ["30d", "monthly_usage_usd", "monthly_limit_usd"]
  ]) {
    const used = subscription?.[usedKey];
    const limit = subscription?.[limitKey];
    if (typeof used === "number" && typeof limit === "number" && limit > 0 && used >= limit) {
      let resetAt;
      if (window === "7d" && isIsoDateTime(subscription.weekly_window_start)) {
        resetAt = new Date(Date.parse(subscription.weekly_window_start) + 7 * 24 * 60 * 60 * 1000).toISOString();
      }
      values.push({ window, resetAt });
    }
  }
  return values;
}

function quotaWindow(error) {
  const text = `${error?.code || ""} ${error?.message || ""}`;
  if (/API_KEY_RATE_5H|5\s*(?:h\b|小时)/i.test(text)) return "5h";
  if (/API_KEY_RATE_1D|\bdaily\b|日限额|日额度/i.test(text)) return "1d";
  if (/API_KEY_RATE_7D|USER_PLATFORM_WEEKLY|\bweekly\b|7\s*天|周限额|周额度/i.test(text)) return "7d";
  if (/USER_PLATFORM_MONTHLY|\bmonthly\b|月限额|月额度/i.test(text)) return "30d";
  if (/API_KEY_QUOTA|总额度/i.test(text)) return "quota";
  return null;
}

function quotaWindowLabel(window) {
  return ({ "5h": "5 小时", "1d": "日", "7d": "周", "30d": "月", quota: "总额度" })[window] || "模型";
}

function quotaExhaustionMessage(error, usage, { now = Date.now(), timeZone } = {}) {
  if (error?.type !== "rate_limit_exceeded" && error?.code !== "rate_limit_exceeded") return null;
  const candidates = exhaustedWindows(usage);
  const code = typeof error.code === "string" ? error.code : "";
  const message = typeof error.message === "string" ? error.message : "";
  const resetAt = isIsoDateTime(error.resetAt) ? error.resetAt : "";
  const quotaSignal = Boolean(resetAt) || candidates.length > 0 ||
    /quota|限额|额度|exhausted/i.test(`${code} ${message}`);
  if (!quotaSignal) return null;

  const explicitWindow = quotaWindow(error);
  const selected = explicitWindow ? candidates.filter((item) => item.window === explicitWindow) : candidates;
  const windows = selected.length ? [...new Set(selected.map((item) => item.window))] :
    explicitWindow ? [explicitWindow] : [];
  const label = windows.length ? windows.map(quotaWindowLabel).join("/") : "模型";

  let resetTime = resetAt ? Date.parse(resetAt) : NaN;
  if (!Number.isFinite(resetTime) && Number.isFinite(error.retryAfter) && error.retryAfter > 0 && error.retryAfter <= 30 * 24 * 60 * 60) {
    resetTime = now + error.retryAfter * 1000;
  }
  if (!Number.isFinite(resetTime)) {
    const resets = selected.map((item) => Date.parse(item.resetAt)).filter(Number.isFinite);
    if (resets.length) resetTime = Math.min(...resets);
  }
  if (!Number.isFinite(resetTime)) return `额度已用尽（${label}），恢复时间暂不可用`;
  const formatOptions = { dateStyle: "medium", timeStyle: "medium" };
  if (timeZone) formatOptions.timeZone = timeZone;
  const formatted = new Intl.DateTimeFormat("zh-CN", formatOptions).format(new Date(resetTime));
  return `额度已用尽（${label}），将于 ${formatted} 恢复`;
}

module.exports = {
  MAX_USAGE_ITEMS,
  USAGE_TIMEOUT_MS,
  isIsoDateTime,
  quotaExhaustionMessage,
  requestUsage,
  sanitizeUsage
};
