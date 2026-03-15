#!/usr/bin/env node
import fs from "node:fs";
import path from "node:path";

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

const parseArgs = (argv) => {
  const out = {
    report: "yandex-form-runs-report-69a84218-764.json",
    quotaTargets: "quota-targets-764-photo.json",
    requireComplete: false,
  };
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--report" && argv[i + 1]) {
      out.report = argv[i + 1];
      i += 1;
      continue;
    }
    if (arg === "--quota-targets" && argv[i + 1]) {
      out.quotaTargets = argv[i + 1];
      i += 1;
      continue;
    }
    if (arg === "--require-complete") {
      out.requireComplete = true;
      continue;
    }
  }
  return out;
};

const readJson = (filePath) => {
  const absPath = path.resolve(process.cwd(), filePath);
  const raw = fs.readFileSync(absPath, "utf8");
  return JSON.parse(raw);
};

const toNumber = (value) => {
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
};

const collectSuccessfulRuns = (report) => {
  const fromSuccessfulRuns = Array.isArray(report?.successful_runs)
    ? report.successful_runs
    : [];
  const fromAttempts = Array.isArray(report?.attempts)
    ? report.attempts.filter((item) => String(item?.status) === "success")
    : [];
  const out = [];
  const seen = new Set();
  for (const item of [...fromSuccessfulRuns, ...fromAttempts]) {
    const id = String(item?.attempt_id ?? item?.run_id ?? `${out.length + 1}`);
    if (seen.has(id)) continue;
    seen.add(id);
    out.push(item);
  }
  return out;
};

const auditRandomSources = (runs) => {
  const allowedRandomSources = new Set(["income_random"]);
  const violations = [];
  const seenRandomSources = new Map();

  for (const run of runs) {
    const attemptId = run?.attempt_id ?? run?.run_id ?? null;
    const answers = Array.isArray(run?.answers_summary) ? run.answers_summary : [];
    for (const answer of answers) {
      const source = String(answer?.source ?? "");
      if (!source.includes("random")) continue;
      seenRandomSources.set(source, (seenRandomSources.get(source) ?? 0) + 1);
      if (allowedRandomSources.has(source)) continue;
      violations.push({
        attempt_id: attemptId,
        source,
        page: answer?.page ?? null,
        type: answer?.type ?? null,
        value: answer?.value ?? null,
      });
    }
  }

  return {
    ok: violations.length === 0,
    seen_random_sources: Object.fromEntries(seenRandomSources),
    violations,
  };
};

const auditLikertNotStuckToOne = (runs) => {
  const values = [];
  for (const run of runs) {
    const answers = Array.isArray(run?.answers_summary) ? run.answers_summary : [];
    for (const answer of answers) {
      if (answer?.type !== "likert_table_row") continue;
      const score = toNumber(answer?.value);
      if (score === null) continue;
      values.push(score);
    }
  }

  const uniqueValues = [...new Set(values)];
  const stuckToOne = values.length >= 20 && uniqueValues.length === 1 && uniqueValues[0] === 1;
  return {
    ok: !stuckToOne,
    total_likert_rows: values.length,
    unique_values: uniqueValues.sort((a, b) => a - b),
  };
};

const auditQuotaConsistency = (summary, targetsRoot) => {
  const targets = targetsRoot?.targets && typeof targetsRoot.targets === "object"
    ? targetsRoot.targets
    : targetsRoot;
  const completed = summary?.quota_completed && typeof summary.quota_completed === "object"
    ? summary.quota_completed
    : {};
  const remaining = summary?.quota_remaining && typeof summary.quota_remaining === "object"
    ? summary.quota_remaining
    : {};

  const violations = [];
  let targetTotal = 0;
  let completedTotal = 0;
  for (const [cellKey, targetRaw] of Object.entries(targets ?? {})) {
    const target = toNumber(targetRaw) ?? 0;
    const done = toNumber(completed[cellKey]) ?? 0;
    const rem = toNumber(remaining[cellKey]);
    targetTotal += target;
    completedTotal += done;
    if (done > target) {
      violations.push({
        type: "completed_exceeds_target",
        cell: cellKey,
        target,
        completed: done,
      });
    }
    if (rem !== null) {
      const expectedRemaining = target - done;
      if (rem !== expectedRemaining) {
        violations.push({
          type: "remaining_mismatch",
          cell: cellKey,
          target,
          completed: done,
          remaining: rem,
          expected_remaining: expectedRemaining,
        });
      }
    }
  }

  const summaryRemainingTotal = toNumber(summary?.quota_remaining_total);
  if (summaryRemainingTotal !== null) {
    const expectedRemainingTotal = targetTotal - completedTotal;
    if (summaryRemainingTotal !== expectedRemainingTotal) {
      violations.push({
        type: "remaining_total_mismatch",
        summary_remaining_total: summaryRemainingTotal,
        expected_remaining_total: expectedRemainingTotal,
      });
    }
  }

  return {
    ok: violations.length === 0,
    target_total_from_targets: targetTotal,
    completed_total_from_summary: completedTotal,
    violations,
  };
};

const auditAgeRanges = (runs) => {
  const violations = [];
  let checked = 0;

  for (const run of runs) {
    const dims = run?.quota_dimensions ?? null;
    const tripType = dims?.trip_type;
    const childrenMode = dims?.children_mode;
    const range = AGE_RULES?.[tripType]?.[childrenMode];
    if (!range) continue;
    const answers = Array.isArray(run?.answers_summary) ? run.answers_summary : [];
    const ageAnswer = answers.find(
      (entry) =>
        entry?.type === "number" && String(entry?.source ?? "").includes("age"),
    );
    if (!ageAnswer) continue;
    const age = toNumber(ageAnswer?.value);
    if (age === null) continue;
    checked += 1;
    if (age < range[0] || age > range[1]) {
      violations.push({
        attempt_id: run?.attempt_id ?? run?.run_id ?? null,
        age,
        range,
        trip_type: tripType,
        children_mode: childrenMode,
      });
    }
  }

  return {
    ok: violations.length === 0,
    checked_attempts: checked,
    violations,
  };
};

const auditCompletion = (summary, requireComplete) => {
  if (!requireComplete) {
    return {
      ok: true,
      required: false,
      run_exit_reason: summary?.run_exit_reason ?? null,
      quota_remaining_total: summary?.quota_remaining_total ?? null,
    };
  }

  const quotaRemainingTotal = toNumber(summary?.quota_remaining_total);
  const ok =
    quotaRemainingTotal === 0 && String(summary?.run_exit_reason ?? "") === "quota_reached";
  return {
    ok,
    required: true,
    run_exit_reason: summary?.run_exit_reason ?? null,
    quota_remaining_total: summary?.quota_remaining_total ?? null,
  };
};

const auditTraceCoverage = (summary, successfulRunsAnalyzed) => {
  const successfulSubmits = toNumber(summary?.successful_submits) ?? 0;
  const ok = !(successfulSubmits > 0 && successfulRunsAnalyzed === 0);
  return {
    ok,
    successful_submits_summary: successfulSubmits,
    successful_runs_analyzed: successfulRunsAnalyzed,
  };
};

const main = () => {
  const args = parseArgs(process.argv.slice(2));
  const report = readJson(args.report);
  const targets = readJson(args.quotaTargets);
  const runs = collectSuccessfulRuns(report);

  const randomSourcesAudit = auditRandomSources(runs);
  const likertAudit = auditLikertNotStuckToOne(runs);
  const quotaAudit = auditQuotaConsistency(report?.summary ?? {}, targets);
  const ageAudit = auditAgeRanges(runs);
  const completionAudit = auditCompletion(report?.summary ?? {}, args.requireComplete);
  const traceCoverageAudit = auditTraceCoverage(report?.summary ?? {}, runs.length);

  const checks = {
    random_sources: randomSourcesAudit,
    likert_not_stuck_to_one: likertAudit,
    quota_consistency: quotaAudit,
    age_ranges: ageAudit,
    completion: completionAudit,
    trace_coverage: traceCoverageAudit,
  };
  const failedChecks = Object.entries(checks)
    .filter(([, item]) => item?.ok === false)
    .map(([name]) => name);
  const ok = failedChecks.length === 0;

  const output = {
    ok,
    report_path: path.resolve(process.cwd(), args.report),
    quota_targets_path: path.resolve(process.cwd(), args.quotaTargets),
    successful_runs_analyzed: runs.length,
    failed_checks: failedChecks,
    checks,
  };
  console.log(JSON.stringify(output, null, 2));
  process.exit(ok ? 0 : 1);
};

main();
