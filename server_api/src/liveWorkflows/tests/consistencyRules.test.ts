import test from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import { parse } from "csv-parse/sync";
import { analysisResponse } from "./analysisFixture.js";
import { loadAgent } from "./skillsFirstFixture.js";
import { verifySkillsFirstValues } from "./manual/skillsFirstValues.js";
import { renderWorkbook } from "../workbooks.js";
import { stableRequestId } from "../runner/socketClient.js";
import type { Context } from "../runner/run.js";
import type { UsageReservation } from "../budget.js";

const { canonicalJob, consistencyKeys } = await import(
  new URL(
    "../../../workflow-packages/skills-first/source/memory.ts",
    import.meta.url,
  ).href
);
const analyze = await loadAgent("analysis");
const exportResults = await loadAgent("export");
const check = "alternativeQualificationsConsistency";
const description =
  "Applicants must qualify in at least one specialty. Managing commercial seed operations OR providing consultation OR conducting inspections. Obtain AOSA certification OR SCST certification within two years.";
const base = {
  id: "agriculture-8018",
  titleCode: "180",
  name: "Agriculture Program Manager 8018",
  text: description,
  originalText: description,
  source: "fixture-upload",
  jurisdiction: "Idaho, USA",
  status: "complete" as const,
  analysis: {},
  citations: [],
  errors: [],
};
const context = (overrides: Partial<Context> = {}): Context => ({
  nodeId: "analyze",
  mode: "trial",
  configuration: {},
  model: async () => {
    throw new Error("Unexpected model call");
  },
  connector: async () => {
    throw new Error("Unexpected connector call");
  },
  progress: async () => {},
  checkpoint: async () => {},
  restore: async () => undefined,
  artifact: async () => {},
  cancelled: async () => false,
  ...overrides,
});
type Verdict = "pass" | "fail" | "n/a" | "unknown";
type AnalyzedJob = YpSkillsFirstJobDescription & {
  analysis: Record<string, YpWorkflowJson>;
  errors: string[];
  status: "pending" | "complete" | "unknown";
};
function checks(verdict: Verdict) {
  return Object.fromEntries(
    consistencyKeys.map((key: string) => [
      key,
      key === check
        ? verdict
        : key === "cscRevisedConsistency"
          ? "n/a"
          : "pass",
    ]),
  );
}
function mandatory(alternatives: string[] | null) {
  return {
    isDegreeMandatory: null,
    hasAlternativeQualifications: false,
    alternativeQualifications: alternatives,
    multipleQualificationPaths: true,
    isDegreeAbsolutelyRequired: null,
    substitutionPossible: null,
    degreePreferenceQuote: null,
    mandatoryStatusExplanations: {
      degreeRequirementExplanation:
        "No degree is stated. Specialty experience and certification options are not stated degree substitutions.",
      bothTrueExplanation: null,
      bothFalseExplanation: null,
    },
  };
}
function savedJob(alternatives: string[] | null, verdict: Verdict) {
  return {
    ...base,
    analysis: {
      mandatory: {
        status: "known",
        findings: mandatory(alternatives),
        evidence: [description],
      },
      consistency: {
        status: "known",
        findings: { validationChecks: checks(verdict) },
        evidence: [description],
      },
    },
  };
}

test("contradictory qualification paths fail for both empty and null alternatives, regardless of model verdict", () => {
  for (const alternatives of [[], null]) {
    for (const verdict of ["pass", "fail", "n/a", "unknown"] as const) {
      const input = savedJob(alternatives, verdict);
      const before = structuredClone(input);
      const job = canonicalJob(input);
      assert.equal(job.degreeAnalysis.validationChecks[check], "fail");
      assert.equal(job.validationScore, 7);
      assert.equal(job.status, "unknown");
      assert.equal(
        job.analysis.consistency.findings.validationChecks[check],
        verdict,
      );
      assert.deepEqual(
        job.degreeAnalysis.degreeRequirementStatus.alternativeQualifications,
        alternatives,
      );
      assert.equal(
        job.degreeAnalysis.degreeRequirementStatus.hasAlternativeQualifications,
        false,
      );
      assert.equal(
        job.degreeAnalysis.degreeRequirementStatus.multipleQualificationPaths,
        true,
      );
      assert.equal(job.analysis.consistencyRules.originalResult, verdict);
      assert.equal(job.analysis.consistencyRules.enforcedResult, "fail");
      assert.match(
        job.analysis.consistencyRules.issues[0],
        /multipleQualificationPaths is true but hasAlternativeQualifications is false/,
      );
      assert.deepEqual(
        canonicalJob(job),
        job,
        "Repeated projection retains the correction and original verdict",
      );
      assert.deepEqual(
        input,
        before,
        "Saved input and original model findings remain immutable",
      );
    }
  }
  const legacy = canonicalJob({
    ...base,
    degreeAnalysis: {
      degreeRequirementStatus: mandatory(null),
      validationChecks: checks("pass"),
    },
  });
  assert.equal(legacy.degreeAnalysis.validationChecks[check], "fail");
  assert.equal(legacy.analysis.consistencyRules.originalResult, "pass");
  assert.deepEqual(
    canonicalJob(legacy),
    legacy,
    "Nested-only legacy results keep the original verdict on re-export",
  );
});

test("consistent and unknown qualification flags are not coerced into failures", () => {
  for (const [multiplePaths, hasAlternatives] of [
    [true, true],
    [false, false],
    [true, null],
    [null, false],
    [null, null],
  ]) {
    const input = savedJob(null, "pass");
    Object.assign(input.analysis.mandatory.findings, {
      multipleQualificationPaths: multiplePaths,
      hasAlternativeQualifications: hasAlternatives,
      alternativeQualifications:
        hasAlternatives === true ? ["Explicit alternative"] : null,
    });
    const job = canonicalJob(input);
    assert.equal(job.degreeAnalysis.validationChecks[check], "pass");
    assert.equal(job.validationScore, 8);
    assert.equal(job.analysis.consistencyRules, undefined);
    assert.equal(job.status, "complete");
    assert.equal(
      job.degreeAnalysis.degreeRequirementStatus.multipleQualificationPaths,
      multiplePaths,
    );
    assert.equal(
      job.degreeAnalysis.degreeRequirementStatus.hasAlternativeQualifications,
      hasAlternatives,
    );
  }
});

test("a targeted mandatory analysis does not manufacture a consistency review", async () => {
  let calls = 0;
  const result = (await analyze(
    context({
      configuration: { tasks: ["mandatory"] },
      model: async (key) => {
        calls++;
        assert.ok(key.endsWith(":mandatory"));
        return {
          status: "known",
          findings: mandatory(null),
          evidence: [description],
        };
      },
    }),
    [base],
  )) as unknown as AnalyzedJob[];
  const job = result[0];
  assert.equal(calls, 1);
  assert.deepEqual(
    Object.values(job.degreeAnalysis.validationChecks),
    Array(9).fill(null),
  );
  assert.equal(job.validationScore, null);
  assert.equal(job.analysis.consistencyRules, undefined);
});

test("both Idaho-style rows keep the enforced failure through analysis, exports and independent value verification", async () => {
  const runId = "qualification-consistency";
  const usage: UsageReservation[] = [];
  const output = await analyze(
    context({
      model: async (key, _prompt, input) => {
        const task = key.split(":").at(-1)!;
        const first = key.startsWith(base.id + ":");
        const reply = analysisResponse(task, description) as {
          findings: Record<string, YpWorkflowJson>;
        } & Record<string, YpWorkflowJson>;
        // Exercise the current source-passage binding, not just legacy quoted replies.
        const passages = (input as Record<string, YpWorkflowJson>)
          .sourcePassages as { id: string }[];
        delete reply.evidence;
        reply.evidenceIds = passages.map((p) => p.id);
        if (task === "mandatory") {
          reply.findings = mandatory(first ? [] : null);
          delete reply.findings.degreePreferenceQuote;
          reply.findings.degreePreferenceEvidenceId = null;
        }
        if (task === "consistency")
          reply.findings = {
            validationChecks: checks(first ? "fail" : "pass"),
          };
        if (task === "degrees" || task === "licenses") {
          reply.findings.documentStatement = "not_stated";
          reply.findings.documentStatementExplanation =
            "The supplied description states no degree or professional license requirement.";
        }
        if (task === "readingLevel") {
          delete reply.findings.difficultPassages;
          reply.findings.difficultPassageIds = null;
        }
        usage.push({
          id: stableRequestId(`${runId}:analyze:${key}`),
          runId,
          nodeId: "analyze",
          fingerprint: "fixture",
          status: "settled",
          maximumMicros: 1,
          actualMicros: 1,
          result: { content: JSON.stringify(reply) },
          description: task,
          createdAt: new Date(0).toISOString(),
        });
        return reply;
      },
    }),
    [
      base,
      {
        ...base,
        id: "agriculture-8102",
        name: "Agriculture Program Manager 8102",
      },
    ],
  );
  assert.equal(
    usage.length,
    14,
    "No extra model calls are needed for the correction",
  );
  for (const job of output as unknown as AnalyzedJob[]) {
    assert.deepEqual(job.errors, []);
    assert.equal(job.degreeAnalysis.validationChecks[check], "fail");
    assert.equal(job.validationScore, 7);
    assert.equal(job.status, "unknown");
  }
  const artifacts = new Map<string, string>();
  const records = await exportResults(
    context({
      mode: "bulk",
      configuration: { outputBindings: ["must-not-publish-unresolved-jobs"] },
      artifact: async (name, content) => {
        artifacts.set(name, content);
      },
    }),
    output,
  );
  const memory = JSON.parse(artifacts.get("skills-first.json")!);
  const verificationInput = {
    runId,
    records: records as unknown as Record<string, unknown>[],
    usage,
    memory,
    csv: Buffer.from(artifacts.get("skills-first.csv")!),
    xlsx: await renderWorkbook(artifacts.get("skills-first.xlsx")!),
  };
  const verified = await verifySkillsFirstValues(verificationInput);
  assert.equal(
    verified.passed,
    true,
    JSON.stringify(verified.missingOrChangedValues),
  );
  assert.equal(verified.rawResponses, 14);
  assert.match(
    artifacts.get("skills-first.md")!,
    /7 passed · 1 failed · 1 not applicable/,
  );
  const csv = parse(verificationInput.csv) as string[][];
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(
    verificationInput.xlsx as unknown as Parameters<typeof book.xlsx.load>[0],
  );
  for (const [index, job] of memory.jobDescriptions.entries()) {
    assert.equal(job.degreeAnalysis.validationChecks[check], "fail");
    assert.equal(job.validationScore, 7);
    assert.equal(csv[index + 2][45], "fail");
    for (const sheet of ["Master sheet old", "merged (master + legal)"]) {
      assert.equal(
        book.getWorksheet(sheet)!.getCell(`AT${index + 3}`).value,
        "fail",
      );
      assert.equal(
        book.getWorksheet(sheet)!.getCell(`AZ${index + 3}`).value,
        7,
      );
    }
    assert.equal(
      book.getWorksheet("Validation summary")!.getCell(`E${index + 2}`).value,
      "7 passed · 1 failed · 1 not applicable",
    );
    const audit = JSON.parse(
      String(
        book.getWorksheet("Review and provenance")!.getCell(`J${index + 2}`)
          .value,
      ),
    );
    assert.equal(
      audit.consistency.findings.validationChecks[check],
      index === 0 ? "fail" : "pass",
    );
    assert.equal(
      audit.consistencyRules.originalResult,
      index === 0 ? "fail" : "pass",
    );
    assert.equal(audit.consistencyRules.enforcedResult, "fail");
  }
  // The acceptance test must catch a regression instead of exempting the field.
  for (const damage of ["verdict", "audit"]) {
    const damaged = structuredClone(memory);
    if (damage === "verdict")
      damaged.jobDescriptions[1].degreeAnalysis.validationChecks[check] =
        "pass";
    else delete damaged.jobDescriptions[1].analysis.consistencyRules;
    const caught = await verifySkillsFirstValues({
      ...verificationInput,
      memory: damaged,
      records: damaged.jobDescriptions,
    });
    assert.equal(caught.passed, false);
    assert.ok(
      caught.missingOrChangedValues.some(
        (issue) =>
          issue.field ===
          (damage === "verdict"
            ? "consistency.validationChecks→degreeAnalysis.validationChecks"
            : "consistencyRules.rule"),
      ),
    );
  }
});
