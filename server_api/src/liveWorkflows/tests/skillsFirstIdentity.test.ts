import test from "node:test";
import assert from "node:assert/strict";
import ExcelJS from "exceljs";
import { parse } from "csv-parse/sync";
import { loadAgent } from "./skillsFirstFixture.js";
import { renderWorkbook } from "../workbooks.js";
import type { Context } from "../runner/run.js";

const { canonicalJob } = await import(
  new URL(
    "../../../workflow-packages/skills-first/source/memory.ts",
    import.meta.url,
  ).href
);
const { matchedLicenseReferences } = await import(
  new URL(
    "../../../workflow-packages/skills-first/source/licenseComparisons.ts",
    import.meta.url,
  ).href
);
const normalize = await loadAgent("normalize"),
  levels = await loadAgent("levels"),
  exportResults = await loadAgent("export");
const base = {
  id: "upload:1",
  name: "Agriculture Program Manager 8018",
  titleCode: "00180",
  text: "Level I: Experience required. Level II: Advanced experience required.",
  source: "upload",
  jurisdiction: "Idaho, USA",
  status: "complete",
  analysis: {},
  citations: [],
  errors: [],
};
const context = (overrides: Partial<Context> = {}): Context => ({
  nodeId: "fixture",
  mode: "trial",
  configuration: {},
  model: async () => {
    throw new Error("Unexpected model call");
  },
  connector: async () => {
    throw new Error("Unexpected connector");
  },
  artifact: async () => {},
  checkpoint: async () => {},
  restore: async () => undefined,
  progress: async () => {},
  cancelled: async () => false,
  ...overrides,
});

test("business IDs distinguish Idaho titles and states while preserving codes and stable re-import identity", async () => {
  const first = canonicalJob(base);
  const second = canonicalJob({
    ...base,
    id: "upload:2",
    name: "Agriculture Program Manager 8102",
  });
  const nj = canonicalJob({
    ...base,
    id: "upload:3",
    jurisdiction: "New Jersey, USA",
  });
  assert.equal(
    first.uniqueId,
    "Idaho, USA | 00180 | Agriculture Program Manager 8018",
  );
  assert.equal(
    second.uniqueId,
    "Idaho, USA | 00180 | Agriculture Program Manager 8102",
  );
  assert.equal(
    nj.uniqueId,
    "New Jersey, USA | 00180 | Agriculture Program Manager 8018",
  );
  assert.equal(new Set([first.uniqueId, second.uniqueId, nj.uniqueId]).size, 3);
  for (const job of [first, second, nj]) assert.equal(job.titleCode, "00180");
  assert.equal(
    canonicalJob({ ...first, id: "new-upload:7", source: "new-upload" })
      .uniqueId,
    first.uniqueId,
  );
  assert.equal(
    canonicalJob({ ...first, uniqueId: "stale", name: second.name }).uniqueId,
    second.uniqueId,
  );
  assert.equal(
    canonicalJob({ ...base, titleCode: null, classCode: "00180" }).uniqueId,
    first.uniqueId,
  );
  assert.equal(
    canonicalJob({ ...base, titleCode: null, classCode: null }).uniqueId,
    "Idaho, USA |  | Agriculture Program Manager 8018 | upload:1",
  );
  assert.equal(
    canonicalJob({ ...first, name: "Agriculture  Program Manager 8018\n" })
      .uniqueId,
    first.uniqueId,
  );
  const escaped = canonicalJob({ ...base, titleCode: "A | B", name: "C" });
  const other = canonicalJob({ ...base, titleCode: "A", name: "B | C" });
  assert.notEqual(
    escaped.uniqueId,
    other.uniqueId,
    "Component delimiters must not introduce collisions",
  );
  assert.equal(
    canonicalJob(escaped).uniqueId,
    escaped.uniqueId,
    "Escaping must not accumulate on re-export",
  );
  assert.notEqual(
    canonicalJob({ ...base, name: "A-B" }).uniqueId,
    canonicalJob({ ...base, name: "A B" }).uniqueId,
  );
  const rows = (await normalize(context(), [
    {
      name: first.name,
      titleCode: "00180",
      text: base.text,
      jurisdiction: base.jurisdiction,
    },
    {
      name: second.name,
      titleCode: "00180",
      text: base.text,
      jurisdiction: base.jurisdiction,
    },
  ])) as unknown as Record<string, unknown>[];
  assert.deepEqual(
    rows.map((job) => job.id),
    ["1", "2"],
    "Repeated codes must not become duplicate runner IDs",
  );
  assert.deepEqual(
    rows.map((job) => job.uniqueId),
    [first.uniqueId, second.uniqueId],
  );
});

test("split levels include their parent title so matching level names under a repeated code stay distinct", async () => {
  const parents = [
    canonicalJob(base),
    canonicalJob({
      ...base,
      id: "upload:2",
      name: "Agriculture Program Manager 8102",
    }),
  ];
  const output = (await levels(
    context({
      model: async (key): Promise<YpWorkflowJson> =>
        key.endsWith(":review")
          ? {
              allLevelsCovered: true,
              sharedRequirementsPreserved: true,
              noCrossLevelRequirements: true,
              issues: [],
            }
          : {
              isMultiLevel: true,
              sharedQuotes: [],
              levels: [
                {
                  name: "Level I",
                  evidenceQuotes: ["Level I: Experience required."],
                },
                {
                  name: "Level II",
                  evidenceQuotes: ["Level II: Advanced experience required."],
                },
              ],
            },
    }),
    parents,
  )) as unknown as Record<string, any>[];
  assert.equal(output.length, 4);
  assert.equal(new Set(output.map((job) => job.uniqueId)).size, 4);
  assert.deepEqual(
    output.map((job) => job.uniqueId),
    parents.flatMap((parent) => [
      parent.uniqueId + " | Level I",
      parent.uniqueId + " | Level II",
    ]),
  );
  for (const job of output) {
    assert.equal(job.titleCode, "00180");
    assert.equal(canonicalJob(job).uniqueId, job.uniqueId);
    assert.ok(parents.some((parent) => parent.id === job.parentId));
  }
});

test("Unique ID leads every job sheet and binds repeated-code research and license rows to their owning job", async () => {
  const jobs = [
    canonicalJob(base),
    canonicalJob({
      ...base,
      id: "upload:2",
      name: "Agriculture Program Manager 8102",
    }),
  ];
  for (const job of jobs) {
    job.degreeAnalysis.deepResearchResults = ["one", "two"].map((suffix) => ({
      uniqueId: "wrong inherited key",
      titleCode: job.titleCode,
      jobTitle: "Partial researched title",
      statedDegreeRequirement: null,
      degreeRequirementType: null,
      typeOfOfficialDocument: "classification",
      matchTypeForJobTitle: "partial",
      typeOfDegreeRequirement: null,
      mandatoryDegreeRequirementStatus: null,
      reasoning: suffix,
      sourceUrl: `https://example.gov/${suffix}`,
      evidenceQuote: "Exact supplied evidence",
      status: "unknown",
    }));
    job.licenseComparisons = [
      {
        uniqueId: "wrong inherited key",
        licenseType: "Surveyor",
        manualMinimumEducationRequirement: "Supplied",
      },
    ];
  }
  const artifacts = new Map<string, string>();
  await exportResults(
    context({
      artifact: async (name, content) => {
        artifacts.set(name, content);
      },
    }),
    jobs,
  );
  const memory = JSON.parse(artifacts.get("skills-first.json")!);
  const expectedIds = jobs.map((job) => job.uniqueId);
  const researchIds = expectedIds.flatMap((id) => [id, id]);
  assert.deepEqual(
    memory.jobDescriptions.map((job: any) => job.uniqueId),
    expectedIds,
  );
  assert.deepEqual(
    memory.educationRequirementResearchResults.map((row: any) => row.uniqueId),
    researchIds,
  );
  assert.deepEqual(
    memory.licenseComparisons.map((row: any) => row.uniqueId),
    expectedIds,
  );
  const csv = parse(artifacts.get("skills-first.csv")!) as string[][];
  assert.equal(csv[1].length, 75);
  assert.deepEqual(csv[1].slice(0, 3), ["Unique ID", "agentId", "titleCode"]);
  assert.deepEqual(
    csv.slice(2).map((row) => row[0]),
    researchIds,
  );
  assert.ok(csv.slice(2).every((row) => row[2] === "00180"));
  const book = new ExcelJS.Workbook();
  const bytes = await renderWorkbook(artifacts.get("skills-first.xlsx")!);
  await book.xlsx.load(
    bytes as unknown as Parameters<typeof book.xlsx.load>[0],
  );
  for (const sheet of book.worksheets) {
    const headerRow = ["merged (master + legal)", "Master sheet old"].includes(
      sheet.name,
    )
      ? 2
      : 1;
    assert.equal(sheet.getCell(headerRow, 1).value, "Unique ID", sheet.name);
    const expected = ["merged (master + legal)", "legal research"].includes(
      sheet.name,
    )
      ? researchIds
      : expectedIds;
    assert.deepEqual(
      expected.map((_, i) => sheet.getCell(headerRow + 1 + i, 1).value),
      expected,
      sheet.name,
    );
    assert.equal(sheet.rowCount, expected.length + headerRow, sheet.name);
  }
});

test("job-specific license references use Unique ID; a repeated title code cannot pull another title or state's reference", () => {
  const job = canonicalJob(base);
  job.degreeAnalysis.professionalLicenseRequirement.licenseType = "Surveyor";
  const refs = [
    { uniqueId: job.uniqueId, licenseType: "Surveyor", notes: "this job" },
    {
      uniqueId: "Idaho, USA | 00180 | Agriculture Program Manager 8102",
      licenseType: "Surveyor",
      notes: "another title",
    },
    {
      "Unique ID": "New Jersey, USA | 00180 | Agriculture Program Manager 8018",
      licenseType: "Surveyor",
      notes: "another state",
    },
    { licenseType: "Surveyor", notes: "generic license reference" },
  ];
  const matched = matchedLicenseReferences(job, refs);
  assert.deepEqual(
    matched.map((ref: any) => ref.notes),
    ["this job", "generic license reference"],
  );
  assert.ok(matched.every((ref: any) => ref.uniqueId === job.uniqueId));
});
