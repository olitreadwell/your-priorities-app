import type { Job } from "./sdk.js";

function part(value: string): string {
  // Keep punctuation and leading zeros; escape separators so distinct components
  // cannot accidentally create the same composite key.
  return value
    .normalize("NFC")
    .trim()
    .replace(/\s+/gu, " ")
    .replaceAll("\\", "\\\\")
    .replaceAll("|", "\\|");
}

/** Business identity across runs; the runner's id still identifies the input row. */
export function uniqueJobId(job: Job): string {
  const code = job.titleCode?.trim() || job.classCode?.trim() || "";
  const title = job.level ? (job.parentName ?? job.name) : job.name;
  const parts = [job.jurisdiction, code, title];
  if (job.level) parts.push(job.level.name);
  // Without a source code/title, do not imply a reliable cross-upload match.
  if (!code || !title.trim()) parts.push(job.id);
  return parts.map(part).join(" | ");
}
