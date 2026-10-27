import assert from "node:assert/strict";
import { existsSync, readFileSync, readdirSync } from "node:fs";
import { test } from "node:test";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { tools } from "../src/tools.ts";

const repo = dirname(dirname(fileURLToPath(import.meta.url)));

/**
 * The documentation an agent acts on without a human checking it: the files that
 * ship, and the skill an installer copies. `docs/plans/` is excluded on purpose —
 * those are records of decisions as they were made, they disagree with the code
 * in places, and nothing that ships cites them (`test/clients.test.ts` holds that
 * line).
 */
const shipped: string[] = [
  join(repo, "README.md"),
  join(repo, "SECURITY.md"),
  join(repo, "docs", "clients.md"),
  join(repo, "skills", "outlook-mcp", "SKILL.md"),
  ...readdirSync(join(repo, "skills", "outlook-mcp", "references")).map(
    (file) => join(repo, "skills", "outlook-mcp", "references", file),
  ),
];

/** A file's lines that are not inside a fenced code block. */
function prose(path: string): string[] {
  const out: string[] = [];
  let fenced = false;
  for (const line of readFileSync(path, "utf8").split(/\r?\n/)) {
    if (line.startsWith("```")) fenced = !fenced;
    else if (!fenced) out.push(line);
  }
  return out;
}

/** GitHub's slug: lowercase, punctuation gone, spaces to hyphens. */
function slug(heading: string): string {
  return heading
    .toLowerCase()
    .replace(/[^\w\s-]/g, "")
    .trim()
    .replace(/\s+/g, "-");
}

test("every link in the shipped documentation points at something that exists", () => {
  for (const path of shipped) {
    const text = prose(path).join("\n");
    for (const [, target] of text.matchAll(/\]\(([^)\s]+)\)/g)) {
      if (/^(https?:|mailto:)/.test(target)) continue;
      const [file, fragment] = target.split("#");
      if (file === undefined || file === "") {
        // An in-page anchor is the link most likely to outlive the heading it
        // named, which is exactly what a reader cannot check without clicking.
        const headings = prose(path)
          .filter((line) => /^#{1,6} /.test(line))
          .map((line) => slug(line.replace(/^#+ /, "")));
        assert.ok(
          headings.includes(fragment ?? ""),
          `${path}: #${fragment} matches no heading in this file`,
        );
        continue;
      }
      const absolute = resolve(dirname(path), file);
      assert.ok(
        existsSync(absolute),
        `${path}: ${file} does not exist (fragment ${fragment ?? "none"})`,
      );
    }
  }
});

test("the documentation pins no image tag, because a pinned tag becomes a lie", () => {
  // Every one of these numbers was true once and then described a server three
  // releases behind the page that named it — a server without the tools the same
  // page listed. `latest` is what a document can say about a moving release;
  // pinning is the reader's job at the point of running, where `config docker`
  // prints the number of the code they have.
  for (const path of [
    ...shipped,
    join(repo, "docs", "plans", "plan-mcp-skill.md"),
  ]) {
    const text = readFileSync(path, "utf8");
    assert.equal(
      /ghcr\.io\/[^\s`)"]*:\d/.test(text),
      false,
      `${path}: pin an image tag nowhere in the docs; say \`latest\``,
    );
    // Releases here carry no assets, so any concrete download URL is a 404.
    assert.equal(
      /releases\/download\/v/.test(text),
      false,
      `${path}: releases publish no tarball asset; describe the shape, do not link one`,
    );
  }
});

test("the README's tool table is the registry's own, in both directions", () => {
  const registered = tools.map((tool) => tool.name).sort();
  const documented: string[] = [];
  for (const line of prose(join(repo, "README.md"))) {
    const row = line.match(/^\|\s*`([a-z_]+)`\s*\|/);
    if (row) documented.push(row[1]);
  }
  assert.deepEqual(
    documented.slice().sort(),
    registered,
    "a tool appeared in src/tools.ts and not in the README table, or the other way round",
  );
});

test("the skill names every tool the registry has, and no tool it has not got", () => {
  const skill = prose(join(repo, "skills", "outlook-mcp", "SKILL.md")).join(
    "\n",
  );
  const registered = new Set(tools.map((tool) => tool.name));
  for (const name of registered) {
    assert.ok(
      skill.includes(`\`${name}\``),
      `the skill never mentions ${name}, so an agent with only the skill cannot reach it`,
    );
  }
  // Inside the skill's tables only: prose elsewhere quotes field names, tool
  // results and Microsoft's own vocabulary, none of which are tools. A table row
  // is the skill asserting a name to be looked up, so every snake_case name in
  // one has to exist.
  const table = skill
    .split("\n")
    .filter((line) => line.startsWith("|"))
    .filter((line) => !/^\|[\s:|-]+$/.test(line))
    .filter((line) => !/^\|\s*Intent\s*\|/.test(line));
  const invented = [...table.join("\n").matchAll(/`([a-z][a-z_]*_[a-z_]*)`/g)]
    .map((match) => match[1])
    .filter((name) => !registered.has(name));
  assert.deepEqual(invented, [], "the skill points at tools that do not exist");
});

test("every environment variable is documented, and none is invented", () => {
  const fromTable = (path: string) =>
    prose(path)
      .filter((line) => line.startsWith("| `OUTLOOK_"))
      .map((line) => line.match(/`(OUTLOOK_[A-Z_]+)`/)![1]);
  const documented = new Set(fromTable(join(repo, "README.md")));

  const read: string[] = [];
  const walk = (dir: string) => {
    for (const item of readdirSync(dir, { withFileTypes: true })) {
      const path = join(dir, item.name);
      if (item.isDirectory()) walk(path);
      else if (path.endsWith(".ts"))
        read.push(
          ...[
            ...readFileSync(path, "utf8").matchAll(
              /\b(OUTLOOK_(?:MCP|LOG|OAUTH)_[A-Z][A-Z_]+)\b/g,
            ),
          ].map((match) => match[1]),
        );
    }
  };
  walk(join(repo, "src"));

  assert.deepEqual(
    [...new Set(read)].sort(),
    [...documented].sort(),
    "src reads an environment variable the README table does not carry, or documents one nothing reads",
  );
});

test("the README's contents list is its own headings, in order", () => {
  const lines = prose(join(repo, "README.md"));
  const headings = lines
    .filter((line) => /^## [^#]/.test(line) && line !== "## Contents")
    .map((line) => line.replace(/^## /, ""));
  const listed = lines
    .slice(lines.indexOf("## Contents"))
    .map((line) => line.match(/^- \[(.+?)\]\(#(.+?)\)$/))
    .filter(Boolean)
    .map((match) => ({ label: match![1], href: match![2] }));
  assert.deepEqual(
    listed.map((entry) => entry.label),
    headings,
    "the contents list and the page's sections disagree",
  );
  for (const entry of listed)
    assert.equal(
      entry.href,
      slug(entry.label),
      `the contents entry [${entry.label}](#${entry.href}) has the wrong anchor`,
    );
});

test("the skill's frontmatter is what an installer reads", () => {
  const path = join(repo, "skills", "outlook-mcp", "SKILL.md");
  const text = readFileSync(path, "utf8");
  const frontmatter = text.match(/^---\n([\s\S]*?)\n---\n/);
  assert.ok(frontmatter, `${path} starts with frontmatter`);
  const fields = Object.fromEntries(
    frontmatter![1]
      .split("\n")
      .filter((line) => line.includes(":"))
      .map((line) => [
        line.slice(0, line.indexOf(":")).trim(),
        line.slice(line.indexOf(":") + 1).trim(),
      ]),
  );
  // The installer names the skill after this, and discovers it by the description.
  assert.equal(fields.name, "outlook-mcp");
  assert.ok(
    fields.description.length > 60,
    "the description carries the trigger",
  );
  // An agent reads the runtime floor before it reads anything else, because the
  // first thing it runs is `node`.
  assert.match(fields.compatibility, /Node 22/);
  assert.match(fields.description, /cannot send|never send|person sends/i);
});
