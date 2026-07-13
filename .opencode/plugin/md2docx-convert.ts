import type { Plugin } from "opencode";

const CLI_BINARY_CANDIDATES = [
  process.env.MD2DOCX_CLI,
  "/Applications/MD2DOCX.app/cli",
  "/usr/local/bin/md2docx",
];

function findCli(): string | null {
  for (const candidate of CLI_BINARY_CANDIDATES) {
    if (candidate) {
      try {
        const result = Bun.sync`which ${candidate}`;
        if (result.exitCode === 0) return candidate;
      } catch {}
      try {
        const stat = Bun.sync`stat ${candidate}`;
        if (stat.exitCode === 0) return candidate;
      } catch {}
    }
  }
  return "md2docx";
}

const plugin: Plugin = {
  name: "md2docx-convert",
  hooks: {
    "tool:execute:after": async ({ tool, args, result }) => {
      if (tool !== "write" || result.type !== "success") return;

      const path = args.filePath;
      if (typeof path !== "string") return;

      const ext = path.toLowerCase().match(/\.(\w+)$/)?.[1];
      if (!ext || !["md", "markdown", "docx"].includes(ext)) return;

      const isMd = ext === "md" || ext === "markdown";
      const isDocx = ext === "docx";
      if (!isMd && !isDocx) return;

      const outExt = isMd ? "docx" : "md";
      const outPath = path.replace(/\.\w+$/, `.${outExt}`);

      try {
        const stat = Bun.sync`stat ${outPath}`;
        if (stat.exitCode === 0) {
          console.log(`[md2docx] output exists, skipping: ${outPath}`);
          return;
        }
      } catch {}

      const cli = findCli();
      if (!cli) {
        console.log("[md2docx] no CLI binary found, skipping conversion");
        return;
      }

      try {
        const conv = Bun.spawnSync([cli, path], { timeout: 30000 });
        if (conv.exitCode === 0) {
          console.log(`[md2docx] converted: ${path} → ${outPath}`);
        } else {
          console.log(`[md2docx] conversion failed: ${conv.stderr.toString()}`);
        }
      } catch (e) {
        console.log(`[md2docx] conversion error: ${e}`);
      }
    },
  },
};

export default plugin;
