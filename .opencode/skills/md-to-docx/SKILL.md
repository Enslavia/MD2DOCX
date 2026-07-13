# MD ↔ DOCX Auto-Conversion

When you write a `.md` or `.docx` file, the opencode plugin
(`md2docx-convert.ts`) automatically produces the paired format.

## Behaviour

- Writing a `.md` file produces a `.docx` alongside it.
- Writing a `.docx` file produces a `.md` alongside it.
- If the output file already exists, the plugin **skips** the conversion
  and logs a message.

## How to Use

**Before overwriting an existing output file, ask the user for confirmation.**

Example:
> The output file `document.docx` already exists. Overwrite it?
> Reply with `yes` to proceed.

If the user confirms, delete the existing output file before writing the
source, or run the conversion manually:

```
MD2DOCX.app/cli document.md
```

## CLI Binary Location

The plugin searches for the CLI binary in this order:

1. `$MD2DOCX_CLI` environment variable
2. `/Applications/MD2DOCX.app/cli`
3. `md2docx` on PATH

## Manual Conversion

You can also convert files manually:

```bash
MD2DOCX.app/cli input.md      # → input.docx
MD2DOCX.app/cli input.docx    # → input.md
```
