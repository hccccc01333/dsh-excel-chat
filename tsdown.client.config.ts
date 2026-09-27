import { createHash } from 'node:crypto'
import { existsSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { dirname, resolve as resolvePath } from 'node:path'

/** Module-table specifiers the web shell shares (platform seed + runtime store exemption). */
const EXTERNALS = [
  'react',
  'react/jsx-runtime',
  'react-dom',
  'react-dom/client',
  '@deepseek-ai/cordis',
  '@deepseek-ai/dsh-client-ui-slots',
  '@deepseek-ai/dsh-client-web-react',
  '@deepseek-ai/dsh-client-ui-primitives',
  '@deepseek-ai/dsh-client-ui-attachment',
  '@deepseek-ai/dsh-client-schema-form',
  '@deepseek-ai/dsh-client-runtime/client',
]

/** Maps a stylesheet's hashed virtual id back to its real path on disk. */
const cssPathById = new Map<string, string>()

/**
 * Browser client bundle for dsh-excel-chat: emits a closure-factory artifact
 * that registers with window.__ModuleLoader__.load so the harness web shell
 * mounts our toolviews into the right details column.
 */
export default {
  name: 'dsh-excel-chat/client',
  entry: { client: 'src/client/index.tsx' },
  outDir: 'bundle/dist-client',
  format: 'cjs',
  platform: 'browser',
  dts: false,
  sourcemap: true,
  clean: true,
  external: EXTERNALS,
  noExternal: (id: string) => (EXTERNALS.includes(id) ? undefined : true),
  define: {
    'process.env.NODE_ENV': JSON.stringify('production'),
    'import.meta.env.MODE': JSON.stringify('production'),
    'import.meta.env': JSON.stringify({ MODE: 'production' }),
  },
  outputOptions: {
    entryFileNames: 'client.js',
    banner: 'window.__ModuleLoader__.load({ id: "dsh-excel-chat", factory: (require) => {',
    footer: 'return module.exports; } });',
    intro: 'var module = { exports: {} }; var exports = module.exports;',
  },
  plugins: [{
    name: 'css-text-inline',
    resolveId(source: string, importer: string | undefined) {
      if (!/\.css$/.test(source) || importer === undefined) return null
      const absolute = resolveBare(source, importer)
      if (absolute === null) return null
      // The virtual id must NOT embed the absolute path. Rolldown writes the
      // module id into a `//#region` comment, which leaked the build machine's
      // local path (e.g. D:\vera\node_modules\...) into the shipped bundle.
      // Hash it instead, and keep the real path in cssPathById for `load`.
      const id = `\0css:${createHash('sha1').update(absolute).digest('hex').slice(0, 12)}.mjs`
      cssPathById.set(id, absolute)
      return id
    },
    async load(id: string) {
      const file = cssPathById.get(id)
      if (!file) return null
      const text = await readFile(file, 'utf8')
      return `export default ${JSON.stringify(text)}`
    },
  }],
}

function resolveBare(source: string, importer: string): string | null {
  const start = source.startsWith('.') || source.startsWith('/')
    ? resolvePath(dirname(importer), source)
    : null
  if (start !== null && existsSync(start)) return start
  if (!source.startsWith('.')) {
    let dir = dirname(importer)
    for (;;) {
      const candidate = resolvePath(dir, 'node_modules', source)
      if (existsSync(candidate)) return candidate
      const parent = dirname(dir)
      if (parent === dir) break
      dir = parent
    }
  }
  return null
}
