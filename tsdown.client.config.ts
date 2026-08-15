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
      return absolute === null ? null : `\0css:${absolute}.mjs`
    },
    async load(id: string) {
      if (!id.startsWith('\0css:')) return null
      const file = id.slice(5).replace(/\.mjs$/, '')
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
