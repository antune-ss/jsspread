import { defineConfig } from 'vite';
import gas from 'vite-plugin-gas';

export default defineConfig({
  plugins: [
    gas({
      autoDetect: true,
      include: ['src', 'lib'],
      exclude: ['**/*.test.ts', '**/*.spec.ts'],
      outDir: 'dist',

      transformLogger: true,
      copyAppsscriptJson: true,

      enablePathAliases: true,
      autoDetectPathAliases: true,
      pathAliases: {
        '@': './src',
        '@lib': './lib',
        '~': './src'
      }
    })
  ]
});