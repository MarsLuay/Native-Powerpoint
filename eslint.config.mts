import sdl from '@microsoft/eslint-plugin-sdl';
import obsidianmd from 'eslint-plugin-obsidianmd';
import noUnsanitizedPlugin from 'eslint-plugin-no-unsanitized';
import globals from 'globals';
import tsParser from '@typescript-eslint/parser';
import { defineConfig, globalIgnores } from 'eslint/config';

export default defineConfig([
	globalIgnores([
		'node_modules',
		'dist',
		'main.js',
		'package-lock.json',
	]),
	{
		files: ['src/**/*.ts', 'src/**/*.tsx'],
		languageOptions: {
			parser: tsParser,
			parserOptions: {
				project: './tsconfig.json',
				tsconfigRootDir: import.meta.dirname,
			},
			globals: {
				...globals.browser,
			},
		},
		plugins: {
			'@microsoft/sdl': sdl,
			'no-unsanitized': noUnsanitizedPlugin,
			obsidianmd,
		},
		rules: {
			'@microsoft/sdl/no-inner-html': 'error',
			'no-unsanitized/method': 'error',
			'no-unsanitized/property': 'error',
			'no-alert': 'error',
			'no-debugger': 'error',
			'no-var': 'error',
			'prefer-const': 'warn',
			'obsidianmd/no-static-styles-assignment': 'error',
			'obsidianmd/no-tfile-tfolder-cast': 'error',
			'obsidianmd/no-unsupported-api': 'error',
			'obsidianmd/regex-lookbehind': 'error',
			'obsidianmd/vault/iterate': 'warn',
		},
	},
]);
