import obsidianmd from 'eslint-plugin-obsidianmd';
import globals from 'globals';
import tsParser from '@typescript-eslint/parser';
import { globalIgnores } from 'eslint/config';

export default [
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
			globals: {
				...globals.browser,
			},
		},
		plugins: {
			obsidianmd,
		},
		rules: {
			'no-alert': 'error',
			'no-debugger': 'error',
			'no-var': 'error',
			'prefer-const': 'warn',
			'obsidianmd/no-tfile-tfolder-cast': 'error',
			'obsidianmd/regex-lookbehind': 'error',
			'obsidianmd/vault/iterate': 'warn',
		},
	},
];
